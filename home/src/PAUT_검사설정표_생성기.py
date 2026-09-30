# -*- coding: utf-8 -*-
"""독립 실행형 PAUT 검사설정표 생성기."""

from __future__ import annotations

import argparse
import json
import math
import os
import tempfile
import tkinter as tk
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import openpyxl
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


APP_TITLE = "PAUT 검사설정표 생성기"
APP_VERSION = "1.0.0"
CONFIG_PATH = Path(__file__).with_name("paut_scan_plan_settings.json")
PROBE_SOURCE = "https://ims.evidentscientific.com/en/products/u8330072/u8330072"


@dataclass
class ScanSettings:
    probe_model: str = "5L64-A2"
    frequency_mhz: float = 5.0
    total_elements: int = 64
    pitch_mm: float = 0.60
    elevation_mm: float = 10.0
    active_elements: int = 32
    first_element: int = 1
    thickness_mm: float = 15.87
    bevel_angle_deg: float = 37.5
    bevel_tolerance_deg: float = 2.5
    min_angle_deg: float = 46.0
    max_angle_deg: float = 61.0
    angle_step_deg: float = 1.0
    wave_type: str = "SW"
    law_config: str = "Sectorial"
    focus_type: str = "True depth"
    focal_depth_mm: float = 15.87
    scan_direction: str = "양측 주사"
    index_offsets: str = "15, 30, 40"

    def index_values(self) -> list[float]:
        values = []
        for token in self.index_offsets.replace("/", ",").split(","):
            token = token.strip()
            if token:
                values.append(float(token))
        return values

    def validate(self) -> None:
        if not self.probe_model.strip():
            raise ValueError("탐촉자 모델을 입력하세요.")
        positive = {
            "주파수": self.frequency_mhz,
            "총 소자 수": self.total_elements,
            "Pitch": self.pitch_mm,
            "Elevation": self.elevation_mm,
            "활성 소자 수": self.active_elements,
            "두께": self.thickness_mm,
            "각도 간격": self.angle_step_deg,
            "초점 깊이": self.focal_depth_mm,
        }
        for label, value in positive.items():
            if value <= 0:
                raise ValueError(f"{label}은(는) 0보다 커야 합니다.")
        if self.active_elements > self.total_elements:
            raise ValueError("활성 소자 수는 총 소자 수보다 클 수 없습니다.")
        if self.first_element < 1:
            raise ValueError("First Element는 1 이상이어야 합니다.")
        if self.first_element + self.active_elements - 1 > self.total_elements:
            raise ValueError("First Element와 활성 소자 범위가 총 소자 수를 초과합니다.")
        if not 0 < self.min_angle_deg < 90 or not 0 < self.max_angle_deg < 90:
            raise ValueError("빔각은 0° 초과 90° 미만이어야 합니다.")
        if self.min_angle_deg > self.max_angle_deg:
            raise ValueError("최소 빔각은 최대 빔각보다 클 수 없습니다.")
        law_count = (self.max_angle_deg - self.min_angle_deg) / self.angle_step_deg
        if abs(law_count - round(law_count)) > 1e-8:
            raise ValueError("빔각 범위가 각도 간격으로 정확히 나누어지지 않습니다.")
        indices = self.index_values()
        if not indices or any(value < 0 for value in indices):
            raise ValueError("Index는 0 이상의 숫자를 하나 이상 입력하세요.")


def _font(size: int = 14, bold: bool = False):
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\malgunbd.ttf" if bold else "",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def make_beam_diagram(settings: ScanSettings, output_path: str, width=1120, height=430) -> None:
    """입력값을 기반으로 개략적인 양측 주사/반사 경로 그림을 만든다."""
    settings.validate()
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = _font(14)
    small = _font(12)

    left, right = 70, width - 70
    top, bottom = 170, 330
    center = (left + right) / 2
    plate_height = bottom - top
    draw.rectangle((left, top, right, bottom), outline="#263238", width=2, fill="#FAFAFA")

    bevel_px = max(18, min(75, plate_height / math.tan(math.radians(settings.bevel_angle_deg))))
    draw.polygon(
        [(center - bevel_px, top), (center, bottom), (center + bevel_px, top)],
        fill="#ECEFF1", outline="#455A64",
    )

    scale = plate_height / settings.thickness_mm
    colors = ["#00A651", "#F4C20D", "#D81B60", "#1565C0", "#EF6C00"]
    indices = settings.index_values()
    probe_y = top - 48

    for side in (-1, 1):
        for idx_pos, index_mm in enumerate(indices):
            probe_x = center + side * (bevel_px + index_mm * scale * 0.72)
            probe_w, probe_h = 64, 38
            draw.rectangle(
                (probe_x - probe_w / 2, probe_y, probe_x + probe_w / 2, probe_y + probe_h),
                fill="#A7A4FF", outline="#303F9F", width=2,
            )
            draw.text((probe_x - 27, probe_y - 18), f"Index {index_mm:g}", fill="#212121", font=small)

            for angle_index, angle in enumerate((settings.min_angle_deg, settings.max_angle_deg)):
                color = colors[(idx_pos * 2 + angle_index) % len(colors)]
                horizontal = settings.thickness_mm * math.tan(math.radians(angle)) * scale
                start = (probe_x, top)
                hit_x = probe_x - side * horizontal
                hit = (hit_x, bottom)
                draw.line((start, hit), fill=color, width=3)
                reflected_x = hit_x - side * horizontal
                draw.line((hit, (reflected_x, top)), fill=color, width=2)

    draw.line((left, top, left - 22, top), fill="#2E7D32", width=2)
    draw.line((left, bottom, left - 22, bottom), fill="#2E7D32", width=2)
    draw.line((left - 15, top, left - 15, bottom), fill="#2E7D32", width=2)
    draw.text((10, (top + bottom) / 2 - 8), f"t={settings.thickness_mm:g} mm", fill="#2E7D32", font=font)
    title = (
        f"{settings.probe_model} | {settings.wave_type} | "
        f"{settings.min_angle_deg:g}°~{settings.max_angle_deg:g}° | {settings.scan_direction}"
    )
    draw.text((width / 2, 28), title, fill="#111827", font=_font(18, True), anchor="ma")
    draw.text(
        (width / 2, height - 35),
        "개략도: 실제 적용 전 웨지 출사점, 용접부 형상 및 교정시험편으로 커버리지를 확인할 것",
        fill="#B91C1C", font=small, anchor="ma",
    )
    image.save(output_path, format="PNG")


def create_workbook(settings: ScanSettings, output_path: str) -> str:
    settings.validate()
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PAUT 설정표"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws.print_area = "A1:K42"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35

    navy = "1F4E78"
    blue = "D9EAF7"
    pale = "EAF3F8"
    input_fill = "FFF2CC"
    white = "FFFFFF"
    thin = Side(style="thin", color="8091A5")
    medium = Side(style="medium", color="1F2937")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    outer = Border(left=medium, right=medium, top=medium, bottom=medium)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = [18, 15, 16, 18, 22, 20, 18, 18, 18, 25, 25]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.merge_cells("A1:K2")
    ws["A1"] = "PAUT 검사 설정 및 빔 방향 검토표"
    ws["A1"].font = Font(name="맑은 고딕", size=18, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 30

    def section(row: int, text: str):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
        cell = ws.cell(row, 1, text)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.font = Font(name="맑은 고딕", size=12, bold=True, color="17365D")
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, 12):
            ws.cell(row, col).border = outer

    section(4, "1. 기본 입력조건")
    basic = [
        ("부재 두께 (mm)", settings.thickness_mm, "공칭 개선각 (°)", settings.bevel_angle_deg,
         "개선각 공차 (±°)", settings.bevel_tolerance_deg, "Angle step (°)", settings.angle_step_deg),
        ("최소 빔각 (°)", settings.min_angle_deg, "최대 빔각 (°)", settings.max_angle_deg,
         "검사 방향", settings.scan_direction, "Wave type", settings.wave_type),
    ]
    for r, row_data in enumerate(basic, start=5):
        for label_col, value_col, label, value in (
            (1, 2, row_data[0], row_data[1]), (4, 5, row_data[2], row_data[3]),
            (7, 8, row_data[4], row_data[5]), (10, 11, row_data[6], row_data[7]),
        ):
            ws.cell(r, label_col, label)
            ws.cell(r, value_col, value)
            ws.cell(r, label_col).fill = PatternFill("solid", fgColor=pale)
            ws.cell(r, value_col).fill = PatternFill("solid", fgColor=input_fill)
            for col in (label_col, value_col):
                ws.cell(r, col).border = border
                ws.cell(r, col).alignment = center
                ws.cell(r, col).font = Font(name="맑은 고딕", size=10, bold=(col == label_col))

    headers = ["Probe", "Wave type", "Law config.", "Focus type", "Active aperture",
               "Sweep angle range", "Law count", "Focal depth", "Index 1", "Index 2", "Index 3+"]
    for col, value in enumerate(headers, 1):
        cell = ws.cell(8, col, value)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="맑은 고딕", size=9, bold=True, color=white)
        cell.alignment = center
        cell.border = border
    indices = settings.index_values()
    index_values = indices + [None] * max(0, 3 - len(indices))
    values = [settings.probe_model, settings.wave_type, settings.law_config, settings.focus_type,
              None, f"{settings.min_angle_deg:g}°~{settings.max_angle_deg:g}°", None,
              settings.focal_depth_mm, index_values[0], index_values[1],
              " / ".join(f"{v:g}" for v in index_values[2:] if v is not None)]
    for col, value in enumerate(values, 1):
        ws.cell(9, col, value)
        ws.cell(9, col).alignment = center
        ws.cell(9, col).border = border
        ws.cell(9, col).font = Font(name="맑은 고딕", size=10)
    ws["E9"] = "=F14*E14"
    ws["G9"] = "=INT((E6-B6)/K5)+1"
    ws["E9"].number_format = '0.00" mm"'
    ws["G9"].number_format = '0" laws"'
    ws.row_dimensions[8].height = 31
    ws.row_dimensions[9].height = 38

    section(12, "2. 탐촉자 및 Focal Law 정보")
    probe_headers = ["탐촉자 형식", "주파수 (MHz)", "진동자 수", "Total aperture (mm)",
                     "Pitch (mm)", "Active elements", "First Element", "Focusing Type",
                     "Focal Depth (mm)", "Index", "비고"]
    for col, value in enumerate(probe_headers, 1):
        cell = ws.cell(13, col, value)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="맑은 고딕", size=9, bold=True, color=white)
        cell.alignment = center
        cell.border = border
    probe_values = [settings.probe_model, settings.frequency_mhz, settings.total_elements, None,
                    settings.pitch_mm, settings.active_elements, settings.first_element,
                    settings.focus_type, settings.focal_depth_mm, settings.index_offsets, None]
    for col, value in enumerate(probe_values, 1):
        ws.cell(14, col, value)
        ws.cell(14, col).alignment = center
        ws.cell(14, col).border = border
        ws.cell(14, col).font = Font(name="맑은 고딕", size=10)
    ws["D14"] = "=C14*E14"
    ws["K14"] = '=F14&" elements × "&TEXT(E14,"0.00")&" mm = "&TEXT(F14*E14,"0.00")&" mm"'
    ws["D14"].number_format = '0.00" mm"'
    ws["A14"].comment = Comment(f"Manufacturer source: {PROBE_SOURCE}", "User")

    section(16, "3. PAUT Scan Plan")
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
        diagram_path = temp_file.name
    try:
        make_beam_diagram(settings, diagram_path)
        diagram = XLImage(diagram_path)
        diagram.width = 1000
        diagram.height = 384
        ws.add_image(diagram, "A17")
        for row in range(17, 37):
            ws.row_dimensions[row].height = 16

        section(38, "4. 빔 경로 계산 및 검토사항")
        calc_headers = ["항목", "최소각", "최대각", "검토"]
        for col, value in enumerate(calc_headers, 1):
            ws.cell(39, col, value)
            ws.cell(39, col).fill = PatternFill("solid", fgColor=navy)
            ws.cell(39, col).font = Font(name="맑은 고딕", size=9, bold=True, color=white)
            ws.cell(39, col).alignment = center
            ws.cell(39, col).border = border
        ws["A40"] = "1-leg 표면거리"
        ws["B40"] = "=B5*TAN(RADIANS(B6))"
        ws["C40"] = "=B5*TAN(RADIANS(E6))"
        ws["D40"] = "내면(ID) 도달 시 표면 투영거리"
        ws["A41"] = "Full-skip 거리"
        ws["B41"] = "=2*B40"
        ws["C41"] = "=2*C40"
        ws["D41"] = "1회 반사 후 상면 도달거리"
        ws["A42"] = "주의"
        ws.merge_cells("B42:K42")
        ws["B42"] = "실제 적용 전 웨지 출사점, 용접부 폭·캡·HAZ 및 교정시험편으로 커버리지를 확인할 것"
        for row in range(40, 43):
            for col in range(1, 12):
                cell = ws.cell(row, col)
                cell.border = border
                cell.alignment = center
                cell.font = Font(name="맑은 고딕", size=9, color="B91C1C" if row == 42 else "000000")
        for cell in (ws["B40"], ws["C40"], ws["B41"], ws["C41"]):
            cell.number_format = '0.00" mm"'

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.save(output)
    finally:
        try:
            os.unlink(diagram_path)
        except OSError:
            pass
    return str(output)


class PautScanPlanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1080x760")
        self.minsize(980, 700)
        self.vars: dict[str, tk.StringVar] = {}
        self._build_ui()
        self._set_values(self._load_default_settings())
        self.after(150, self.refresh_preview)

    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("맑은 고딕", 16, "bold"))
        style.configure("Section.TLabelframe.Label", font=("맑은 고딕", 10, "bold"))

        header = ttk.Frame(self, padding=(16, 12))
        header.pack(fill="x")
        ttk.Label(header, text="PAUT 검사설정표 생성기", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="입력 → 검증 → 빔 경로 확인 → Excel 생성").pack(side="left", padx=18)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        form_outer = ttk.Frame(body, padding=4)
        preview_outer = ttk.Frame(body, padding=4)
        body.add(form_outer, weight=2)
        body.add(preview_outer, weight=3)

        canvas = tk.Canvas(form_outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_outer, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas)
        form.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        groups = [
            ("기본 조건", [
                ("thickness_mm", "부재 두께 (mm)"), ("bevel_angle_deg", "공칭 개선각 (°)"),
                ("bevel_tolerance_deg", "개선각 공차 (±°)"), ("scan_direction", "검사 방향"),
            ]),
            ("탐촉자", [
                ("probe_model", "탐촉자 모델"), ("frequency_mhz", "주파수 (MHz)"),
                ("total_elements", "총 소자 수"), ("pitch_mm", "Pitch (mm)"),
                ("elevation_mm", "Elevation (mm)"), ("active_elements", "활성 소자 수"),
                ("first_element", "First Element"),
            ]),
            ("Focal Law", [
                ("min_angle_deg", "최소 빔각 (°)"), ("max_angle_deg", "최대 빔각 (°)"),
                ("angle_step_deg", "Angle step (°)"), ("wave_type", "Wave type"),
                ("law_config", "Law config."), ("focus_type", "Focus type"),
                ("focal_depth_mm", "Focal depth (mm)"), ("index_offsets", "Index (쉼표 구분)"),
            ]),
        ]
        combo_values = {
            "scan_direction": ("양측 주사", "단측 주사"),
            "wave_type": ("SW", "LW"),
            "law_config": ("Sectorial", "Linear"),
            "focus_type": ("True depth", "Sound path", "Projection"),
        }
        for title, items in groups:
            box = ttk.LabelFrame(form, text=title, style="Section.TLabelframe", padding=10)
            box.pack(fill="x", pady=5)
            for row, (key, label) in enumerate(items):
                ttk.Label(box, text=label, width=20).grid(row=row, column=0, sticky="w", pady=3)
                var = tk.StringVar()
                self.vars[key] = var
                if key in combo_values:
                    widget = ttk.Combobox(box, textvariable=var, values=combo_values[key], state="readonly", width=22)
                else:
                    widget = ttk.Entry(box, textvariable=var, width=25)
                widget.grid(row=row, column=1, sticky="ew", pady=3)
                widget.bind("<FocusOut>", lambda _e: self.refresh_preview())
                widget.bind("<Return>", lambda _e: self.refresh_preview())
            box.columnconfigure(1, weight=1)

        action = ttk.Frame(form, padding=(0, 10))
        action.pack(fill="x")
        ttk.Button(action, text="설정 불러오기", command=self.load_settings).pack(side="left", padx=3)
        ttk.Button(action, text="설정 저장", command=self.save_settings).pack(side="left", padx=3)
        ttk.Button(action, text="기본값", command=self.reset_defaults).pack(side="left", padx=3)

        preview_box = ttk.LabelFrame(preview_outer, text="빔 경로 미리보기", padding=8)
        preview_box.pack(fill="both", expand=True)
        self.preview = tk.Canvas(preview_box, bg="white", highlightthickness=1, highlightbackground="#CBD5E1")
        self.preview.pack(fill="both", expand=True)
        self.preview.bind("<Configure>", lambda _e: self.refresh_preview())
        self.status = tk.StringVar(value="입력값을 확인하세요.")
        ttk.Label(preview_outer, textvariable=self.status, foreground="#334155").pack(fill="x", pady=6)

        footer = ttk.Frame(self, padding=(12, 0, 12, 12))
        footer.pack(fill="x")
        ttk.Button(footer, text="Excel 생성", command=self.export_excel).pack(side="right", ipadx=18, ipady=6)
        ttk.Button(footer, text="미리보기 갱신", command=self.refresh_preview).pack(side="right", padx=8, ipady=6)

    def _load_default_settings(self) -> ScanSettings:
        if CONFIG_PATH.exists():
            try:
                return ScanSettings(**json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
            except Exception:
                pass
        return ScanSettings()

    def _set_values(self, settings: ScanSettings):
        for item in fields(settings):
            self.vars[item.name].set(str(getattr(settings, item.name)))

    def _get_settings(self) -> ScanSettings:
        numeric_float = {
            "frequency_mhz", "pitch_mm", "elevation_mm", "thickness_mm",
            "bevel_angle_deg", "bevel_tolerance_deg", "min_angle_deg",
            "max_angle_deg", "angle_step_deg", "focal_depth_mm",
        }
        numeric_int = {"total_elements", "active_elements", "first_element"}
        data = {}
        for item in fields(ScanSettings):
            value = self.vars[item.name].get().strip()
            if item.name in numeric_float:
                data[item.name] = float(value)
            elif item.name in numeric_int:
                data[item.name] = int(float(value))
            else:
                data[item.name] = value
        settings = ScanSettings(**data)
        settings.validate()
        return settings

    def refresh_preview(self):
        self.preview.delete("all")
        try:
            settings = self._get_settings()
        except Exception as exc:
            self.status.set(f"입력 오류: {exc}")
            self.preview.create_text(20, 20, text=str(exc), anchor="nw", fill="#B91C1C")
            return
        width = max(400, self.preview.winfo_width())
        height = max(300, self.preview.winfo_height())
        margin = 45
        top, bottom = height * 0.42, height * 0.75
        center = width / 2
        self.preview.create_rectangle(margin, top, width - margin, bottom, fill="#F8FAFC", outline="#334155", width=2)
        plate_h = bottom - top
        bevel = max(15, min(60, plate_h / math.tan(math.radians(settings.bevel_angle_deg))))
        self.preview.create_polygon(center - bevel, top, center, bottom, center + bevel, top,
                                    fill="#E2E8F0", outline="#475569")
        scale = plate_h / settings.thickness_mm
        colors = ("#16A34A", "#D97706", "#DB2777", "#2563EB")
        for side in (-1, 1):
            for pos, index in enumerate(settings.index_values()):
                x = center + side * (bevel + index * scale * 0.55)
                self.preview.create_rectangle(x - 24, top - 34, x + 24, top, fill="#A5B4FC", outline="#3730A3")
                for angle_pos, angle in enumerate((settings.min_angle_deg, settings.max_angle_deg)):
                    dx = settings.thickness_mm * math.tan(math.radians(angle)) * scale
                    hit_x = x - side * dx
                    color = colors[(pos + angle_pos) % len(colors)]
                    self.preview.create_line(x, top, hit_x, bottom, fill=color, width=2)
                    self.preview.create_line(hit_x, bottom, hit_x - side * dx, top, fill=color, width=2)
        law_count = int(round((settings.max_angle_deg - settings.min_angle_deg) / settings.angle_step_deg)) + 1
        active_aperture = settings.active_elements * settings.pitch_mm
        self.status.set(f"검증 완료 | Focal laws {law_count}개 | 활성 개구 {active_aperture:.2f} mm")

    def reset_defaults(self):
        self._set_values(ScanSettings())
        self.refresh_preview()

    def save_settings(self):
        try:
            settings = self._get_settings()
        except Exception as exc:
            messagebox.showerror("입력 오류", str(exc), parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="설정 저장", defaultextension=".json",
            initialfile="PAUT_검사설정.json", filetypes=[("JSON", "*.json")],
        )
        if path:
            Path(path).write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")

    def load_settings(self):
        path = filedialog.askopenfilename(parent=self, title="설정 불러오기", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            settings = ScanSettings(**json.loads(Path(path).read_text(encoding="utf-8")))
            settings.validate()
            self._set_values(settings)
            self.refresh_preview()
        except Exception as exc:
            messagebox.showerror("불러오기 오류", str(exc), parent=self)

    def export_excel(self):
        try:
            settings = self._get_settings()
        except Exception as exc:
            messagebox.showerror("입력 오류", str(exc), parent=self)
            return
        default_name = (
            f"PAUT_검사설정_{settings.probe_model}_{settings.thickness_mm:g}mm_"
            f"{settings.min_angle_deg:g}-{settings.max_angle_deg:g}도.xlsx"
        )
        path = filedialog.asksaveasfilename(
            parent=self, title="Excel 저장", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            create_workbook(settings, path)
            CONFIG_PATH.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("완료", f"검사설정표를 생성했습니다.\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("생성 오류", str(exc), parent=self)


def self_test(output_path: str) -> None:
    settings = ScanSettings()
    created = create_workbook(settings, output_path)
    check = openpyxl.load_workbook(created, data_only=False)
    sheet = check["PAUT 설정표"]
    assert sheet["B5"].value == 15.87
    assert sheet["E9"].value == "=F14*E14"
    assert sheet["G9"].value == "=INT((E6-B6)/K5)+1"
    assert len(sheet._images) == 1
    assert sheet.print_area == "'PAUT 설정표'!$A$1:$K$42"
    check.close()


def main():
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", metavar="OUTPUT", help="GUI 없이 샘플 Excel을 생성하고 검증")
    args = parser.parse_args()
    if args.self_test:
        self_test(args.self_test)
        print(f"SELF TEST OK: {args.self_test}")
        return
    app = PautScanPlanApp()
    app.mainloop()


if __name__ == "__main__":
    main()
