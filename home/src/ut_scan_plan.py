# -*- coding: utf-8 -*-
"""초음파탐상시험(UT) 검사설정표 및 스캔플랜 생성 모듈.

사각탐상(Angle Beam, 45°/60°/70°) 및 수직탐상(Straight Beam, 0°)을 모두 지원하며,
모재 두께, 베벨 각도, 비드 폭에 따른 0.5/1.0 Skip 거리, Sound Path, 주사 범위를 자동 계산하고
빔 다이어그램 시각화 및 openpyxl 기반 엑셀 검사설정표를 생성합니다.
"""

from __future__ import annotations

import json
import math
import os
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


# ==============================================================================
# 1. 데이터 모델 (Data Models)
# ==============================================================================

@dataclass
class UTProbeConfig:
    """초음파 탐촉자 설정."""
    probe_id: str = "P1"
    probe_type: str = "사각 (Angle)"           # "사각 (Angle)" 또는 "수직 (Straight)"
    angle_deg: float = 45.0                     # 0.0 (수직), 45.0, 60.0, 70.0 (사각)
    frequency_mhz: float = 4.0                  # 2.0, 4.0, 5.0 MHz 등
    crystal_size_mm: str = "10×10"              # "10×10", "Φ10", "Φ20" 등
    wave_type: str = "SW (횡파)"                # "SW (횡파)" 또는 "LW (종파)"
    velocity_ms: float = 3240.0                 # 탄소강: 횡파 3240 m/s, 종파 5920 m/s
    index_offset_mm: float = 10.0               # 탐촉자 전단부에서 입사점까지 거리 (사각)
    probe_model: str = "4Z10x10A45"             # 모델명
    wedge_delay_us: float = 0.0                 # 웻지 지연 시간
    is_active: bool = True                      # 활성화 여부


@dataclass
class UTScanPlanSettings:
    """UT 검사설정표 및 스캔플랜 종합 설정."""
    # 관리 정보
    project_name: str = "한국지역난방공사 중앙지사 열수송관 공사"
    report_title: str = "초음파탐상시험(UT) 검사설정표 및 스캔플랜"
    applicable_spec: str = "KS B 0896 / ASME Sec.V Art.4"
    acceptance_criteria: str = "KS B 0896 2류 / ASME B31.1"
    inspection_date: str = "2026-11-03"
    inspector_name: str = "비파괴검사원"
    inspector_cert: str = "UT Level II"

    # 시험체 정보
    component_name: str = "열배관 현장 원주 용접부 (Circumferential Weld)"
    material_spec: str = "SPPS 380 / Carbon Steel"
    thickness_mm: float = 15.87
    pipe_od_mm: float = 318.5                   # 배관 외경 (0 또는 음수이면 평판)
    bevel_type: str = "Single-V (V형 맞대기)"
    bevel_angle_deg: float = 35.0               # V홈 편각 (포함각 70°)
    root_gap_mm: float = 3.0                    # 루트 간격
    root_face_mm: float = 1.5                   # 루트 면
    cap_width_mm: float = 20.0                  # 비드 폭
    cap_height_mm: float = 2.0                  # 덧살 높이
    haz_width_mm: float = 10.0                  # 열영향부(HAZ) 검사 확장폭

    # 탐상 장비 및 교정 정보
    instrument_name: str = "EPOCH 650 (Olympus)"
    couplant: str = "초음파 탐상 전용 글리세린 / 페이스트"
    cal_block: str = "STB-A1 / IIW Type 1 & ASME Basic Block"
    cal_method: str = "DAC (거리진폭보정곡선)"
    reference_level: str = "DAC 100% (기준감도 H선)"
    scan_level: str = "기준감도(DAC) + 6 dB"
    transfer_correction_db: float = 2.0         # 전송 손실 보정

    # 탐촉자 목록 (기본: 사각 45°, 60°, 70° 및 수직 0°)
    probes: List[UTProbeConfig] = field(default_factory=lambda: [
        UTProbeConfig(
            probe_id="P1",
            probe_type="사각 (Angle)",
            angle_deg=45.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=10.0,
            probe_model="4Z10x10A45",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P2",
            probe_type="사각 (Angle)",
            angle_deg=60.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=11.0,
            probe_model="4Z10x10A60",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P3",
            probe_type="사각 (Angle)",
            angle_deg=70.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=12.0,
            probe_model="4Z10x10A70",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P4",
            probe_type="수직 (Straight)",
            angle_deg=0.0,
            frequency_mhz=4.0,
            crystal_size_mm="Φ10",
            wave_type="LW (종파)",
            velocity_ms=5920.0,
            index_offset_mm=0.0,
            probe_model="4Z10N",
            is_active=True
        ),
    ])

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "UTScanPlanSettings":
        data = dict(data)
        raw_probes = data.pop("probes", None)
        probes = []
        if raw_probes:
            for p in raw_probes:
                if isinstance(p, UTProbeConfig):
                    probes.append(p)
                else:
                    probes.append(UTProbeConfig(**p))
        else:
            probes = [
                UTProbeConfig("P1", "사각 (Angle)", 45.0, 4.0, "10×10", "SW (횡파)", 3240.0, 10.0, "4Z10x10A45"),
                UTProbeConfig("P2", "사각 (Angle)", 60.0, 4.0, "10×10", "SW (횡파)", 3240.0, 11.0, "4Z10x10A60"),
                UTProbeConfig("P3", "사각 (Angle)", 70.0, 4.0, "10×10", "SW (횡파)", 3240.0, 12.0, "4Z10x10A70"),
                UTProbeConfig("P4", "수직 (Straight)", 0.0, 4.0, "Φ10", "LW (종파)", 5920.0, 0.0, "4Z10N"),
            ]
        data["probes"] = probes
        return cls(**data)

    def to_dict(self) -> Dict[str, Any]:
        result = asdict(self)
        return result


# ==============================================================================
# 2. 기하학적 계산 엔진 (Geometric Calculation Engine)
# ==============================================================================

@dataclass
class UTCalculationResult:
    """단일 탐촉자에 대한 기하학 계산 결과."""
    probe_id: str
    probe_type: str
    angle_deg: float
    skip_05_mm: float           # 0.5 Skip 거리 (mm)
    skip_10_mm: float           # 1.0 Skip 거리 (mm)
    sound_path_05_mm: float     # 0.5 Skip Sound Path (mm)
    sound_path_10_mm: float     # 1.0 Skip Sound Path (mm)
    scan_min_mm: float          # 탐촉자 입사점 최소 거리 (mm, 용접중심선 기준)
    scan_max_mm: float          # 탐촉자 입사점 최대 거리 (mm, 용접중심선 기준)
    scan_range_desc: str        # 주사 범위 요약 설명
    divergence_deg: float       # 빔 확산각 (반각, 도)
    recommended_range_mm: float # 장비 시간축 권장 Range (mm)
    coverage_target: str        # 주 검출 대상 영역


def calculate_ut_geometry(settings: UTScanPlanSettings) -> List[UTCalculationResult]:
    """설정된 모든 활성 탐촉자에 대해 Skip, Sound Path, 주사 범위를 계산한다."""
    t = float(settings.thickness_mm)
    half_cap = float(settings.cap_width_mm) / 2.0
    haz = float(settings.haz_width_mm)

    results = []
    for probe in settings.probes:
        if not probe.is_active:
            continue

        theta = float(probe.angle_deg)
        if probe.probe_type.startswith("수직") or theta == 0.0:
            # 수직 탐상 (Straight Beam, 0°)
            skip_05 = 0.0
            skip_10 = 0.0
            sp_05 = t
            sp_10 = 2.0 * t
            # 모재 라미네이션 및 두께 검사: 용접 비드 양쪽 모재부
            scan_min = half_cap + float(probe.index_offset_mm)
            scan_max = half_cap + haz + 50.0
            scan_desc = f"{scan_min:.1f} ~ {scan_max:.1f} mm (양측 모재 검사)"
            rec_range = 3.0 * t
            target = "모재 라미네이션, 판두께 측정, 용접선 인접부 모재 결함"
            divergence = 2.5
        else:
            # 사각 탐상 (Angle Beam, 45°, 60°, 70° 등)
            rad = math.radians(theta)
            tan_theta = math.tan(rad)
            cos_theta = math.cos(rad)

            skip_05 = t * tan_theta
            skip_10 = 2.0 * t * tan_theta
            sp_05 = t / cos_theta
            sp_10 = (2.0 * t) / cos_theta

            # 빔 확산각 근사 계산
            # 파장 lambda = v / f
            v = probe.velocity_ms
            f = probe.frequency_mhz * 1e6
            wavelength_mm = (v / f) * 1000.0  # mm
            # 진동자 폭 파싱 (예: "10x10" -> 10.0)
            elem_w = 10.0
            size_clean = probe.crystal_size_mm.replace("×", "x").replace("Φ", "")
            if "x" in size_clean:
                try:
                    elem_w = float(size_clean.split("x")[0])
                except ValueError:
                    pass
            elif size_clean:
                try:
                    elem_w = float(size_clean)
                except ValueError:
                    pass
            try:
                divergence = math.degrees(math.asin(min(1.0, 0.7 * wavelength_mm / elem_w)))
            except Exception:
                divergence = 3.0

            # 주사 범위 산출:
            # 비드 토우 접근 한계 = half_cap + probe.index_offset_mm
            # 0.5 스킵으로 루트부 탐상 = skip_05
            # 1.0 스킵으로 상단 HAZ/베벨 탐상 = skip_10 + half_cap + haz
            scan_min = max(half_cap + float(probe.index_offset_mm), skip_05 - half_cap)
            scan_max = skip_10 + half_cap + haz
            scan_desc = f"{scan_min:.1f} ~ {scan_max:.1f} mm (0.5~1.0 Skip)"
            rec_range = sp_10 * 1.25

            if theta >= 68.0:
                target = "용접부 상층부 및 표면/언더컷, 표면 근처 HAZ"
            elif theta >= 55.0:
                target = "용접부 중간층, 융합불량(Lack of Fusion), 슬래그 혼입"
            else:
                target = "루트부(Lack of Penetration, Root Crack) 및 대향면 결함"

        results.append(UTCalculationResult(
            probe_id=probe.probe_id,
            probe_type=probe.probe_type,
            angle_deg=theta,
            skip_05_mm=round(skip_05, 2),
            skip_10_mm=round(skip_10, 2),
            sound_path_05_mm=round(sp_05, 2),
            sound_path_10_mm=round(sp_10, 2),
            scan_min_mm=round(scan_min, 2),
            scan_max_mm=round(scan_max, 2),
            scan_range_desc=scan_desc,
            divergence_deg=round(divergence, 1),
            recommended_range_mm=round(rec_range, 1),
            coverage_target=target
        ))

    return results


# ==============================================================================
# 3. 빔 다이어그램 시각화 (Pillow Beam Diagram Renderer)
# ==============================================================================

def make_ut_beam_diagram(
    settings: UTScanPlanSettings,
    output_path: str,
    width: int = 1160,
    height: int = 360,
) -> None:
    """모재 단면, 용접부, 사각/수직 탐촉자 빔 경로를 도식화한 이미지를 생성한다."""
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    def get_font(size: int = 12, bold: bool = False):
        font_names = [
            "malgunbd.ttf" if bold else "malgun.ttf",
            "gulim.ttc",
            "arialbd.ttf" if bold else "arial.ttf",
        ]
        for name in font_names:
            try:
                return ImageFont.truetype(name, size)
            except OSError:
                continue
        return ImageFont.load_default()

    font_title = get_font(15, bold=True)
    font_bold = get_font(11, bold=True)
    font_normal = get_font(10, bold=False)
    font_small = get_font(9, bold=False)

    margin_left = 110
    margin_right = 90
    plate_top = 115
    t = float(settings.thickness_mm)

    # 1.0 Skip 최대 거리를 감안한 스케일 계산
    max_span_mm = max(90.0, t * math.tan(math.radians(70.0)) * 2.1 + float(settings.cap_width_mm) + 25.0)
    avail_width = width - margin_left - margin_right
    scale = avail_width / (max_span_mm * 1.3)
    plate_bottom = plate_top + max(30, int(t * scale))

    # 용접 중심선 x좌표
    center_x = margin_left + int(avail_width * 0.40)

    # 1. 배경 타이틀 바
    draw.rectangle([8, 8, width - 8, 38], fill=(240, 244, 250), outline=(180, 200, 225), width=1)
    pipe_info = f"Pipe OD {settings.pipe_od_mm:g}mm" if settings.pipe_od_mm > 0 else "Plate"
    title_text = f"UT Scan Plan Diagram  |  {settings.component_name} (T={t:g}mm, {pipe_info}, Bevel {settings.bevel_angle_deg:g}°)"
    draw.text((20, 14), title_text, fill=(20, 50, 100), font=font_title)

    # 2. 모재 및 용접부 단면
    left_x = margin_left - 25
    right_x = width - margin_right + 25
    draw.rectangle([left_x, plate_top, right_x, plate_bottom], fill=(248, 249, 250), outline=(120, 130, 140), width=2)

    half_cap = (float(settings.cap_width_mm) / 2.0) * scale
    half_root_gap = (float(settings.root_gap_mm) / 2.0) * scale
    root_face = float(settings.root_face_mm) * scale
    cap_height = float(settings.cap_height_mm) * scale
    haz_px = float(settings.haz_width_mm) * scale

    weld_poly = [
        (center_x - half_cap, plate_top),
        (center_x, plate_top - cap_height),
        (center_x + half_cap, plate_top),
        (center_x + half_root_gap, plate_bottom - root_face),
        (center_x + half_root_gap, plate_bottom),
        (center_x - half_root_gap, plate_bottom),
        (center_x - half_root_gap, plate_bottom - root_face),
    ]
    draw.polygon(weld_poly, fill=(235, 238, 245), outline=(90, 110, 140))
    draw.line([(center_x - half_cap, plate_top), (center_x, plate_top - cap_height), (center_x + half_cap, plate_top)],
              fill=(70, 90, 120), width=2)

    # HAZ 점선
    haz_left = center_x - half_cap - haz_px
    haz_right = center_x + half_cap + haz_px
    for y in range(plate_top, plate_bottom, 5):
        draw.line([(haz_left, y), (haz_left, min(y + 2, plate_bottom))], fill=(190, 160, 120), width=1)
        draw.line([(haz_right, y), (haz_right, min(y + 2, plate_bottom))], fill=(190, 160, 120), width=1)

    # 용접 중심선
    for y in range(plate_top - 25, plate_bottom + 30, 7):
        draw.line([(center_x, y), (center_x, y + 4)], fill=(220, 60, 60), width=1)
    draw.text((center_x - 28, plate_bottom + 32), "용접 중심선", fill=(200, 40, 40), font=font_small)

    # 두께 치수선
    dim_x = left_x + 12
    draw.line([(dim_x, plate_top), (dim_x, plate_bottom)], fill=(80, 80, 80), width=1)
    draw.line([(dim_x - 3, plate_top), (dim_x + 3, plate_top)], fill=(80, 80, 80), width=1)
    draw.line([(dim_x - 3, plate_bottom), (dim_x + 3, plate_bottom)], fill=(80, 80, 80), width=1)
    draw.text((dim_x - 55, (plate_top + plate_bottom) // 2 - 6), f"T={t:g}mm", fill=(50, 50, 50), font=font_bold)

    # 3. 빔 경로 및 주사 범위
    calculations = calculate_ut_geometry(settings)
    colors = [
        (30, 100, 220),   # 45도 파랑
        (30, 160, 70),    # 60도 초록
        (220, 110, 20),   # 70도 주황
        (150, 50, 180),   # 0도 보라
    ]
    calc_by_angle = {res.angle_deg: res for res in calculations}

    # 범례 박스
    legend_y = height - 85
    draw.rectangle([left_x, legend_y - 6, right_x, height - 10], fill=(250, 250, 252), outline=(210, 215, 220), width=1)
    draw.text((left_x + 10, legend_y - 2), "[범례 및 계산 결과 요약]", fill=(40, 40, 40), font=font_bold)

    angle_list = [45.0, 60.0, 70.0, 0.0]
    drawn_count = 0
    for idx, angle in enumerate(angle_list):
        if angle not in calc_by_angle:
            continue
        c_res = calc_by_angle[angle]
        color = colors[idx % len(colors)]

        col_x = left_x + 10 + (drawn_count % 2) * 460
        row_y = legend_y + 18 + (drawn_count // 2) * 28
        drawn_count += 1

        draw.rectangle([col_x, row_y + 2, col_x + 14, row_y + 12], fill=color)
        if angle == 0.0:
            leg_txt = f"수직 0°: Sound Path={c_res.sound_path_05_mm:.1f}mm | 주사범위: {c_res.scan_range_desc}"
        else:
            leg_txt = (f"사각 {angle:g}°: 0.5S={c_res.skip_05_mm:.1f}mm, 1.0S={c_res.skip_10_mm:.1f}mm | "
                       f"SP={c_res.sound_path_10_mm:.1f}mm | 주사: {c_res.scan_min_mm:.1f}~{c_res.scan_max_mm:.1f}mm")
        draw.text((col_x + 20, row_y), leg_txt, fill=(40, 40, 40), font=font_normal)

        if angle == 0.0:
            str_probe_x = center_x - half_cap - haz_px - int(22.0 * scale)
            probe_w, probe_h = 20, 16
            draw.rectangle([str_probe_x - probe_w // 2, plate_top - probe_h, str_probe_x + probe_w // 2, plate_top],
                           fill=(230, 220, 245), outline=color, width=2)
            draw.line([(str_probe_x, plate_top), (str_probe_x, plate_bottom)], fill=color, width=2)
            draw.text((str_probe_x - 16, plate_top - probe_h - 14), "0° 수직", fill=color, font=font_small)
            draw.line([(str_probe_x - 3, plate_bottom - 5), (str_probe_x, plate_bottom), (str_probe_x + 3, plate_bottom - 5)], fill=color, width=2)
        else:
            rad = math.radians(angle)
            skip_05_px = int(c_res.skip_05_mm * scale)
            skip_10_px = int(c_res.skip_10_mm * scale)

            index_x = center_x + skip_10_px
            refl_x = center_x + skip_05_px

            draw.line([(index_x, plate_top), (refl_x, plate_bottom)], fill=color, width=2)
            draw.line([(refl_x, plate_bottom), (center_x, plate_top)], fill=color, width=2)

            wedge_w = 22
            wedge_h = 15
            wedge_poly = [
                (index_x - 5, plate_top),
                (index_x + wedge_w, plate_top),
                (index_x + wedge_w, plate_top - wedge_h),
                (index_x + 3, plate_top - wedge_h),
            ]
            draw.polygon(wedge_poly, fill=(245, 245, 250), outline=color, width=2)
            draw.text((index_x + 5, plate_top - wedge_h - 13), f"{angle:g}°", fill=color, font=font_bold)

            y_dim = plate_top - 28 - int(idx * 14)
            draw.line([(center_x, y_dim), (index_x, y_dim)], fill=color, width=1)
            draw.line([(center_x, y_dim - 2), (center_x, y_dim + 2)], fill=color, width=1)
            draw.line([(refl_x, y_dim - 2), (refl_x, y_dim + 2)], fill=color, width=1)
            draw.line([(index_x, y_dim - 2), (index_x, y_dim + 2)], fill=color, width=1)
            draw.text((refl_x - 16, y_dim - 12), f"0.5S={c_res.skip_05_mm:.1f}", fill=color, font=font_small)
            draw.text((index_x - 18, y_dim - 12), f"1.0S={c_res.skip_10_mm:.1f}", fill=color, font=font_small)

    img.save(output_path, "PNG")


def create_ut_workbook(settings: UTScanPlanSettings, output_path: str) -> str:
    """UT 검사설정표 및 스캔플랜을 A4 1페이지에 완벽히 맞춘 엑셀 문서로 생성한다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UT 검사설정표"
    ws.views.sheetView[0].showGridLines = True

    # 스타일 정의
    font_main_title = Font(name="맑은 고딕", size=14, bold=True, color="1F4E78")
    font_sec_title = Font(name="맑은 고딕", size=10, bold=True, color="203764")
    font_header = Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="맑은 고딕", size=9, bold=True)
    font_normal = Font(name="맑은 고딕", size=8.5)
    font_small = Font(name="맑은 고딕", size=7.5)

    fill_sec_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_tbl_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_sub_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="BFBFBF")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom_side = Side(border_style="medium", color="2F5597")
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True)

    # 열 너비 설정 (B~L 총 11열 구성)
    col_widths = {
        "A": 3, "B": 17, "C": 16, "D": 13, "E": 13,
        "F": 13, "G": 17, "H": 14, "I": 14, "J": 14,
        "K": 17, "L": 18, "M": 4
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # 1. 메인 타이틀 (Row 2)
    ws.row_dimensions[1].height = 6
    ws.merge_cells("B2:L2")
    title_cell = ws["B2"]
    title_cell.value = "초음파탐상시험(UT) 검사설정표 및 스캔플랜"
    title_cell.font = font_main_title
    title_cell.alignment = align_center
    ws.row_dimensions[2].height = 22

    current_row = 4

    def add_section_header(title: str):
        nonlocal current_row
        ws.merge_cells(f"B{current_row}:L{current_row}")
        cell = ws[f"B{current_row}"]
        cell.value = title
        cell.font = font_sec_title
        cell.fill = fill_sec_header
        cell.alignment = align_left
        ws.row_dimensions[current_row].height = 17
        current_row += 1

    def style_range(start_col: str, start_r: int, end_col: str, end_r: int, border=border_cell):
        for r in range(start_r, end_r + 1):
            for c in range(openpyxl.utils.column_index_from_string(start_col),
                           openpyxl.utils.column_index_from_string(end_col) + 1):
                ws.cell(row=r, column=c).border = border

    def add_spacer():
        nonlocal current_row
        ws.row_dimensions[current_row].height = 5
        current_row += 1

    # --------------------------------------------------------------------------
    # 2. 섹션 1: 일반 및 관리 정보
    # --------------------------------------------------------------------------
    add_section_header("1. 일반 및 관리 정보 (General Information)")

    general_info = [
        ("공 사 명", settings.project_name, "검 사 일 자", settings.inspection_date),
        ("적 용 규 격", settings.applicable_spec, "합 격 기 준", settings.acceptance_criteria),
        ("검 사 원", f"{settings.inspector_name} ({settings.inspector_cert})", "보고서 번호", "UT-SP-2026-001"),
    ]

    for label1, val1, label2, val2 in general_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 3. 섹션 2: 시험체 및 용접부 제원
    # --------------------------------------------------------------------------
    add_section_header("2. 시험체 및 용접부 제원 (Component & Weld Details)")

    pipe_desc = f"OD {settings.pipe_od_mm:g} mm" if settings.pipe_od_mm > 0 else "평판 (Plate)"
    joint_info = [
        ("부품 / 구간", settings.component_name, "재 질 (Material)", settings.material_spec),
        ("모재 두께 (T)", f"{settings.thickness_mm:g} mm", "배관 외경 / 형태", pipe_desc),
        ("홈 형상 (Bevel)", settings.bevel_type, "홈 편각 (Angle)", f"{settings.bevel_angle_deg:g}° (포함각 {settings.bevel_angle_deg*2:g}°)"),
        ("루트 간격 (Gap)", f"{settings.root_gap_mm:g} mm", "루트 면 (Root Face)", f"{settings.root_face_mm:g} mm"),
        ("비드 폭 (Width)", f"{settings.cap_width_mm:g} mm", "덧살 높이 (Cap)", f"{settings.cap_height_mm:g} mm"),
        ("열영향부 (HAZ)", f"비드 양단 + {settings.haz_width_mm:g} mm", "표면 가공 상태", "쇼트 블라스트 / 그라인딩"),
    ]

    for label1, val1, label2, val2 in joint_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 4. 섹션 3: 탐상 장비 및 교정 조건
    # --------------------------------------------------------------------------
    add_section_header("3. 탐상 장비 및 교정 조건 (Equipment & Calibration)")

    equip_info = [
        ("초음파 탐상기", settings.instrument_name, "접 촉 매 질", settings.couplant),
        ("대비 시험편", settings.cal_block, "교정 기법", settings.cal_method),
        ("기준 감도 (Ref.)", settings.reference_level, "검사 감도 (Scan)", settings.scan_level),
        ("전송 손실 보정", f"+ {settings.transfer_correction_db:g} dB", "주사 속도 (Max)", "150 mm/s 이하"),
    ]

    for label1, val1, label2, val2 in equip_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 5. 섹션 4: 스캔플랜 및 탐촉자별 기하학 계산표
    # --------------------------------------------------------------------------
    add_section_header("4. 스캔플랜 및 기하학 계산표 (Scan Plan Calculations)")

    tbl_headers = [
        ("B", "탐촉자 구분"),
        ("C", "굴절각(°)"),
        ("D", "주파수(MHz)"),
        ("E", "진동자 크기"),
        ("F", "0.5 Skip(mm)"),
        ("G", "1.0 Skip(mm)"),
        ("H", "Sound Path(mm)"),
        ("I", "입사점(Xmin)"),
        ("J", "입사점(Xmax)"),
        ("K", "주사 범위 요약"),
        ("L", "주요 검출 대상"),
    ]

    ws.row_dimensions[current_row].height = 18
    for col, h_name in tbl_headers:
        c = ws[f"{col}{current_row}"]
        c.value = h_name
        c.font = font_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = border_header
    current_row += 1

    calc_results = calculate_ut_geometry(settings)
    probe_map = {p.probe_id: p for p in settings.probes}

    for res in calc_results:
        p_cfg = probe_map.get(res.probe_id)
        ws.row_dimensions[current_row].height = 16

        ws[f"B{current_row}"] = f"{res.probe_type} ({res.probe_id})"
        ws[f"C{current_row}"] = f"{res.angle_deg:g}°"
        ws[f"D{current_row}"] = f"{p_cfg.frequency_mhz:g}" if p_cfg else "-"
        ws[f"E{current_row}"] = p_cfg.crystal_size_mm if p_cfg else "-"
        ws[f"F{current_row}"] = f"{res.skip_05_mm:.1f}" if res.angle_deg > 0 else "-"
        ws[f"G{current_row}"] = f"{res.skip_10_mm:.1f}" if res.angle_deg > 0 else "-"
        ws[f"H{current_row}"] = f"{res.sound_path_10_mm:.1f}"
        ws[f"I{current_row}"] = f"{res.scan_min_mm:.1f}"
        ws[f"J{current_row}"] = f"{res.scan_max_mm:.1f}"
        ws[f"K{current_row}"] = res.scan_range_desc
        ws[f"L{current_row}"] = res.coverage_target

        for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            ws[f"{col}{current_row}"].alignment = align_center
            ws[f"{col}{current_row}"].font = font_bold if col in ["C", "F", "G", "H"] else font_normal
        ws[f"K{current_row}"].alignment = align_center
        ws[f"K{current_row}"].font = font_normal
        ws[f"L{current_row}"].alignment = align_left
        ws[f"L{current_row}"].font = font_small

        if res.angle_deg == 60.0:
            for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]:
                ws[f"{col}{current_row}"].fill = fill_highlight

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 6. 섹션 5: 스캔플랜 다이어그램 임베드
    # --------------------------------------------------------------------------
    add_section_header("5. 스캔플랜 빔 다이어그램 (Scan Plan Beam Path Diagram)")

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_img_path = tmp.name

    try:
        make_ut_beam_diagram(settings, tmp_img_path, width=1160, height=360)
        xl_img = XLImage(tmp_img_path)
        xl_img.width = 750
        xl_img.height = 232
        ws.add_image(xl_img, f"B{current_row}")
    finally:
        pass

    # 다이어그램 아래 빈 행들 높이 지정
    for offset in range(13):
        ws.row_dimensions[current_row + offset].height = 18

    # 인쇄 설정 (A4 가로 1페이지 강제 맞춤)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 여백 최소화 (인치 단위)
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1

    # 인쇄 정중앙 맞춤
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True

    # 인쇄 영역 확정
    ws.print_area = f"$B$2:$L${current_row + 12}"

    wb.save(output_path)
    return output_path
