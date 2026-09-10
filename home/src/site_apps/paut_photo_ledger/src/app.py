from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, XDRPositiveSize2D
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps, ImageTk


APP_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = APP_DIR / "data"
PHOTO_DIR = APP_DIR / "photos"
OUTPUT_DIR = APP_DIR / "outputs"
DB_PATH = DATA_DIR / "items.json"


class PautPhotoLedgerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PAUT Probe 및 Wedge 입고 사진대장")
        self.root.geometry("1220x760")
        self.root.minsize(980, 620)
        for folder in (DATA_DIR, PHOTO_DIR, OUTPUT_DIR):
            folder.mkdir(parents=True, exist_ok=True)
        self.data = self._load_data()
        self.current_index = 0
        self.preview_refs: list[ImageTk.PhotoImage] = []
        self.status_var = tk.StringVar(value="준비")
        self._build_ui()
        self._refresh_list()

    def _load_data(self) -> dict:
        if DB_PATH.exists():
            return json.loads(DB_PATH.read_text(encoding="utf-8"))
        return {"project": "A25 Probe 및 Wedge 입고", "received_date": "2026-09-07", "items": []}

    def _save_data(self):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        DB_PATH.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="PAUT 입고 사진대장", font=("맑은 고딕", 18, "bold")).pack(side="left")
        ttk.Button(top, text="사진 가져오기", command=self.import_photos).pack(side="right", padx=4)
        ttk.Button(top, text="자동 인식·분류", command=self.auto_classify_current).pack(side="right", padx=4)
        ttk.Button(top, text="엑셀 사진대장 출력", command=self.export_excel).pack(side="right", padx=4)
        ttk.Button(top, text="저장", command=self.save_current).pack(side="right", padx=4)

        body = ttk.Panedwindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(12, 4)).pack(fill="x", side="bottom")

        left = ttk.Frame(body, padding=6)
        right = ttk.Frame(body, padding=10)
        body.add(left, weight=1)
        body.add(right, weight=3)

        ttk.Label(left, text="입고 품목", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(0, 6))
        self.tree = ttk.Treeview(left, columns=("kind", "qty"), show="tree headings", height=25)
        self.tree.heading("#0", text="규격 / 관리번호")
        self.tree.heading("kind", text="구분")
        self.tree.heading("qty", text="사진")
        self.tree.column("#0", width=190)
        self.tree.column("kind", width=70, anchor="center")
        self.tree.column("qty", width=45, anchor="center")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        controls = ttk.Frame(left)
        controls.pack(fill="x", pady=6)
        ttk.Button(controls, text="품목 추가", command=self.add_item).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(controls, text="품목 삭제", command=self.delete_item).pack(side="left", expand=True, fill="x", padx=(3, 0))
        ttk.Button(left, text="선택 품목 합치기", command=self.merge_selected_items).pack(fill="x", pady=(0, 6))

        form = ttk.LabelFrame(right, text="품목 정보", padding=10)
        form.pack(fill="x")
        self.vars = {key: tk.StringVar() for key in ("kind", "model", "size", "serial", "quantity", "note")}
        fields = [("구분", "kind"), ("모델명", "model"), ("규격", "size"), ("S/N", "serial"), ("수량", "quantity"), ("비고", "note")]
        for i, (label, key) in enumerate(fields):
            row, col = divmod(i, 3)
            ttk.Label(form, text=label).grid(row=row * 2, column=col, sticky="w", padx=5)
            ttk.Entry(form, textvariable=self.vars[key]).grid(row=row * 2 + 1, column=col, sticky="ew", padx=5, pady=(0, 8))
            form.columnconfigure(col, weight=1)

        photo_bar = ttk.Frame(right)
        photo_bar.pack(fill="x", pady=(10, 5))
        ttk.Label(photo_bar, text="연결된 사진", font=("맑은 고딕", 11, "bold")).pack(side="left")
        ttk.Button(photo_bar, text="선택 사진 제거", command=self.remove_selected_photo).pack(side="right")
        ttk.Button(photo_bar, text="사진 추가", command=self.add_photos_to_item).pack(side="right", padx=5)

        self.canvas = tk.Canvas(right, bg="#eef2f5", highlightthickness=0)
        scroll = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        self.photo_frame = ttk.Frame(self.canvas)
        self.photo_frame.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.photo_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _refresh_list(self, select: int | None = None):
        self.tree.delete(*self.tree.get_children())
        for i, item in enumerate(self.data.get("items", [])):
            title = item.get("size") or item.get("model") or "미분류"
            self.tree.insert("", "end", iid=str(i), text=title, values=(item.get("kind", ""), len(item.get("photos", []))))
        if self.data.get("items"):
            idx = min(self.current_index if select is None else select, len(self.data["items"]) - 1)
            self.tree.selection_set(str(idx))
            self.tree.focus(str(idx))
            self._show_item(idx)
        else:
            self._clear_form()

    def _on_select(self, _event=None):
        selected = self.tree.selection()
        if selected:
            self.save_current(silent=True)
            self._show_item(int(selected[0]))

    def _show_item(self, index: int):
        self.current_index = index
        item = self.data["items"][index]
        for key, var in self.vars.items():
            var.set(str(item.get(key, "")))
        self._render_photos(item.get("photos", []))

    def _clear_form(self):
        for var in self.vars.values():
            var.set("")
        self._render_photos([])

    def _render_photos(self, photos: list[str]):
        for child in self.photo_frame.winfo_children():
            child.destroy()
        self.preview_refs.clear()
        self.photo_selection = tk.StringVar(value="")
        for i, rel in enumerate(photos):
            path = APP_DIR / rel
            card = ttk.Frame(self.photo_frame, padding=5, relief="ridge")
            card.grid(row=i // 3, column=i % 3, padx=5, pady=5, sticky="n")
            try:
                with Image.open(path) as source:
                    img = ImageOps.exif_transpose(source).convert("RGB")
                    img.thumbnail((250, 175))
                tk_img = ImageTk.PhotoImage(img)
                self.preview_refs.append(tk_img)
                ttk.Label(card, image=tk_img).pack()
            except Exception:
                ttk.Label(card, text="사진을 열 수 없음", width=28).pack(pady=50)
            ttk.Radiobutton(card, text=path.name, variable=self.photo_selection, value=rel).pack(anchor="w")

    def save_current(self, silent=False):
        if self.data.get("items"):
            item = self.data["items"][self.current_index]
            for key, var in self.vars.items():
                item[key] = var.get().strip()
            self._save_data()
            if not silent:
                self._refresh_list(select=self.current_index)
        if not silent:
            messagebox.showinfo("저장", "품목 정보가 저장되었습니다.")

    def add_item(self):
        self.save_current(silent=True)
        self.data.setdefault("items", []).append({"kind": "Wedge", "model": "", "size": "", "serial": "", "quantity": "1", "note": "", "photos": []})
        self.current_index = len(self.data["items"]) - 1
        self._save_data()
        self._refresh_list(select=self.current_index)

    def delete_item(self):
        if not self.data.get("items"):
            return
        if messagebox.askyesno("품목 삭제", "선택한 품목과 사진 연결을 삭제할까요?\n복사된 사진 파일은 유지됩니다."):
            self.data["items"].pop(self.current_index)
            self.current_index = max(0, self.current_index - 1)
            self._save_data()
            self._refresh_list()

    def merge_selected_items(self):
        selected = sorted({int(iid) for iid in self.tree.selection()})
        if len(selected) < 2:
            messagebox.showinfo("품목 합치기", "합칠 품목을 Ctrl 키로 2개 이상 선택해 주세요.")
            return
        items = self.data.get("items", [])
        primary_index = next((i for i in selected if items[i].get("kind") not in {"미분류", "확인 필요"}), selected[0])
        primary = items[primary_index]
        for index in selected:
            if index == primary_index:
                continue
            other = items[index]
            for key in ("kind", "model", "size", "serial", "quantity", "note"):
                if not primary.get(key) and other.get(key):
                    primary[key] = other[key]
            for photo in other.get("photos", []):
                if photo not in primary.setdefault("photos", []):
                    primary["photos"].append(photo)
            if other.get("ocr_text"):
                primary["ocr_text"] = "\n".join(filter(None, [primary.get("ocr_text", ""), other["ocr_text"]]))
        for index in reversed(selected):
            if index != primary_index:
                items.pop(index)
        new_index = primary_index - sum(1 for index in selected if index < primary_index)
        self.current_index = new_index
        self._save_data()
        self._refresh_list(select=new_index)
        self.status_var.set(f"{len(selected)}개 품목을 하나로 합쳤습니다.")

    def _copy_photos(self, paths: list[str]) -> list[str]:
        copied = []
        for src_text in paths:
            src = Path(src_text)
            if src.suffix.lower() not in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}:
                continue
            dest = PHOTO_DIR / src.name
            n = 1
            while dest.exists() and dest.resolve() != src.resolve():
                dest = PHOTO_DIR / f"{src.stem}_{n}{src.suffix.lower()}"
                n += 1
            if not dest.exists():
                shutil.copy2(src, dest)
            copied.append(dest.relative_to(APP_DIR).as_posix())
        return copied

    def import_photos(self):
        folder = filedialog.askdirectory(title="사진 폴더 선택")
        if not folder:
            return
        paths = [str(p) for p in sorted(Path(folder).iterdir()) if p.is_file()]
        copied = self._copy_photos(paths)
        self.data.setdefault("items", []).append({"kind": "미분류", "model": "", "size": Path(folder).name, "serial": "", "quantity": "", "note": "가져온 사진 - 품목 정보를 입력하세요", "photos": copied})
        self.current_index = len(self.data["items"]) - 1
        self._save_data()
        self._refresh_list(select=self.current_index)
        messagebox.showinfo("가져오기 완료", f"사진 {len(copied)}장을 복사했습니다.\n이어서 사진 정보를 자동 분석합니다.")
        self.root.after(100, self.auto_classify_current)

    def auto_classify_current(self):
        if not self.data.get("items"):
            messagebox.showinfo("자동 인식", "분석할 사진이 없습니다.")
            return
        item = self.data["items"][self.current_index]
        photos = list(item.get("photos", []))
        if not photos:
            messagebox.showinfo("자동 인식", "선택한 품목에 사진이 없습니다.")
            return
        source_index = self.current_index
        self.status_var.set(f"OCR 준비 중... (총 {len(photos)}장)")

        def progress(done, total):
            self.root.after(0, lambda: self.status_var.set(f"사진 정보 인식 중... {done}/{total}"))

        def worker():
            try:
                from .ocr_service import analyze_and_group
                groups = analyze_and_group(APP_DIR, photos, progress)
                self.root.after(0, lambda: self._apply_ocr_groups(source_index, groups))
            except Exception as exc:
                self.root.after(0, lambda: self._ocr_failed(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_ocr_groups(self, source_index: int, groups: list[dict]):
        if source_index >= len(self.data.get("items", [])):
            self.status_var.set("분석 결과를 적용할 원본 항목을 찾지 못했습니다.")
            return
        self.data["items"][source_index:source_index + 1] = groups
        self.current_index = source_index
        self._save_data()
        self._refresh_list(select=source_index)
        uncertain = sum(1 for group in groups if group.get("kind") == "확인 필요")
        self.status_var.set(f"자동 분류 완료: {len(groups)}개 묶음, 확인 필요 {uncertain}개")
        messagebox.showinfo("자동 인식 완료", f"사진을 {len(groups)}개 항목으로 분류했습니다.\n인식 결과를 확인한 뒤 저장해 주세요.")

    def _ocr_failed(self, exc: Exception):
        self.status_var.set("자동 인식 실패")
        messagebox.showerror("자동 인식 오류", str(exc))

    def add_photos_to_item(self):
        if not self.data.get("items"):
            self.add_item()
        paths = filedialog.askopenfilenames(title="사진 선택", filetypes=[("사진", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff")])
        if paths:
            self.data["items"][self.current_index].setdefault("photos", []).extend(self._copy_photos(list(paths)))
            self._save_data()
            self._refresh_list(select=self.current_index)

    def remove_selected_photo(self):
        rel = getattr(self, "photo_selection", tk.StringVar()).get()
        if rel and self.data.get("items"):
            photos = self.data["items"][self.current_index].get("photos", [])
            if rel in photos:
                photos.remove(rel)
                self._save_data()
                self._show_item(self.current_index)

    def export_excel(self):
        self.save_current(silent=True)
        default = OUTPUT_DIR / f"PAUT_입고_사진대장_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        target = filedialog.asksaveasfilename(title="사진대장 저장", initialdir=OUTPUT_DIR, initialfile=default.name, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not target:
            return
        try:
            self._write_excel(Path(target))
            if messagebox.askyesno("출력 완료", "엑셀 사진대장을 만들었습니다.\n지금 열까요?"):
                self._open_path(Path(target))
        except Exception as exc:
            messagebox.showerror("출력 오류", str(exc))

    def _write_excel(self, target: Path):
        temp_dir = Path(tempfile.mkdtemp(prefix="paut_ledger_"))
        wb = Workbook()
        ws = wb.active
        ws.title = "입고사진대장"
        ws.sheet_view.showGridLines = False
        for col in range(1, 13):
            ws.column_dimensions[get_column_letter(col)].width = 11
        ws.merge_cells("A1:L2")
        ws["A1"] = "PAUT PROBE 및 WEDGE 입고 사진대장"
        ws["A1"].font = Font(name="맑은 고딕", size=18, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("A3:H3"); ws["A3"] = f"건명: {self.data.get('project', '')}"
        ws.merge_cells("I3:L3"); ws["I3"] = f"입고일: {self.data.get('received_date', '')}"
        blue = PatternFill("solid", fgColor="4F81BD")
        white = Font(name="맑은 고딕", color="FFFFFF", bold=True)
        thin = Side(style="thin", color="666666")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        row = 5
        items = self.data.get("items", [])
        for no, item in enumerate(items, 1):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
            title = f"{no}. {item.get('kind', '')} | {item.get('model', '')} | {item.get('size', '')} | 수량 {item.get('quantity', '')}"
            if item.get("serial"):
                title += f" | S/N {item['serial']}"
            ws.cell(row, 1, title)
            ws.cell(row, 1).fill = blue; ws.cell(row, 1).font = white; ws.cell(row, 1).alignment = Alignment(horizontal="center")
            for c in range(1, 13): ws.cell(row, c).border = border
            image_top = row + 1
            image_bottom = row + 9
            ws.merge_cells(start_row=image_top, start_column=1, end_row=image_bottom, end_column=6)
            ws.merge_cells(start_row=image_top, start_column=7, end_row=image_bottom, end_column=12)
            for c in range(1, 13):
                for r in range(image_top, image_bottom + 1): ws.cell(r, c).border = border
            photos = item.get("photos", [])[:2]
            for idx, rel in enumerate(photos):
                path = APP_DIR / rel
                if path.exists():
                    optimized = temp_dir / f"{no}_{idx}.jpg"
                    with Image.open(path) as source:
                        source = ImageOps.exif_transpose(source).convert("RGB")
                        source.thumbnail((1280, 900))
                        source.save(optimized, "JPEG", quality=82, optimize=True)
                    img = XLImage(str(optimized)); img.width = 360; img.height = 180
                    # Excel에서 각 사진 영역(A:F, G:L)은 약 528px이다.
                    # 360px 사진의 좌우 여백을 각각 84px로 맞춘다.
                    marker = AnchorMarker(
                        col=0 if idx == 0 else 6,
                        colOff=pixels_to_EMU(84),
                        row=image_top - 1,
                        rowOff=0,
                    )
                    img.anchor = OneCellAnchor(
                        _from=marker,
                        ext=XDRPositiveSize2D(cx=pixels_to_EMU(360), cy=pixels_to_EMU(180)),
                    )
                    ws.add_image(img)
            ws.merge_cells(start_row=row + 10, start_column=1, end_row=row + 10, end_column=12)
            ws.cell(row + 10, 1, item.get("note", ""))
            ws.cell(row + 10, 1).alignment = Alignment(horizontal="center")
            for c in range(1, 13): ws.cell(row + 10, c).border = border
            for r in range(image_top, image_bottom + 1): ws.row_dimensions[r].height = 16
            ws.row_dimensions[row].height = 24
            ws.row_dimensions[row + 10].height = 22
            if no < len(items) and no % 2 == 0:
                ws.row_breaks.append(Break(id=row + 10))
            row += 11
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.35, bottom=0.35, header=0.1, footer=0.1)
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        ws.print_title_rows = "1:3"
        ws.oddFooter.center.text = "Page &P / &N"
        ws.oddFooter.center.size = 9
        ws.oddFooter.center.font = "Arial"
        ws.print_area = f"A1:L{max(4, row - 2)}"
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            wb.save(target)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    @staticmethod
    def _open_path(path: Path):
        if sys.platform == "win32":
            subprocess.Popen(["explorer", str(path)])
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
