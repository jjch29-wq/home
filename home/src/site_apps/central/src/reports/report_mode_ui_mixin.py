# -*- coding: utf-8 -*-
import sys
import os
from pathlib import Path
import glob
import math
import traceback
import re
import warnings
import json
import tempfile
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import threading
import xlsxwriter
import datetime
from PIL import Image as PILImage, ImageChops, ImageOps
import io
import base64
import time
import copy





# --- Versioning ---
APP_VERSION = "v260404.03"
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.pagebreak import Break
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple, column_index_from_string
# Ignore library warnings
warnings.simplefilter("ignore")

# --- Constants & Paths ---
NAN_PATTERN = re.compile(r'^nan(\.0+)?$|^none$|^null$|^0\.0+|-0\.0+$', re.IGNORECASE)
DOT_ZERO_PATTERN = re.compile(r'\.0$')


def find_exact_column(df, keywords):
    """Return a column only when its header exactly matches a keyword.

    PT source files often contain headers such as ``Spool No`` and
    ``Request No`` even when they do not have an independent sequence-number
    column.  A substring match for ``NO`` therefore maps unrelated identifiers
    into the preview's No field.
    """
    normalized_keywords = {
        re.sub(r'\s+', '', str(keyword)).upper() for keyword in keywords
    }
    for column in df.columns:
        normalized_column = re.sub(r'\s+', '', str(column)).upper()
        for keyword in normalized_keywords:
            # Multi-row Excel headers can repeat the same label when the
            # header rows are combined (for example,
            # "Request Date Request Date").  Treat only repetitions of the
            # complete keyword as exact; unrelated names such as Spool No
            # must still not match NO.
            if normalized_column == keyword:
                return column
            if keyword and len(normalized_column) > len(keyword):
                repeat_count, remainder = divmod(len(normalized_column), len(keyword))
                if remainder == 0 and repeat_count > 1 and normalized_column == keyword * repeat_count:
                    return column
    return None


def format_thickness_range(items):
    """Format the minimum and maximum numeric PT thickness values."""
    thicknesses = []
    for item in items:
        raw_value = item.get("Thk.", "") if isinstance(item, dict) else ""
        if isinstance(raw_value, (int, float)) and not pd.isna(raw_value):
            thicknesses.append(float(raw_value))
            continue
        match = re.fullmatch(
            r'\s*(\d+(?:\.\d+)?)\s*(?:mm|t)?\s*',
            str(raw_value),
            re.IGNORECASE,
        )
        if match:
            thicknesses.append(float(match.group(1)))

    if not thicknesses:
        return ""

    def _display(value):
        return f"{value:.10f}".rstrip("0").rstrip(".")

    return f"{_display(min(thicknesses))} ~ {_display(max(thicknesses))}mm"


def first_report_date(items):
    """Return the first valid extracted report date as a date value."""
    for item in items:
        raw_value = item.get("Date", "") if isinstance(item, dict) else ""
        if isinstance(raw_value, pd.Timestamp):
            return raw_value.to_pydatetime()
        if isinstance(raw_value, datetime.datetime):
            return raw_value
        if isinstance(raw_value, datetime.date):
            return datetime.datetime.combine(raw_value, datetime.time.min)

        text = str(raw_value).strip()
        for date_format in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
            try:
                return datetime.datetime.strptime(text[:10], date_format)
            except ValueError:
                continue
    return None


def remove_unused_report_sheets(workbook, used_sheets):
    """Remove template sheets that were not used by the generated report."""
    used_ids = {id(sheet) for sheet in used_sheets}
    for sheet in list(workbook.worksheets):
        if id(sheet) not in used_ids:
            workbook.remove(sheet)

# [PT] SCH -> 두께(mm) 변환 테이블
SCH_TO_THK = {
    "1/2": {"5S": 1.65, "10S": 2.11, "40": 2.77, "80": 3.73, "160": 4.75, "XXS": 7.47},
    "3/4": {"5S": 1.65, "10S": 2.11, "40": 2.87, "80": 3.91, "160": 5.56, "XXS": 7.82},
    "1": {"5S": 1.65, "10S": 2.77, "40": 3.38, "80": 4.55, "160": 6.35, "XXS": 9.09},
    "1-1/4": {"5S": 1.65, "10S": 2.77, "40": 3.56, "80": 4.85, "160": 6.35, "XXS": 9.70},
    "1-1/2": {"5S": 1.65, "10S": 2.77, "40": 3.68, "80": 5.08, "160": 7.14, "XXS": 10.16},
    "2": {"5S": 1.65, "10S": 2.77, "40": 3.91, "80": 5.54, "160": 8.74, "XXS": 11.07},
    "2-1/2": {"5S": 2.11, "10S": 3.05, "40": 5.16, "80": 7.01, "160": 9.53, "XXS": 14.02},
    "3": {"5S": 2.11, "10S": 3.05, "40": 5.49, "80": 7.62, "160": 11.13, "XXS": 15.24},
    "4": {"5S": 2.11, "10S": 3.05, "40": 6.02, "80": 8.56, "160": 13.49, "XXS": 17.12},
    "5": {"5S": 2.77, "10S": 3.40, "40": 6.55, "80": 9.53, "120": 12.70, "160": 15.88},
    "6": {"5S": 2.77, "10S": 3.40, "40": 7.11, "80": 10.97, "120": 14.27, "160": 18.26, "XXS": 21.95},
    "8": {"5S": 2.77, "10S": 3.76, "20": 6.35, "30": 7.04, "40": 8.18, "60": 10.31, "80": 12.70, "100": 15.09, "120": 18.26, "140": 20.62, "160": 23.01, "XXS": 22.23},
    "10": {"5S": 3.40, "10S": 4.19, "20": 6.35, "30": 7.80, "40": 9.27, "60": 12.70, "80": 15.09, "100": 18.26, "120": 21.44, "140": 25.40, "160": 28.58},
    "12": {"5S": 3.96, "10S": 4.57, "20": 6.35, "30": 8.38, "40": 10.31, "60": 14.27, "80": 17.48, "100": 21.44, "120": 25.40, "140": 28.58, "160": 33.32},
    "14": {"5S": 3.96, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 9.53, "40": 11.13, "60": 15.09, "80": 19.05, "100": 23.83, "120": 27.79, "140": 31.75, "160": 35.71},
    "16": {"5S": 4.19, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 9.53, "40": 12.70, "60": 16.66, "80": 21.44, "100": 26.19, "120": 30.96, "140": 36.53, "160": 40.49},
    "18": {"5S": 4.19, "10S": 4.78, "10": 6.35, "20": 7.92, "30": 11.13, "40": 14.27, "60": 19.05, "80": 23.83, "100": 29.36, "120": 34.93, "140": 39.67, "160": 45.24},
    "20": {"5S": 4.78, "10S": 5.54, "10": 6.35, "20": 9.53, "30": 12.70, "40": 15.09, "60": 20.62, "80": 26.19, "100": 32.54, "120": 38.10, "140": 44.45, "160": 50.01},
    "24": {"5S": 5.54, "10S": 6.35, "10": 6.35, "20": 9.53, "30": 14.27, "40": 17.48, "60": 24.61, "80": 30.96, "100": 38.89, "120": 46.02, "140": 52.37, "160": 59.54},
}

def convert_sch_to_thk(size_val, thk_val):
    """SCH 값을 두께(mm)로 변환"""
    if pd.isna(thk_val) or str(thk_val).strip() == "": return ""
    thk_str = str(thk_val).strip().upper()
    try:
        val = float(thk_str.replace("MM", "").replace("T", "").strip())
        if 0 < val < 100: return f"{val:.2f}"
    except: pass
    sch_match = re.search(r'(?:SCH[.\s]?|S/)?(\d+S?|XXS|XS)', thk_str, re.IGNORECASE)
    if not sch_match: return thk_str
    sch = sch_match.group(1).upper()
    if sch.endswith('S') and sch not in ['5S', '10S', 'XXS', 'XS']: sch = sch[:-1]
    if pd.isna(size_val) or str(size_val).strip() == "": return thk_str
    size_str = str(size_val).strip().replace('"', '').replace("'", "")
    size_str = re.sub(r'\s+', '-', size_str)
    if size_str in SCH_TO_THK and sch in SCH_TO_THK[size_str]:
        return f"{SCH_TO_THK[size_str][sch]:.2f}"
    try:
        size_int = str(int(float(size_str)))
        if size_int in SCH_TO_THK and sch in SCH_TO_THK[size_int]:
            return f"{SCH_TO_THK[size_int][sch]:.2f}"
    except: pass
    return thk_str

# 공통 스타일 (테두리 등)
thin_side = Side(style='thin')
medium_side = Side(style='medium')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

def find_project_home(start_path):
    current = os.path.abspath(start_path)
    while True:
        if os.path.isdir(os.path.join(current, "config")) and os.path.isdir(os.path.join(current, "resources")):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            return os.path.abspath(start_path)
        current = parent


if getattr(sys, 'frozen', False):
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    SCRIPT_HOME = base_path
    BASE_DIR = base_path
    CENTRAL_HOME = base_path
    CONFIG_DIR = base_path
    RESOURCE_DIR = os.path.join(base_path, "resources")
else:
    SCRIPT_HOME = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = find_project_home(SCRIPT_HOME)
    CENTRAL_HOME = os.path.abspath(os.path.join(SCRIPT_HOME, "..", ".."))
    CONFIG_DIR = os.path.join(CENTRAL_HOME, "config")
    _central_res = os.path.join(CENTRAL_HOME, "resources")
    _base_res = os.path.join(BASE_DIR, "resources")
    RESOURCE_DIR = _central_res if os.path.exists(os.path.join(_central_res, "logo.png")) else _base_res

SETTINGS_FILE = os.path.join(CONFIG_DIR, "logo_settings_unified.json")
LEGACY_SETTINGS_FILE = os.path.join(BASE_DIR, "config", "logo_settings_unified.json")
PHOTOLOG_CONFIG_FILE = os.path.join(CONFIG_DIR, "photolog_config.json")
SRC_ROOT = str(Path(__file__).resolve().parents[4]) if 'Path' in globals() else os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))



# Embedded Logo Fallback
DEFAULT_SITCO_LOGO_B64 = "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"



class ReportModeUIMixin:
    def _create_scrollable_sidebar(self, parent):
        """Creates a scrollable canvas/scrollbar container for sidebars."""
        canvas = tk.Canvas(parent, background="#f9fafb", highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, background="#f9fafb", padx=10, pady=0, highlightthickness=0, borderwidth=0)

        def _update_scrollregion(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Auto-hide scrollbar if content fits
            if canvas.bbox("all")[3] <= canvas.winfo_height():
                scrollbar.pack_forget()
            else:
                scrollbar.pack(side="right", fill="y")
        
        scrollable_frame.bind("<Configure>", _update_scrollregion)

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_window, width=e.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _on_mousewheel(event):
            # Only allow scrolling if content actually exceeds the visible area
            if canvas.bbox("all")[3] > canvas.winfo_height():
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"
        
        canvas.bind('<Enter>', lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        return scrollable_frame

    def _setup_pmi_ui(self, parent):
        # [FORCE] Ensure parent (pmi_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.pmi_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.pmi_paned.pack(fill='both', expand=True)

        # [LEFT] Filter/Settings Sidebar
        left_container = tk.Frame(self.pmi_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.pmi_paned.add(left_container, width=425)
        self._add_sidebar_toggle(self.pmi_paned, left_container)

        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header (Lifts over scrollable area)
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PMI 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        tk.Label(inner_header, text=f"v{APP_VERSION}", font=("Arial", 8), 
                 background="#f1f5f9", foreground="#64748b", padx=5, pady=0).pack(side='left', padx=10)
        
        # Bottom separator
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift() # Ensure it's on top

        # Spacer in scrollable area so first group starts below floating header
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_container = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_container.pack(fill='x', pady=(0, 10))

        def _add_very_compact_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            # Use width=1 for entry to allow proactive shrinking
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda : self._browse_dir(var)) if is_dir else (lambda : self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_very_compact_row(file_container, "데이터:", self.target_file_path, 0, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_very_compact_row(file_container, "양식:", self.template_file_path, 1, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_container, background="#f0f0f0")
        btn_frame.grid(row=2, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("PMI")).pack(side='left', padx=2)

        # [NEW] PMI 양식 모드 선택 (일반 1행 vs 모재+용접부+모재 3단)
        layout_mode_frame = tk.LabelFrame(file_container, text=" 양식 모드 (측정 방식) ", background="#f0f0f0", padx=4, pady=2)
        layout_mode_frame.grid(row=3, column=0, columnspan=3, sticky='ew', padx=2, pady=3)
        ttk.Radiobutton(layout_mode_frame, text="일반 단일행 (1포인트 = 1행)", variable=self.pmi_layout_mode, value="STANDARD").pack(side='left', padx=6)
        ttk.Radiobutton(layout_mode_frame, text="플랜트 배관용 (모재+용접부+모재 3단 / 1p당 9포인트)", variable=self.pmi_layout_mode, value="PLANT_TRIPLE").pack(side='left', padx=6)

        # 2. Configuration Notebook (Cover, Data, Rows - NO Preview tab here)
        config_frame = ttk.LabelFrame(left_pane, text=" 설정 (Config) ", padding=2)
        config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.pmi_tab_notebook = ttk.Notebook(config_frame)
        self.pmi_tab_notebook.pack(fill='both', expand=True)

        tab_cover = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_data = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_rows = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_cols = ttk.Frame(self.pmi_tab_notebook, padding=5)
        tab_filter = ttk.Frame(self.pmi_tab_notebook, padding=10)  # 내부 호환용 위젯 유지 (노트북 탭 미노출)
        
        # [CRITICAL] Allow children to expand within tabs
        for t in [tab_cover, tab_data, tab_rows, tab_cols]: t.columnconfigure(0, weight=1)

        self.pmi_tab_notebook.add(tab_cover, text="갑지")
        self.pmi_tab_notebook.add(tab_data, text="을지")
        self.pmi_tab_notebook.add(tab_rows, text="행 설정")
        self.pmi_tab_notebook.add(tab_cols, text="컬럼 설정")
        # [CLEANUP] 사용 빈도가 낮고 혼선을 주는 레거시 '필터/옵션' 탭은 화면에서 숨김 (재질 자동 판정은 기본 True 유지)
        # self.pmi_tab_notebook.add(tab_filter, text="필터/옵션")
        if hasattr(self, 'auto_verify'):
            self.auto_verify.set(True)

        # [FIX] Do NOT reset setting_vars here to prevent data loss when switching tabs
        # [ALIGNED] Mode-specific context for logo grid
        self._create_gapji_meta_ui(tab_cover, use_pack=False, mode="PMI")
        self.pmi_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        # PMI logos are placed automatically from the saved configuration/template.
        # Keep the settings and insertion engine, but do not expose manual logo
        # path/position controls in the PMI tabs.
        self._create_margin_settings(tab_cover, "PMI_COVER", use_pack=False)
        self._create_margin_settings(tab_data, "PMI_DATA", use_pack=False)
        self._create_row_settings(tab_rows, mode="PMI")
        
        pmi_items = [
            ("No:", "PMI_COL_NO", 13, "PMI_NAME_NO", "No", "No"),
            ("Drawing No:", "PMI_COL_DWG", 1, "PMI_NAME_DWG", "Drawing No.", "Dwg"),
            ("Joint No:", "PMI_COL_JOINT", 6, "PMI_NAME_JOINT", "Joint No.", "Joint"),
            ("Location:", "PMI_COL_LOC", 7, "PMI_NAME_LOC", "Location", "Loc"),
            ("Ni:", "PMI_COL_NI", 8, "PMI_NAME_NI", "Ni", "Ni"),
            ("Cr:", "PMI_COL_CR", 9, "PMI_NAME_CR", "Cr", "Cr"),
            ("Mo:", "PMI_COL_MO", 10, "PMI_NAME_MO", "Mo", "Mo"),
            ("판정 결과:", "PMI_COL_RES", 13, "PMI_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(tab_cols, "PMI", pmi_items)
        
        # Spacer for tabs
        tk.Frame(tab_rows, background="#f9fafb").pack(fill='x')

        # 3. Quick Options Section (Moved to Settings Notebook)
        # Filter Tag Area (Self-updating)
        filter_box = tk.Frame(tab_filter, background="#f9fafb")
        filter_box.pack(fill='x', pady=2)
        ttk.Label(filter_box, text="🔍 성분 필터:", background="#f9fafb", font=("Malgun Gothic", 9, "bold")).pack(side='left', padx=(0, 5))
        
        def add_filter():
            self.element_filters.append({'key': tk.StringVar(value=''), 'min': tk.StringVar(value=''), 'max': tk.StringVar(value='')})
            self._update_pmi_filter_ui()
        ttk.Button(filter_box, text="+", width=2, command=add_filter).pack(side='right')

        self.filter_container = tk.Frame(filter_box, background="#f9fafb")
        self.filter_container.pack(side='left', fill='x', expand=True)

        # Sequence Filter & Auto Check
        tk.Frame(tab_filter, height=1, background="#e5e7eb").pack(fill='x', pady=10)
        tk.Checkbutton(tab_filter, text="✅ 재질 자동 판정", variable=self.auto_verify, background="#f9fafb", font=("Malgun Gothic", 8)).pack(anchor='w', pady=(0, 5))
        seq_row = tk.Frame(tab_filter, background="#f9fafb")
        seq_row.pack(fill='x', pady=2)
        ttk.Label(seq_row, text="📊 특정순번:", background="#f9fafb", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(seq_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        self._update_pmi_filter_ui()

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.pmi_paned, background="#f3f4f6")
        self.pmi_paned.add(right_container, stretch="always")

        # Top Action Bar (생성 시작, 데이터 추출, 일보 NDT 불러오기) - 모던 플랫 툴바 스타일
        action_bar = tk.Frame(right_container, background="#ffffff", padx=8, pady=6)
        action_bar.pack(fill='x')

        btn_run = tk.Button(
            action_bar, text="✨ 생성 시작", command=self.run_process,
            font=("Malgun Gothic", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=5, bd=0
        )
        btn_run.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_run.bind("<Enter>", lambda e: btn_run.configure(bg="#1d4ed8"))
        btn_run.bind("<Leave>", lambda e: btn_run.configure(bg="#2563eb"))

        btn_extract = tk.Button(
            action_bar, text="📝 데이터 추출", command=self.extract_only,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_extract.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_extract.bind("<Enter>", lambda e: btn_extract.configure(bg="#e2e8f0"))
        btn_extract.bind("<Leave>", lambda e: btn_extract.configure(bg="#f8fafc"))

        btn_daily = tk.Button(
            action_bar, text="📋 일보 NDT 불러오기", command=self.load_daily_work_history,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_daily.pack(side='left', fill='x', expand=True)
        btn_daily.bind("<Enter>", lambda e: btn_daily.configure(bg="#e2e8f0"))
        btn_daily.bind("<Leave>", lambda e: btn_daily.configure(bg="#f8fafc"))

        # 툴바와 데이터 화면 사이 구분선
        tk.Frame(right_container, height=1, background="#e2e8f0").pack(fill='x', pady=(0, 4))
        
        # [DATA ONLY] 갑지 탭 제거 -> 100% 데이터 미리보기 전용 화면으로 쾌적화
        tab_data = tk.Frame(right_container, background="#ffffff")
        tab_data.pack(fill='both', expand=True)
        
        self._create_preview_ui(tab_data)
        self._apply_sash_ratio("PMI")
        
        self.template_file_path.trace_add("write", lambda *a: self.load_template_specific_config("PMI"))

    def _update_pmi_ratio(self):
        """Saves current sash position and persists it."""
        try:
            total_w = self.pmi_paned.winfo_width()
            if total_w > 100:
                current_sash = self.pmi_paned.sash_coord(0)[0]
                self.pmi_pane_ratio = current_sash / total_w
                self.config['PMI_SASH_RATIO'] = self.pmi_pane_ratio
                self.save_settings() # Persist immediately
        except: pass

    def _on_pmi_paned_configure(self, event):
        """Maintains the proportional split when the window resizes."""
        try:
            total_w = self.pmi_paned.winfo_width()
            if total_w > 100:
                # During window resize, RE-PLACE sash based on SAVED ratio
                if event and event.widget == self.pmi_paned:
                    new_pos = int(total_w * self.pmi_pane_ratio)
                    self.pmi_paned.sash_place(0, new_pos, 0)
        except: pass

    def _update_pmi_filter_ui(self):
        """Redraws the element filter tags in a modern, elastic style."""
        for widget in self.filter_container.winfo_children():
            widget.destroy()
        
        def remove_filter(idx):
            self.element_filters.pop(idx)
            self._update_pmi_filter_ui()

        # [HYPER-REACTIVE] Use grid for even horizontal distribution (Elastic, Stacked)
        cols_per_row = 1
        for i, f_item in enumerate(self.element_filters):
            self.filter_container.columnconfigure(0, weight=1)
            
            # Create an "Elastic Tag" 
            tag = tk.Frame(self.filter_container, background="#eff6ff", highlightthickness=1, 
                            highlightbackground="#bfdbfe", padx=2, pady=1)
            tag.grid(row=i, column=0, sticky='ew', padx=2, pady=2)
            
            # Inner grid for tag components
            tag.columnconfigure(0, weight=1, minsize=0) # Key
            tag.columnconfigure(2, weight=1, minsize=0) # Min
            tag.columnconfigure(4, weight=1, minsize=0) # Max
            
            # Key
            e_key = tk.Entry(tag, textvariable=f_item['key'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_key.grid(row=0, column=0, sticky='ew')
            tk.Label(tag, text=":", background="#eff6ff", font=("Arial", 8)).grid(row=0, column=1)
            
            # Min
            e_min = tk.Entry(tag, textvariable=f_item['min'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_min.grid(row=0, column=2, sticky='ew')
            tk.Label(tag, text="~", background="#eff6ff", font=("Arial", 8)).grid(row=0, column=3)
            
            # Max
            e_max = tk.Entry(tag, textvariable=f_item['max'], width=1, font=("Arial", 8), 
                             relief='flat', background="#eff6ff", justify='center')
            e_max.grid(row=0, column=4, sticky='ew')
            
            # Delete button (Small)
            btn_del = tk.Button(tag, text="×", command=lambda idx=i: remove_filter(idx), 
                                relief='flat', background="#eff6ff", foreground="#3b82f6", font=("Arial", 9, "bold"),
                                activebackground="#dbeafe", cursor="hand2", padx=0, pady=0)
            btn_del.grid(row=0, column=5, padx=(2, 0))

    def _setup_rt_ui(self, parent):
        """[RE-INTEGRATED] RT 탭 통합 구성 (내부에 서브 미리보기 탭 배치)"""
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.rt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.rt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.rt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.rt_paned.add(left_container, width=425)
        self._add_sidebar_toggle(self.rt_paned, left_container)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED Header
        header_frame = tk.Frame(left_container, background="#f9fafb")
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        tk.Label(header_frame, text="🔬 RT 성적서 통합 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left', padx=20, pady=5)
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # [NEW] 모드 선택 (표준 vs 가스공사)
        mode_select_f = tk.Frame(left_pane, background="#f9fafb")
        mode_select_f.pack(fill='x', pady=(0, 10))
        self.rt_sub_mode = tk.StringVar(value="RT")
        
        style = ttk.Style()
        style.configure("Mode.TRadiobutton", font=("Malgun Gothic", 9, "bold"))
        
        ttk.Radiobutton(mode_select_f, text=" 🔬 표준 RT ", variable=self.rt_sub_mode, value="RT", 
                        style="Mode.TRadiobutton", command=self._on_rt_sub_mode_change).pack(side='left', padx=5)
        ttk.Radiobutton(mode_select_f, text=" 🇰🇷 가스공사 ", variable=self.rt_sub_mode, value="KOGAS", 
                        style="Mode.TRadiobutton", command=self._on_rt_sub_mode_change).pack(side='left', padx=5)

        # 1. File Selection (Adaptive)
        self.rt_file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        self.rt_file_frame.pack(fill='x', pady=(0, 10))
        
        # We need to show different path entries based on mode
        self.rt_path_container = tk.Frame(self.rt_file_frame)
        self.rt_path_container.pack(fill='x')
        
        def _add_file_row(p, lbl, var, r, is_dir=False, types=None):
            p.columnconfigure(1, weight=1)
            ttk.Label(p, text=lbl, font=("Malgun Gothic", 8)).grid(row=r, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(p, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=r, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(p, text="...", width=3, command=cmd).grid(row=r, column=2, padx=2, pady=2)

        # Container for actual path rows (will be cleared/re-added on mode change)
        self._update_rt_path_ui()

        # 2. Tabs for Configuration
        rt_config_frame = ttk.LabelFrame(left_pane, text=" 세부 설정 (Config) ", padding=2)
        rt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.rt_tab_nb = ttk.Notebook(rt_config_frame)
        self.rt_tab_nb.pack(fill='both', expand=True)

        self.rt_tabs = {}
        for t_name in ["갑지", "을지", "로고", "행 설정", "컬럼 설정"]:
            f = ttk.Frame(self.rt_tab_nb, padding=5)
            self.rt_tab_nb.add(f, text=t_name)
            self.rt_tabs[t_name] = f
            if t_name not in ["행 설정", "컬럼 설정"]: f.columnconfigure(0, weight=1)

        self.rt_tab_nb.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        self._refresh_rt_config_tabs() # Initial UI build

        # 3. Actions
        action_f = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_f.pack(fill='x', pady=(0, 10))
        
        # Row 1: Sequence Filter
        r1 = tk.Frame(action_f, background="#ffffff")
        r1.pack(fill='x', pady=2)
        ttk.Label(r1, text="📊 특정순번 필터:", background="#ffffff").pack(side='left')
        ttk.Entry(r1, textvariable=self.sequence_filter, width=15).pack(side='left', padx=5, fill='x', expand=True)
        
        # Row 2: Keyword Filter
        r2 = tk.Frame(action_f, background="#ffffff")
        r2.pack(fill='x', pady=2)
        ttk.Label(r2, text="🔍 추출 키워드:", background="#ffffff").pack(side='left')
        ttk.Entry(r2, textvariable=self.rt_extract_keyword, width=15).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview (Sub-tabs)
        right_f = tk.Frame(self.rt_paned, background="#f3f4f6")
        self.rt_paned.add(right_f, stretch="always")
        
        self._create_top_action_bar(right_f, "RT")
        self.rt_preview_nb = ttk.Notebook(right_f)
        self.rt_preview_nb.pack(fill='both', expand=True)
        
        t_std = tk.Frame(self.rt_preview_nb, background="#ffffff")
        t_kogas = tk.Frame(self.rt_preview_nb, background="#ffffff")
        t_gapji = tk.Frame(self.rt_preview_nb, background="#ffffff")
        
        self.rt_preview_nb.add(t_std, text=" 🔬 표준 미리보기 ")
        self.rt_preview_nb.add(t_kogas, text=" 🇰🇷 가스공사 미리보기 ")
        self.rt_preview_nb.add(t_gapji, text=" 📄 갑지 미리보기 ")
        
        # Link Preview Tab to Mode Switch
        def _on_rt_preview_tab_changed(event):
            tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
            new_mode = None
            if "가스공사" in tab_text: new_mode = "KOGAS"
            elif "표준" in tab_text: new_mode = "RT"
            if new_mode and new_mode != self.rt_sub_mode.get():
                self.rt_sub_mode.set(new_mode)
                self._on_rt_sub_mode_change(from_tab=True)
        self.rt_preview_nb.bind("<<NotebookTabChanged>>", _on_rt_preview_tab_changed)

        self._create_rt_preview_ui(t_std, mode="RT")
        self._create_rt_preview_ui(t_kogas, mode="KOGAS")
        self._create_gapji_preview_ui(t_gapji, "RT") # Gapji uses current sub-mode context internally
        
        self._apply_sash_ratio("RT")
        self.rt_pane_ratio = self.config.get('RT_SASH_RATIO', 0.5)
        self.rt_paned.bind("<Configure>", self._on_rt_paned_configure)
        self.rt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_rt_ratio))
        self.root.after(500, lambda: self._on_rt_paned_configure(None))

    def _update_rt_path_ui(self):
        """모드에 따라 파일 선택 경로 입력 UI 갱신"""
        for w in self.rt_path_container.winfo_children(): w.destroy()
        mode = self.rt_sub_mode.get()
        v_target = self.rt_target_file_path if mode == "RT" else self.kogas_target_file_path
        v_template = self.rt_template_file_path if mode == "RT" else self.kogas_template_file_path

        def _add_file_row(p, lbl, var, r, is_dir=False, types=None):
            p.columnconfigure(1, weight=1)
            ttk.Label(p, text=lbl, font=("Malgun Gothic", 8)).grid(row=r, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(p, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=r, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(p, text="...", width=3, command=cmd).grid(row=r, column=2, padx=2, pady=2)

        _add_file_row(self.rt_path_container, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(self.rt_path_container, f"{mode} 데이터:", v_target, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(self.rt_path_container, f"{mode} 양식:", v_template, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        btn_f = tk.Frame(self.rt_path_container, background="#f0f0f0")
        btn_f.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_f, text="💾 전용 설정 저장", command=lambda: self.save_template_specific_config(mode)).pack(side='left', padx=2)

    def _refresh_rt_config_tabs(self):
        """모드에 따라 행/컬럼 설정 탭 내용 갱신"""
        mode = self.rt_sub_mode.get()
        t_cover = self.rt_tabs["갑지"]
        t_data = self.rt_tabs["을지"]
        t_logo = self.rt_tabs["로고"]
        t_rows = self.rt_tabs["행 설정"]
        t_cols = self.rt_tabs["컬럼 설정"]
        
        for t in [t_rows, t_cols, t_logo, t_cover, t_data]:
            for w in t.winfo_children(): w.destroy()
        
        # Meta & Margins (Cover/Data)
        self._create_gapji_meta_ui(t_cover, use_pack=False, mode="RT")
        ctx_cover = f"{mode}_COVER"
        ctx_data = f"{mode}_DATA"
        self._create_margin_settings(t_cover, ctx_cover, use_pack=False)
        self._create_margin_settings(t_data, ctx_data, use_pack=False)
        
        # Logos
        tk.Label(t_logo, text="🚩 [갑지] 로고", font=("Malgun Gothic", 9, "bold"), background="#f9fafb").grid(row=0, column=0, sticky='w')
        r_next = self._create_setting_grid(t_logo, ctx_cover)
        tk.Label(t_logo, text="🚩 [을지] 로고", font=("Malgun Gothic", 9, "bold"), background="#f9fafb").grid(row=r_next+1, column=0, sticky='w', pady=(10, 0))
        self._create_setting_grid(t_logo, ctx_data)
        
        # Row Settings
        self._create_row_settings(t_rows, mode=mode)
        
        # Column Mapping (Ordered by current preview keys)
        all_possible = {
            "No": ("No:", f"{mode}_COL_NO", 1, f"{mode}_NAME_NO", "No", "No"),
            "Date": ("Date:", f"{mode}_COL_DATE", 2, f"{mode}_NAME_DATE", "Date", "Date"),
            "Sec": ("SEC:", f"{mode}_COL_SEC", 0, f"{mode}_NAME_SEC", "SEC", "Sec"),
            "Dwg": ("Dwg No.:", f"{mode}_COL_DWG", 3, f"{mode}_NAME_DWG", "Drawing No.", "Dwg"),
            "Joint": ("Film Ident. No.:", f"{mode}_COL_JOINT", 4, f"{mode}_NAME_JOINT", "Film Ident. No.", "Joint"),
            "Loc": ("Film Location:", f"{mode}_COL_LOC", 5, f"{mode}_NAME_LOC", "Film Location", "Loc"),
            "T": ("T:", f"{mode}_COL_THK", 6, f"{mode}_NAME_THK", "T", "T"),
            "Mat": ("Mat:", f"{mode}_COL_MAT", 7, f"{mode}_NAME_MAT", "Mat", "Mat"),
            "Size": ("구경(Size):", f"{mode}_COL_SIZE", 0, f"{mode}_NAME_SIZE", "Size", "Size"),
            "Acc": ("Acc:", f"{mode}_COL_ACC", 8, f"{mode}_NAME_ACC", "Acc", "Acc"),
            "Rej": ("Rej:", f"{mode}_COL_REJ", 9, f"{mode}_NAME_REJ", "Rej", "Rej"),
            "Deg": ("Deg:", f"{mode}_COL_DEG", 10, f"{mode}_NAME_DEG", "Deg", "Deg"),
            "D1": ("① Crack:", f"{mode}_COL_D1", 13, f"{mode}_NAME_D1", "① Crack", "D1"),
            "D2": ("② IP:", f"{mode}_COL_D2", 14, f"{mode}_NAME_D2", "② IP", "D2"),
            "D3": ("③ LF:", f"{mode}_COL_D3", 15, f"{mode}_NAME_D3", "③ LF", "D3"),
            "D4": ("④ Slag:", f"{mode}_COL_D4", 16, f"{mode}_NAME_D4", "④ Slag", "D4"),
            "D5": ("⑤ Por:", f"{mode}_COL_D5", 17, f"{mode}_NAME_D5", "⑤ Por", "D5"),
            "⑥ U/C": ("⑥ U/C:", f"{mode}_COL_D6", 18, f"{mode}_NAME_D6", "⑥ U/C", "D6"),
            "D6": ("⑥ U/C:", f"{mode}_COL_D6", 18, f"{mode}_NAME_D6", "⑥ U/C", "D6"),
            "D7": ("⑦ RUC:", f"{mode}_COL_D7", 19, f"{mode}_NAME_D7", "⑦ RUC", "D7"),
            "D8": ("⑧ BT:", f"{mode}_COL_D8", 20, f"{mode}_NAME_D8", "⑧ BT", "D8"),
            "D9": ("⑨ TI:", f"{mode}_COL_D9", 21, f"{mode}_NAME_D9", "⑨ TI", "D9"),
            "D10": ("⑩ CP:", f"{mode}_COL_D10", 22, f"{mode}_NAME_D10", "⑩ CP", "D10"),
            "D11": ("⑪ RC:", f"{mode}_COL_D11", 23, f"{mode}_NAME_D11", "⑪ RC", "D11"),
            "D12": ("⑫ Mis:", f"{mode}_COL_D12", 24, f"{mode}_NAME_D12", "⑫ Mis", "D12"),
            "D13": ("⑬ EP:", f"{mode}_COL_D13", 25, f"{mode}_NAME_D13", "⑬ EP", "D13"),
            "D14": ("⑭ SD:", f"{mode}_COL_D14", 26, f"{mode}_NAME_D14", "⑭ SD", "D14"),
            "D15": ("⑮ Oth:", f"{mode}_COL_D15", 27, f"{mode}_NAME_D15", "⑮ Oth", "D15"),
            "Result": ("판정(Result):", f"{mode}_COL_RES", 28, f"{mode}_NAME_RES", "Result", "Result"),
            "Welder": ("용접사:", f"{mode}_COL_WELDER", 29, f"{mode}_NAME_WELDER", "Welder No", "Welder"),
            "Remarks": ("비고:", f"{mode}_COL_REM", 30, f"{mode}_NAME_REM", "Remarks", "Remarks")
        }
        
        keys_attr = "rt_column_keys" if mode == "RT" else "kogas_column_keys"
        current_keys = getattr(self, keys_attr)
        
        # 1. Start with items that are currently visible/ordered
        rt_items = []
        for k in current_keys:
            if k in all_possible:
                rt_items.append(all_possible[k])
        
        # 2. Add remaining items at the end (so they remain configurable even if hidden)
        for k, v in all_possible.items():
            if v not in rt_items:
                rt_items.append(v)

        if mode == "KOGAS":
            self._create_kogas_separated_column_mapping_ui(t_cols)
        else:
            self._create_column_mapping_ui(t_cols, mode, rt_items)

    def _on_rt_sub_mode_change(self, from_tab=False):
        """RT 내부 서브 모드(표준/가스공사) 변경 시 UI 및 상태 동기화"""
        try:
            # [NEW] 서브 모드 전환 전 기존 설정 내용을 안전하게 자동 저장
            self.save_settings()
        except: pass
        mode = self.rt_sub_mode.get()
        self.current_mode = mode
        self._load_report_info_for_mode(mode)
        self._sub_mode_switching = True  # guard: prevent _on_rt_kogas_mode_change from overwriting paths
        self.rt_kogas_mode.set(mode == "KOGAS")
        self._sub_mode_switching = False
        self.log(f"🔄 RT 서브 모드 전환: {mode}")
        
        # 1. 오른쪽 미리보기 탭 동기화 (라디오 버튼으로 바꾼 경우)
        if not from_tab:
            target_idx = 0 if mode == "RT" else 1
            if self.rt_preview_nb.index("current") != target_idx:
                self.rt_preview_nb.select(target_idx)
        
        # 2. 왼쪽 사이드바 UI 갱신 (파일 경로 및 세부 설정 탭)
        self._update_rt_path_ui()
        self._refresh_rt_config_tabs()
        
        # 3. 템플릿 로드 트리거 [FIX] Pass correct path and avoid forced tab switch
        t_path = self.rt_template_file_path.get() if mode == "RT" else self.kogas_template_file_path.get()
        self.load_template_specific_config(t_path, mode)
        
        # 4. 갑지 미리보기 갱신
        self._update_gapji_preview(mode)

    def _update_rt_ratio(self):
        try:
            total_w = self.rt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.rt_paned.sash_coord(0)[0]
                self.rt_pane_ratio = current_sash / total_w
                self.config['RT_SASH_RATIO'] = self.rt_pane_ratio
                self.save_settings()
        except: pass

    def _on_rt_paned_configure(self, event):
        try:
            total_w = self.rt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.rt_paned:
                    new_pos = int(total_w * self.rt_pane_ratio)
                    self.rt_paned.sash_place(0, new_pos, 0)
        except: pass

    def _on_rt_kogas_mode_change(self):
        """가스공사 모드 변경 시 템플릿 경로 자동 설정 및 미리보기 컬럼 업데이트"""
        # 서브 모드 전환 중이면 경로 자동 설정을 건너뜀 (중복 실행 방지)
        if getattr(self, '_sub_mode_switching', False):
            self._update_rt_preview_columns("KOGAS" if self.rt_kogas_mode.get() else "RT")
            return
        if self.rt_kogas_mode.get():
            # 가스공사 템플릿 경로 탐색
            data_dir = os.path.join(BASE_DIR, "Na-aba", "home", "data")
            if not os.path.exists(data_dir):
                data_dir = os.path.join(os.getcwd(), "Na-aba", "home", "data")
            
            kogas_template = os.path.join(data_dir, "가스공사 의뢰서.xlsx")
            if os.path.exists(kogas_template):
                self.kogas_template_file_path.set(kogas_template)
                self.log("💡 가스공사 전용 템플릿이 자동으로 선택되었습니다.")
            else:
                self.log("⚠️ 가스공사 의뢰서.xlsx 파일을 찾을 수 없습니다. (data 폴더 확인 필요)")
        else:
            # 일반 KS 양식으로 복구 시도 (있는 경우)
            data_dir = os.path.join(BASE_DIR, "Na-aba", "home", "data")
            if not os.path.exists(data_dir):
                data_dir = os.path.join(os.getcwd(), "Na-aba", "home", "data")
                
            ks_template = os.path.join(data_dir, "RT KS양식.xlsx")
            if os.path.exists(ks_template):
                self.rt_template_file_path.set(ks_template)
                self.log("💡 일반 KS 양식 템플릿으로 복구되었습니다.")
        
        # [NEW] 미리보기 컬럼 구성 업데이트
        self._update_rt_preview_columns("KOGAS" if self.rt_kogas_mode.get() else "RT")

    def _update_rt_preview_columns(self, mode="RT"):
        """모드 여부에 따라 RT 미리보기 컬럼을 동적으로 변경 (사용자 설정 키 준수)"""
        tree = getattr(self, "rt_preview_tree", None) if mode == "RT" else getattr(self, "kogas_preview_tree", None)
        if not tree: return
        
        # [DYNAMIC] 고정된 리스트 대신 사용자가 설정한 키(rt_column_keys 등)를 사용
        keys_attr = "rt_column_keys" if mode == "RT" else "kogas_column_keys"
        current_cols = list(getattr(self, keys_attr))
        if 'Report No' not in current_cols:
            current_cols.insert(current_cols.index('Date') + 1 if 'Date' in current_cols else 2, 'Report No')
        setattr(self, keys_attr, current_cols)
            
        tree["columns"] = tuple(current_cols)
        tree["displaycolumns"] = tuple(current_cols)
        
        # [NEW] 기본 표시 이름 맵핑
        default_names = {
            "selected": "V", "No": "No", "Date": "Date", "Report No": "Report No", "Sec": "SEC", "Dwg": "Drawing No.",
            "Joint": "Film Ident. No.", "Loc": "Film Location", "T": "T", "Mat": "Mat",
            "Size": "구경",
            "Acc": "Acc", "Rej": "Rej", "Deg": "Deg", "Welder": "Welder No", "Remarks": "Remarks",
            "Dwg_Sub": "Dwg(Sub)", "Welder_Sub": "Welder(Sub)", "Mat_Sub": "Mat(Sub)",
            "D1": "① Crack", "D2": "② IP", "D3": "③ LF", "D4": "④ Slag", "D5": "⑤ Por",
            "D6": "⑥ U/C", "D7": "⑦ RUC", "D8": "⑧ BT", "D9": "⑨ TI", "D10": "⑩ CP",
            "D11": "⑪ RC", "D12": "⑫ Mis", "D13": "⑬ EP", "D14": "⑭ SD", "D15": "⑮ Oth"
        }
        
        saved_widths = self.config.get(f"{mode}_COL_WIDTHS", {})
        default_widths = {
            "selected": 40, "No": 50, "Date": 90, "Sec": 90, "Dwg": 300, "Joint": 120, "Loc": 100, 
            "Report No": 220,
            "Acc": 40, "Rej": 40, "Deg": 40, "Welder": 100, "Remarks": 120,
            "Dwg_Sub": 200, "Welder_Sub": 100, "Mat_Sub": 100
        }
        
        # [NEW] 내부 컬럼 ID와 설정 Key(rt_items 정의) 간의 맵핑
        key_map = {
            "T": "THK", "Remarks": "REM", "Result": "RES"
        }
        
        for col in current_cols:
            suffix = key_map.get(col, col.upper())
            if mode == "KOGAS":
                name_key = f"KOGAS_R_NAME_{suffix}"
            else:
                name_key = f"{mode}_NAME_{suffix}"
            if col == "selected": name_key = None
            
            display_text = None
            if name_key:
                # 1. 현재 UI 입력값 확인 (최우선)
                if name_key in self.setting_vars:
                    display_text = self.setting_vars[name_key].get()
                
                # 2. 설정 파일(config) 확인
                if not display_text:
                    display_text = self.config.get(name_key)
                
                # 3. 보조 키 형식 확인 (소문자 등 대비)
                if not display_text:
                    if mode == "KOGAS":
                        alt_key = f"KOGAS_R_NAME_{col}"
                    else:
                        alt_key = f"{mode}_NAME_{col}"
                    if alt_key in self.setting_vars:
                        display_text = self.setting_vars[alt_key].get()
                    else:
                        display_text = self.config.get(alt_key)

            fallback_name = default_names.get(col, col)
            if not display_text:
                display_text = fallback_name
                
            tree.heading(col, text=display_text, anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode=mode))
            w = saved_widths.get(col, default_widths.get(col, 80))
            tree.column(col, width=w, anchor='center', stretch=False)
        
        # 데이터가 이미 있다면 리프레시
        data = self.rt_extracted_data if mode == "RT" else self.kogas_extracted_data
        if data:
            self.populate_preview(data, mode=mode, switch_tab=False)
        elif mode in getattr(self, '_rt_check_menus', {}):
            self._refresh_rt_filter_headers(mode)

    def _setup_pt_ui(self, parent):
        # [FORCE] Ensure parent (pt_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.pt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.pt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.pt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.pt_paned.add(left_container, width=425)
        self._add_sidebar_toggle(self.pt_paned, left_container)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PT 성적서 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "PT 데이터:", self.pt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "PT 양식:", self.pt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_frame, background="#f0f0f0")
        btn_frame.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("PT")).pack(side='left', padx=2)

        # 2. Configuration Tabs
        pt_config_frame = ttk.LabelFrame(left_pane, text=" 리포트 세부 설정 ", padding=2)
        pt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.pt_tab_notebook = ttk.Notebook(pt_config_frame)
        self.pt_tab_notebook.pack(fill='both', expand=True)

        pt_tab_cover = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_data = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_rows = ttk.Frame(self.pt_tab_notebook, padding=5)
        pt_tab_cols = ttk.Frame(self.pt_tab_notebook, padding=5)
        
        # [CRITICAL] Allow children to expand
        for t in [pt_tab_cover, pt_tab_data, pt_tab_rows, pt_tab_cols]: t.columnconfigure(0, weight=1)

        self.pt_tab_notebook.add(pt_tab_cover, text="갑지")
        self.pt_tab_notebook.add(pt_tab_data, text="을지")
        self.pt_tab_notebook.add(pt_tab_rows, text="행 설정")
        self.pt_tab_notebook.add(pt_tab_cols, text="컬럼 설정")

        # PT templates already contain their required logos and footer images.
        # Keep only metadata and print/layout controls in the PT tabs; exposing
        # separate image-path/position fields here was redundant and confusing.
        self._create_gapji_meta_ui(pt_tab_cover, use_pack=False, mode="PT")
        self.pt_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        self._create_margin_settings(pt_tab_cover, "PT_COVER", use_pack=False)
        self._create_margin_settings(pt_tab_data, "PT_DATA", use_pack=False)
        self._create_row_settings(pt_tab_rows, mode="PT")
        
        pt_items = [
            ("순번(No):", "PT_COL_NO", 1, "PT_NAME_NO", "No", "No"),
            ("ISO/Dwg:", "PT_COL_DWG", 2, "PT_NAME_DWG", "Dwg", "Dwg"),
            ("Joint No:", "PT_COL_JOINT", 5, "PT_NAME_JOINT", "Joint", "Joint"),
            ("NPS:", "PT_COL_NPS", 6, "PT_NAME_NPS", "NPS", "NPS"),
            ("두께(Th'k):", "PT_COL_THK", 7, "PT_NAME_THK", "Thk.", "Thk."),
            ("재질(Material):", "PT_COL_MAT", 8, "PT_NAME_MAT", "Material", "Material"),
            ("용접사(Welder):", "PT_COL_WELDER", 9, "PT_NAME_WELDER", "Welder", "Welder"),
            ("검사타입(Type):", "PT_COL_TYPE", 10, "PT_NAME_TYPE", "WType", "WType"),
            ("판정 결과:", "PT_COL_RES", 11, "PT_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(pt_tab_cols, "PT", pt_items)
        

        # 3. Action Section
        action_frame = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=2)
        ttk.Label(filter_row, text="📊 특정순번:", background="#ffffff", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.pt_paned, background="#f3f4f6")
        self.pt_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "PT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["PT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_pt_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "PT")
        self._apply_sash_ratio("PT")
        
        self.pt_template_file_path.trace_add("write", lambda *a: self.load_template_specific_config("PT"))

        # Adaptive Resizing Bindings
        self.pt_pane_ratio = self.config.get('PT_SASH_RATIO', 0.5)
        self.pt_paned.bind("<Configure>", self._on_pt_paned_configure)
        self.pt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_pt_ratio))
        self.root.after(500, lambda: self._on_pt_paned_configure(None))

    def _setup_mt_ui(self, parent):
        # [FORCE] Ensure parent (pt_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.mt_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.mt_paned.pack(fill='both', expand=True)

        # [LEFT] Settings Sidebar
        left_container = tk.Frame(self.mt_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.mt_paned.add(left_container, width=425)
        self._add_sidebar_toggle(self.mt_paned, left_container)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 MT 성적서 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "MT 데이터:", self.mt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "MT 양식:", self.mt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])
        
        # [NEW] Template Config Save Button
        btn_frame = tk.Frame(file_frame, background="#f0f0f0")
        btn_frame.grid(row=3, column=1, sticky='ew', pady=2)
        ttk.Button(btn_frame, text="💾 현재 설정을 이 양식 전용으로 저장", command=lambda: self.save_template_specific_config("MT")).pack(side='left', padx=2)

        # 2. Configuration Tabs
        pt_config_frame = ttk.LabelFrame(left_pane, text=" 리포트 세부 설정 ", padding=2)
        pt_config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.mt_tab_notebook = ttk.Notebook(pt_config_frame)
        self.mt_tab_notebook.pack(fill='both', expand=True)

        pt_tab_cover = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_data = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_rows = ttk.Frame(self.mt_tab_notebook, padding=5)
        pt_tab_cols = ttk.Frame(self.mt_tab_notebook, padding=5)
        
        # [CRITICAL] Allow children to expand
        for t in [pt_tab_cover, pt_tab_data, pt_tab_rows, pt_tab_cols]: t.columnconfigure(0, weight=1)

        self.mt_tab_notebook.add(pt_tab_cover, text="갑지")
        self.mt_tab_notebook.add(pt_tab_data, text="을지")
        self.mt_tab_notebook.add(pt_tab_rows, text="행 설정")
        self.mt_tab_notebook.add(pt_tab_cols, text="컬럼 설정")

        # [ALIGNED] Mode-specific context for logo grid
        self._create_gapji_meta_ui(pt_tab_cover, use_pack=False, mode="MT")
        self.mt_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        next_row_mt_cover = self._create_setting_grid(pt_tab_cover, "MT_COVER")
        next_row_mt_data = self._create_setting_grid(pt_tab_data, "MT_DATA")
        self._create_margin_settings(pt_tab_cover, "MT_COVER", use_pack=False)
        self._create_margin_settings(pt_tab_data, "MT_DATA", use_pack=False)
        self._create_row_settings(pt_tab_rows, mode="MT")
        
        pt_items = [
            ("순번(No):", "MT_COL_NO", 1, "MT_NAME_NO", "No", "No"),
            ("ISO/Dwg:", "MT_COL_DWG", 2, "MT_NAME_DWG", "Dwg", "Dwg"),
            ("Joint No:", "MT_COL_JOINT", 5, "MT_NAME_JOINT", "Joint", "Joint"),
            ("NPS:", "MT_COL_NPS", 6, "MT_NAME_NPS", "NPS", "NPS"),
            ("두께(Th'k):", "MT_COL_THK", 7, "MT_NAME_THK", "Thk.", "Thk."),
            ("재질(Material):", "MT_COL_MAT", 8, "MT_NAME_MAT", "Material", "Material"),
            ("용접사(Welder):", "MT_COL_WELDER", 9, "MT_NAME_WELDER", "Welder", "Welder"),
            ("검사타입(Type):", "MT_COL_TYPE", 10, "MT_NAME_TYPE", "WType", "WType"),
            ("판정 결과:", "MT_COL_RES", 11, "MT_NAME_RES", "Result", "Result")
        ]
        self._create_column_mapping_ui(pt_tab_cols, "MT", pt_items)
        

        # 3. Action Section
        action_frame = tk.Frame(left_pane, background="#ffffff", highlightthickness=1, highlightbackground="#d1d5db", padx=10, pady=5)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=2)
        ttk.Label(filter_row, text="📊 특정순번:", background="#ffffff", font=("Malgun Gothic", 8)).pack(side='left')
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=15, font=("Arial", 9), exportselection=False).pack(side='left', padx=5, fill='x', expand=True)

        # Actions moved to right Top Action Bar

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.mt_paned, background="#f3f4f6")
        self.mt_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "MT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["MT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_mt_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "MT")
        self._apply_sash_ratio("MT")
        
        self.mt_template_file_path.trace_add("write", lambda *a: self.load_template_specific_config(self.mt_template_file_path.get(), "MT"))

        # Adaptive Resizing Bindings
        self.mt_pane_ratio = self.config.get('MT_SASH_RATIO', 0.5)
        self.mt_paned.bind("<Configure>", self._on_mt_paned_configure)
        self.mt_paned.bind("<ButtonRelease-1>", lambda e: self.root.after(10, self._update_mt_ratio))
        self.root.after(500, lambda: self._on_mt_paned_configure(None))

    def _update_pt_ratio(self):
        try:
            total_w = self.pt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.pt_paned.sash_coord(0)[0]
                self.pt_pane_ratio = current_sash / total_w
                self.config['PT_SASH_RATIO'] = self.pt_pane_ratio
                self.save_settings()
        except: pass

    def _update_mt_ratio(self):
        try:
            total_w = self.mt_paned.winfo_width()
            if total_w > 100:
                current_sash = self.mt_paned.sash_coord(0)[0]
                self.mt_pane_ratio = current_sash / total_w
                self.config['MT_SASH_RATIO'] = self.mt_pane_ratio
                self.save_settings()
        except: pass

    def _on_pt_paned_configure(self, event):
        try:
            total_w = self.pt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.pt_paned:
                    new_pos = int(total_w * self.pt_pane_ratio)
                    self.pt_paned.sash_place(0, new_pos, 0)
        except: pass

    def _on_mt_paned_configure(self, event):
        try:
            total_w = self.mt_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.mt_paned:
                    new_pos = int(total_w * self.mt_pane_ratio)
                    self.mt_paned.sash_place(0, new_pos, 0)
        except: pass

    def _setup_paut_ui(self, parent):
        # [FORCE] Ensure parent (paut_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        self.paut_paned = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        self.paut_paned.pack(fill='both', expand=True)

        # [LEFT] Settings & Manual Eval Pane
        left_container = tk.Frame(self.paut_paned, background="#f9fafb", highlightthickness=0, borderwidth=0)
        self.paut_paned.add(left_container, width=425)
        self._add_sidebar_toggle(self.paut_paned, left_container)
        
        # Scrollable area (Full size)
        left_pane = self._create_scrollable_sidebar(left_container)

        # FIXED FLOATING Header
        header_frame = tk.Frame(left_container, background="#f9fafb", highlightthickness=0, borderwidth=0)
        header_frame.place(x=0, y=0, relwidth=1, height=40)
        
        inner_header = tk.Frame(header_frame, background="#f9fafb", padx=20)
        inner_header.pack(fill='both', expand=True, pady=(5, 0))
        
        tk.Label(inner_header, text="🔬 PAUT (ASME B31.1) 관리", font=("Malgun Gothic", 15, "bold"), 
                 background="#f9fafb", foreground="#1e3a8a").pack(side='left')
        
        tk.Frame(header_frame, height=1, background="#e5e7eb").pack(side='bottom', fill='x')
        header_frame.lift()

        # Spacer in scrollable area
        tk.Frame(left_pane, height=40, background="#f9fafb").pack(fill='x')

        # 1. File Selection Group
        file_frame = ttk.LabelFrame(left_pane, text=" 데이터 및 양식 (Data) ", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        def _add_file_row(parent_f, label, var, row, is_dir=False, types=None):
            parent_f.columnconfigure(1, weight=1)
            ttk.Label(parent_f, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=0, sticky='e', padx=2, pady=2)
            ttk.Entry(parent_f, textvariable=var, font=("Arial", 9), width=1, exportselection=False).grid(row=row, column=1, padx=2, pady=2, sticky='ew')
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_f, text="...", width=3, command=cmd).grid(row=row, column=2, padx=2, pady=2)

        _add_file_row(file_frame, "기본 로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "PAUT 데이터:", self.paut_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "PAUT 양식:", self.paut_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # --- 2. Configuration Notebook (갑, 을, 행 설정, 개별 판정) ---
        config_frame = ttk.LabelFrame(left_pane, text=" 설정 및 판정 (Config & Eval) ", padding=2)
        config_frame.pack(fill='both', expand=True, pady=(0, 10))

        self.paut_tab_notebook = ttk.Notebook(config_frame)
        self.paut_tab_notebook.pack(fill='both', expand=True)

        tab_cover = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_data = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_rows = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_cols = ttk.Frame(self.paut_tab_notebook, padding=5)
        tab_eval = ttk.Frame(self.paut_tab_notebook, padding=5)

        # Allow expansion
        for t in [tab_cover, tab_data, tab_rows, tab_cols, tab_eval]: t.columnconfigure(0, weight=1)

        self.paut_tab_notebook.add(tab_eval, text="개별 판정") # Default/Main tab
        self.paut_tab_notebook.add(tab_cover, text="갑지")
        self.paut_tab_notebook.add(tab_data, text="을지")
        self.paut_tab_notebook.add(tab_rows, text="행 설정")
        self.paut_tab_notebook.add(tab_cols, text="컬럼 설정")
        
        # [NEW] PAUT Equipment Settings
        tab_equip_outer = ttk.Frame(self.paut_tab_notebook, padding=0)
        self.paut_tab_notebook.add(tab_equip_outer, text="장비/설정")
        
        equip_canvas = tk.Canvas(tab_equip_outer, background="#f9fafb", highlightthickness=0)
        equip_scrollbar = ttk.Scrollbar(tab_equip_outer, orient="vertical", command=equip_canvas.yview)
        tab_equip = ttk.Frame(equip_canvas, padding=5)
        
        tab_equip.bind("<Configure>", lambda e: equip_canvas.configure(scrollregion=equip_canvas.bbox("all")))
        
        # We don't force canvas width to avoid squeezing the Entry fields to 0 width.
        equip_canvas.create_window((0, 0), window=tab_equip, anchor="nw")
        
        # Use existing mousewheel binding if available
        if hasattr(self, '_bind_mousewheel_recursive'):
            self._bind_mousewheel_recursive(tab_equip)
            
        equip_canvas.configure(yscrollcommand=equip_scrollbar.set)
        equip_canvas.pack(side="left", fill="both", expand=True)
        equip_scrollbar.pack(side="right", fill="y")
        
        for col_idx in [1, 3]: # 2 columns layout
            tab_equip.columnconfigure(col_idx, weight=1)
        
        self.paut_equip_vars = {}
        self.paut_equip_map = {
            'equip': ('검사장비', 'A12', 'OmniScan MX2'),
            'module': ('검사모듈', 'H12', '32-128PR'),
            'probe': ('탐촉자', 'O12', '5L64-A2 / 5L16A10'),
            'wedge': ('웻지', 'V12', 'SA2-N55S-IHC / SA10-N55S'),
            'program': ('판독프로그램', 'AC12', 'Tomoview, Omni PC'),
            'method': ('검사기법', 'A15', 'PA/Sectoria'),
            'wave': ('파형', 'H15', '횡파'),
            'velocity': ('음속', 'O15', '3246.8m/s'),
            'filter': ('Band-Pass Filter', 'V15', 'Band-Pass 5MHz'),
            'medium': ('접촉 매질', 'AC15', '물 & 글리세린'),
            'voltage': ('Voltage', 'A17', '80V'),
            'gain': ('Gain', 'H17', '11 + 6dB / 21.4 + 6dB'),
            'mode': ('Mode', 'O17', 'PE(Pulse-Echo)'),
            'offset': ('Index Offset', 'V17', '±17 mm'),
            'ref_block': ('Ref Block', 'AC17', 'Test Level A Block(Φ2.5)'),

            'tcg1_p1': ('TCG1 Pos 1', 'H19', '0'),
            'tcg1_g1': ('TCG1 Gain 1', 'O19', '4.2dB'),
            'tcg1_p2': ('TCG1 Pos 2', 'H20', '8.96mm'),
            'tcg1_g2': ('TCG1 Gain 2', 'O20', '4.5dB'),
            'tcg1_p3': ('TCG1 Pos 3', 'H21', '17.13mm'),
            'tcg1_g3': ('TCG1 Gain 3', 'O21', '7.0dB'),
            'tcg1_p4': ('TCG1 Pos 4', 'H22', '24.79mm'),
            'tcg1_g4': ('TCG1 Gain 4', 'O22', '12.7dB'),

            'tcg2_p1': ('TCG2 Pos 1', 'AA19', '0'),
            'tcg2_g1': ('TCG2 Gain 1', 'AG19', '3.2dB'),
            'tcg2_p2': ('TCG2 Pos 2', 'AA20', '7.61mm'),
            'tcg2_g2': ('TCG2 Gain 2', 'AG20', '2.7dB'),
            'tcg2_p3': ('TCG2 Pos 3', 'AA21', '15.35mm'),
            'tcg2_g3': ('TCG2 Gain 3', 'AG21', '5.3dB'),
            'tcg2_p4': ('TCG2 Pos 4', 'AA22', '23.17mm'),
            'tcg2_g4': ('TCG2 Gain 4', 'AG22', '8.1dB'),

            'set1_qty': ('Gr.1 Used Elem', 'A25', '16'),
            'set1_first': ('Gr.1 First Elem', 'H25', '1'),
            'set1_last': ('Gr.1 Last Elem', 'O25', '16'),
            'set1_scan': ('Gr.1 Scan Res', 'V25', '1'),
            'set1_focal': ('Gr.1 Focal', 'AC25', '10'),
            'set1_start': ('Gr.1 Start Ang', 'A27', '45°'),
            'set1_stop': ('Gr.1 Stop Ang', 'H27', '75°'),
            'set1_ang_res': ('Gr.1 Ang Res', 'O27', '1°'),
            'set1_law': ('Gr.1 Law Config', 'V27', 'Sectorial'),
            
            'set2_qty': ('Gr.2 Used Elem', 'A30', '16'),
            'set2_first': ('Gr.2 First Elem', 'H30', '49'),
            'set2_last': ('Gr.2 Last Elem', 'O30', '64'),
            'set2_scan': ('Gr.2 Scan Res', 'V30', '1'),
            'set2_focal': ('Gr.2 Focal', 'AC30', '10'),
            'set2_start': ('Gr.2 Start Ang', 'A32', '45°'),
            'set2_stop': ('Gr.2 Stop Ang', 'H32', '75°'),
            'set2_ang_res': ('Gr.2 Ang Res', 'O32', '1°'),
            'set2_law': ('Gr.2 Law Config', 'V32', 'Sectorial')
        }
        
        equip_r, equip_c = 0, 0
        paut_standard_revision = "CENTRAL_PAUT_20260910"
        apply_new_standard = (
            self.config.get("PAUT_EQUIP_STANDARD_REV") != paut_standard_revision
        )
        for k, (label, cell, default_val) in self.paut_equip_map.items():
            config_val = self.config.get(f'PAUT_EQUIP_{k.upper()}')
            if apply_new_standard or config_val is None or str(config_val).strip() == '':
                config_val = default_val
                self.config[f'PAUT_EQUIP_{k.upper()}'] = default_val
            var = tk.StringVar(value=config_val)
            self.paut_equip_vars[k] = var
            ttk.Label(tab_equip, text=label).grid(row=equip_r, column=equip_c*2, sticky='e', padx=5, pady=5)
            ent = ttk.Entry(tab_equip, textvariable=var, width=16)
            ent.grid(row=equip_r, column=equip_c*2+1, sticky='ew', padx=5, pady=5)
            ent.bind("<FocusOut>", lambda e: self.save_settings())
            ent.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            
            equip_c += 1
            if equip_c > 1: # 2 columns layout
                equip_c = 0
                equip_r += 1
        self.config["PAUT_EQUIP_STANDARD_REV"] = paut_standard_revision

        # [NEW] Populate Settings
        self._create_gapji_meta_ui(tab_cover, use_pack=False, mode="PAUT")
        self.paut_tab_notebook.bind("<<NotebookTabChanged>>", self._update_gapji_preview_current)
        next_row_cover = self._create_setting_grid(tab_cover, "PAUT_COVER")
        next_row_data = self._create_setting_grid(tab_data, "PAUT_DATA")
        self._create_margin_settings(tab_cover, "PAUT_COVER", use_pack=False)
        self._create_margin_settings(tab_data, "PAUT_DATA", use_pack=False)
        self._create_row_settings(tab_rows, mode="PAUT")
        
        paut_items = [
            ("순번(No):", "PAUT_COL_NO", 1, "PAUT_COL_NO_NAME", "No", "No"),
            ("촬영일자(Date):", "PAUT_COL_DATE", 0, "PAUT_NAME_DATE", "Date", "Date"),
            ("Line No.:", "PAUT_COL_LINE", 2, "PAUT_NAME_LINE", "Line No.", "Line No."),
            ("Joint No.:", "PAUT_COL_JOINT", 3, "PAUT_NAME_JOINT", "Joint No.", "Joint No."),
            ("크기(Size):", "PAUT_COL_SIZE", 0, "PAUT_NAME_SIZE", "Size", "Size"),
            ("두께(Th'k):", "PAUT_COL_THK", 4, "PAUT_NAME_THK", "Th'k(mm)", "Th'k(mm)"),
            ("합격(Acc):", "PAUT_COL_ACC", 0, "PAUT_NAME_ACC", "Acc", "Acc"),
            ("불합격(Rej):", "PAUT_COL_REJ", 0, "PAUT_NAME_REJ", "Rej", "Rej"),
            ("결함높이(H):", "PAUT_COL_H", 5, "PAUT_NAME_H", "Height(mm)", "Height(mm)"),
            ("TCG MAX(%):", "PAUT_COL_TCG", 0, "PAUT_NAME_TCG", "TCG MAX(%)", "TCG MAX(%)"),
            ("결함길이(L):", "PAUT_COL_L", 6, "PAUT_NAME_L", "Length(mm)", "Length(mm)"),
            ("결함깊이(D):", "PAUT_COL_UP", 7, "PAUT_NAME_UP", "Upper", "Upper"),
            ("하부깊이(Low):", "PAUT_COL_LOW", 0, "PAUT_NAME_LOW", "Lower", "Lower"),
            ("시작위치:", "PAUT_COL_START", 0, "PAUT_NAME_START", "Start", "Start"),
            ("종료위치:", "PAUT_COL_END", 0, "PAUT_NAME_END", "End", "End"),
            ("a/l 비:", "PAUT_COL_AL", 0, "PAUT_NAME_AL", "a/l", "a/l"),
            ("a/t 비:", "PAUT_COL_AT", 0, "PAUT_NAME_AT", "a/t", "a/t"),
            ("결함종류:", "PAUT_COL_NAT", 8, "PAUT_NAME_NAT", "Type of Flaw", "Type of Flaw"),
            ("용접사(Welder):", "PAUT_COL_WELDER", 0, "PAUT_NAME_WELDER", "Welder", "Welder"),
            ("검사길이:", "PAUT_COL_TESTED_LEN", 0, "PAUT_NAME_TESTED_LEN", "Tested Length", "Tested Length"),
            ("판정(Eval):", "PAUT_COL_EVAL", 9, "PAUT_NAME_EVAL", "Evaluation", "Evaluation"),
            ("비고(Rem):", "PAUT_COL_REM", 10, "PAUT_NAME_REM", "Remarks", "Remarks")
        ]
        # [NEW] PAUT Column Settings Notebook for different templates
        self.paut_col_nb = ttk.Notebook(tab_cols)
        self.paut_col_nb.pack(fill='both', expand=True, padx=2, pady=2)
        
        for template_name in ["지역난방", "ASME", "KS", "AWS"]:
            frame = ttk.Frame(self.paut_col_nb)
            self.paut_col_nb.add(frame, text=template_name)
            
            # Localize the keys
            localized_items = []
            for item in paut_items:
                col_key = item[1].replace("PAUT_", f"PAUT_{template_name}_")
                name_key = item[3].replace("PAUT_", f"PAUT_{template_name}_")
                default_col = (23 if template_name == "지역난방" else 0) if item[1] == "PAUT_COL_TCG" else item[2]
                localized_items.append((item[0], col_key, default_col, name_key, item[4], item[5]))
                
            self._create_column_mapping_ui(frame, f"PAUT_{template_name}", localized_items)

        # --- 3. Manual Evaluation Content (Inside tab_eval) ---
        manual_container = tk.Frame(tab_eval, background="#f9fafb")
        manual_container.pack(fill='both', expand=True)
        
        # [NEW] Evaluation Mode Selector
        mode_frame = tk.Frame(manual_container, background="#f9fafb")
        mode_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(mode_frame, text="✅ 판정 기준:", font=("Malgun Gothic", 8, "bold")).pack(side='left', padx=2)
        tk.Radiobutton(mode_frame, text="일반(B31.1)", variable=self.paut_eval_mode, value="Standard", background="#f9fafb").pack(side='left', padx=5)
        tk.Radiobutton(mode_frame, text="ECA(파과역학)", variable=self.paut_eval_mode, value="ECA", background="#f9fafb").pack(side='left', padx=5)

        # Sizing Helper Row
        calc_frame = tk.Frame(manual_container, background="#f1f5f9", padx=10, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
        calc_frame.pack(fill='x', pady=(0, 10))
        
        # Row 0: H-Calc Settings
        h_set_row = tk.Frame(calc_frame, background="#f1f5f9")
        h_set_row.pack(fill='x')
        ttk.Label(h_set_row, text="📏 H-Setting:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(h_set_row, textvariable=self.paut_manual_vars['peak'], width=5).pack(side='left', padx=2)
        ttk.Label(h_set_row, text="% /", background="#f1f5f9").pack(side='left')
        tk.Entry(h_set_row, textvariable=self.paut_manual_vars['db'], width=4).pack(side='left', padx=2)
        ttk.Label(h_set_row, text="dB ->", background="#f1f5f9").pack(side='left', padx=2)
        ttk.Label(h_set_row, textvariable=self.paut_manual_vars['target_fsh'], background="#f1f5f9", foreground="#ef4444", font=("Arial", 8, "bold")).pack(side='left', padx=2)
        
        # Row 1: H-Calc Data
        h_data_row = tk.Frame(calc_frame, background="#f1f5f9")
        h_data_row.pack(fill='x', pady=2)
        ttk.Label(h_data_row, text="📏 H-Data:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(h_data_row, textvariable=self.paut_manual_vars['z1'], width=6).pack(side='left', padx=2)
        ttk.Label(h_data_row, text="~", background="#f1f5f9").pack(side='left')
        tk.Entry(h_data_row, textvariable=self.paut_manual_vars['z2'], width=6).pack(side='left', padx=2)
        ttk.Button(h_data_row, text="입력", width=6, command=self._calculate_paut_h).pack(side='right', padx=5)

        # Row 2: L-Calc Data
        l_data_row = tk.Frame(calc_frame, background="#f1f5f9")
        l_data_row.pack(fill='x')
        ttk.Label(l_data_row, text="📏 L-Data:", font=("Malgun Gothic", 8, "bold"), background="#f1f5f9").pack(side='left', padx=2)
        tk.Entry(l_data_row, textvariable=self.paut_manual_vars['L1'], width=6).pack(side='left', padx=2)
        ttk.Label(l_data_row, text="~", background="#f1f5f9").pack(side='left')
        tk.Entry(l_data_row, textvariable=self.paut_manual_vars['L2'], width=6).pack(side='left', padx=2)
        ttk.Button(l_data_row, text="입력", width=6, command=self._calculate_paut_l).pack(side='right', padx=5)

        # Input Grid
        input_grid = tk.Frame(manual_container, background="#f9fafb")
        input_grid.pack(fill='x', pady=5)
        
        m_inputs = [("두께(T):", "t"), ("높이(H):", "h"), ("길이(L):", "l"), ("깊이(d):", "d"), 
                    ("시작(S):", "L1"), ("종료(E):", "L2"), ("상단(Up):", "D_Upper"), ("하단(Low):", "D_Lower")]
        for i, (lbl, key) in enumerate(m_inputs):
            grid_row = i // 2
            grid_col = (i % 2) * 2
            ttk.Label(input_grid, text=lbl, font=("Malgun Gothic", 8)).grid(row=grid_row, column=grid_col, sticky='e', padx=2, pady=1)
            ent = tk.Entry(input_grid, textvariable=self.paut_manual_vars[key], width=10)
            ent.grid(row=grid_row, column=grid_col+1, padx=2, pady=1, sticky='w')
            ent.bind("<Return>", lambda e: self.root.focus_set())

        # Nature & Evaluation Actions
        nature_row = tk.Frame(manual_container, background="#f9fafb")
        nature_row.pack(fill='x', pady=5)
        ttk.Label(nature_row, text="종류:").pack(side='left', padx=2)
        ttk.Combobox(nature_row, textvariable=self.paut_manual_vars['nature'], values=["Crack", "LOF", "IP", "Slag", "Porosity", "Others"], width=9).pack(side='left', padx=2)
        ttk.Label(nature_row, text="위치:").pack(side='left', padx=2)
        ttk.Label(nature_row, textvariable=self.paut_manual_vars['loc'], foreground="#3b82f6", font=("Arial", 9, "bold")).pack(side='left', padx=2)
        
        btn_row = tk.Frame(manual_container, background="#f9fafb")
        btn_row.pack(fill='x', pady=2)
        ttk.Button(btn_row, text="판정", width=8, command=self._run_manual_paut_eval).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(btn_row, text="저장", width=8, command=self._save_paut_manual_to_tree).pack(side='left', expand=True, fill='x', padx=2)

        self.paut_res_label = tk.Label(manual_container, text="결과 대기", font=("Malgun Gothic", 10, "bold"), background="#f3f4f6")
        self.paut_res_label.pack(fill='x', pady=(5, 0))

        # Bottom Actions
        # [NEW] Photo Log Pages Input
        photo_pages_frame = tk.Frame(left_pane, background="#f9fafb")
        photo_pages_frame.pack(fill='x', pady=(5, 2))
        ttk.Label(photo_pages_frame, text="첨부할 사진대장 장수:", font=("Malgun Gothic", 9)).pack(side='left', padx=2)
        self.paut_photo_pages_var = tk.StringVar(value=str(self.config.get("PAUT_PHOTO_PAGES", "0")))
        ttk.Entry(photo_pages_frame, textvariable=self.paut_photo_pages_var, width=5).pack(side='left', padx=2)
        self.paut_photo_pages_var.trace_add(
            "write", lambda *_: self._sync_paut_photo_page_numbers()
        )
        
        # Core Actions moved to right Top Action Bar
        paut_photo_btn_row = tk.Frame(left_pane, background="#f9fafb")
        paut_photo_btn_row.pack(fill='x', pady=(0, 2))
        ttk.Button(paut_photo_btn_row, text=" 📸 사진대장 목록에서 PAUT 추출 (자동파싱) ", command=self._extract_paut_from_photos).pack(side='left', fill='x', expand=True)

        paut_session_row = tk.Frame(left_pane, background="#f9fafb")
        paut_session_row.pack(fill='x', pady=(0, 10))
        ttk.Button(paut_session_row, text=" 💾 세션 저장 ", command=self._export_paut_session).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(paut_session_row, text=" 📂 세션 불러오기 ", command=self._import_paut_session).pack(side='left', fill='x', expand=True)

        # [RIGHT] Multi-Preview Pane (Data & Gapji)
        right_container = tk.Frame(self.paut_paned, background="#f3f4f6")
        self.paut_paned.add(right_container, stretch="always")
        
        self._create_top_action_bar(right_container, "PAUT")
        if not hasattr(self, 'preview_notebooks'): self.preview_notebooks = {}
        nb = ttk.Notebook(right_container)
        nb.pack(fill='both', expand=True)
        self.preview_notebooks["PAUT"] = nb
        
        tab_data = tk.Frame(nb, background="#ffffff")
        tab_gapji = tk.Frame(nb, background="#ffffff")
        nb.add(tab_data, text=" 🔬 데이터 미리보기 ")
        nb.add(tab_gapji, text=" 📄 갑지 미리보기 ")
        
        self._create_paut_preview_ui(tab_data)
        self._create_gapji_preview_ui(tab_gapji, "PAUT")
        self._apply_sash_ratio("PAUT")

    def _update_paut_ratio(self):
        try:
            total_w = self.paut_paned.winfo_width()
            if total_w > 100:
                current_sash = self.paut_paned.sash_coord(0)[0]
                self.paut_pane_ratio = current_sash / total_w
                self.config['PAUT_SASH_RATIO'] = self.paut_pane_ratio
                self.save_settings()
        except: pass

    def _on_paut_paned_configure(self, event):
        try:
            total_w = self.paut_paned.winfo_width()
            if total_w > 100:
                if event and event.widget == self.paut_paned:
                    new_pos = int(total_w * self.paut_pane_ratio)
                    self.paut_paned.sash_place(0, new_pos, 0)
        except: pass

    def _update_paut_auto_loc(self):
        try:
            t = float(self.paut_manual_vars['t'].get() or 0)
            d = float(self.paut_manual_vars['d'].get() or 0)
            h = float(self.paut_manual_vars['h'].get() or 0)
            if t > 0:
                is_surface = (d <= 0.1 * t) or (d + h >= 0.9 * t)
                loc = "Surface" if is_surface else "Subsurface"
                self.paut_manual_vars['loc'].set(loc)
        except: pass

    def _update_paut_target_fsh(self):
        """[NEW] Peak 나 dB 변경 시 목표 % FSH 실시간 업데이트"""
        try:
            peak = float(self.paut_manual_vars['peak'].get() or 0)
            db = float(self.paut_manual_vars['db'].get() or 0)
            if peak > 0:
                drop_factor = pow(10, -db / 20)
                target = peak * drop_factor
                self.paut_manual_vars['target_fsh'].set(f"{target:.1f}%")
            else:
                self.paut_manual_vars['target_fsh'].set("0.0%")
        except:
            self.paut_manual_vars['target_fsh'].set("-")

    def _calculate_paut_h(self):
        """[NEW] dB 강하법을 이용한 h 값 계산 및 자동 입력"""
        try:
            peak = float(self.paut_manual_vars['peak'].get() or 80)
            db = float(self.paut_manual_vars['db'].get() or 6)
            z1 = float(self.paut_manual_vars['z1'].get() or 0)
            z2 = float(self.paut_manual_vars['z2'].get() or 0)
            
            # Calculate Target FSH: Peak * 10^(-dB/20)
            drop_factor = pow(10, -db / 20)
            target_fsh = peak * drop_factor
            self.paut_manual_vars['target_fsh'].set(f"{target_fsh:.1f}%")

            h_calc = abs(z1 - z2)
            self.paut_manual_vars['h'].set(f"{h_calc:.2f}")
            
            # [NEW] D: 중간값, UP: 낮은값, LOW: 높은값 자동 기입
            d_mid = (z1 + z2) / 2
            d_min = min(z1, z2)
            d_max = max(z1, z2)
            
            self.paut_manual_vars['d'].set(f"{d_mid:.2f}")
            self.paut_manual_vars['D_Upper'].set(f"{d_min:.2f}")
            self.paut_manual_vars['D_Lower'].set(f"{d_max:.2f}")

            self._update_paut_auto_loc()
            self.log(f"📐 h 산출 완료 ({peak}% peak -> {db}dB 강하 목표 {target_fsh:.1f}%): {h_calc:.2f} mm (D:{d_mid:.2f}, U:{d_min:.2f}, L:{d_max:.2f})")
        except Exception as e:
            messagebox.showerror("계산 오류", f"입력값을 확인해주세요: {e}")

    def _calculate_paut_l(self):
        """[NEW] L1, L2 값을 이용한 결함 길이(L) 자동 계산"""
        try:
            l1 = float(self.paut_manual_vars['L1'].get() or 0)
            l2 = float(self.paut_manual_vars['L2'].get() or 0)
            l_calc = abs(l1 - l2)
            self.paut_manual_vars['l'].set(f"{l_calc:.2f}")
            self.log(f"📐 L 산출 완료 (L1:{l1} ~ L2:{l2}): {l_calc:.2f} mm")
        except Exception as e:
            messagebox.showerror("계산 오류", f"L1, L2 입력값을 확인해주세요: {e}")

    def _save_paut_manual_to_tree(self):
        """[NEW] 수동 판정 구역의 입력값을 선택된 트리뷰 행(들)에 저장"""
        selected_ids = self.paut_preview_tree.selection()
        if not selected_ids:
            messagebox.showwarning("선택 누락", "저장할 행을 먼저 선택해주세요.")
            return

        try:
            t = self.to_float(self.paut_manual_vars['t'].get() or 0)
            h = self.to_float(self.paut_manual_vars['h'].get() or 0)
            l = self.to_float(self.paut_manual_vars['l'].get() or 0)
            d = self.to_float(self.paut_manual_vars['d'].get() or 0)
            l1 = self.paut_manual_vars['L1'].get()
            l2 = self.paut_manual_vars['L2'].get()
            up = self.paut_manual_vars['D_Upper'].get()
            low = self.paut_manual_vars['D_Lower'].get()
            nat = self.paut_manual_vars['nature'].get()
            rem = "" # Remarks can be added to UI if needed, for now keep existing or clear

            # Run evaluation logic to get the result
            eval_mode = self.paut_eval_mode.get()
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            
            al = f"{h/l:.3f}" if l > 0 and h > 0 else ""
            at = f"{h/t:.3f}" if t > 0 and h > 0 else ""

            mode_info = self._get_mode_info("PAUT")
            if not mode_info: return
            _, idx_map, data_list, _ = mode_info

            for item_id in selected_ids:
                view_idx = self.paut_preview_tree.index(item_id)
                if 0 <= view_idx < len(idx_map):
                    actual_idx = idx_map[view_idx]
                    item = data_list[actual_idx]
                    item["Th'k(mm)"] = t
                    item['Height(mm)'] = h
                    item['Length(mm)'] = l
                    item['Start'] = l1
                    item['End'] = l2
                    item['Upper'] = up
                    item['Lower'] = low
                    item['Type of Flaw'] = nat
                    item['Evaluation'] = f"{res} ({loc})"
                    item['a/l'] = al
                    item['a/t'] = at
            
            self.populate_preview(data_list, mode="PAUT", switch_tab=False)
            self.log(f"💾 선택된 {len(selected_ids)}개 행에 수동 판정 결과 저장 완료.")
            
            # [NEW] 저장 후 입력 필드 초기화
            self._reset_paut_manual_inputs()
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"입력값을 확인해주세요: {e}")

    def _run_manual_paut_eval(self):
        try:
            t = float(self.paut_manual_vars['t'].get() or 0)
            h = float(self.paut_manual_vars['h'].get() or 0)
            l = float(self.paut_manual_vars['l'].get() or 0)
            d = float(self.paut_manual_vars['d'].get() or 0)
            nat = self.paut_manual_vars['nature'].get()
            eval_mode = self.paut_eval_mode.get()
            
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            self.paut_manual_vars['loc'].set(loc)
            if res == "Accept":
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#27ae60")
            else:
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#e74c3c")
        except Exception as e:
            messagebox.showerror("입력 오류", f"입력값을 확인해주세요: {e}")

    def _extract_paut_from_photos(self):
        try:
            if not hasattr(self, 'photo_listbox') or self.photo_listbox.size() == 0:
                messagebox.showwarning("경고", "사진대장 탭에 로드된 사진이 없습니다.")
                return
            
            photo_list = self.photo_listbox.get(0, tk.END)
            self.paut_extracted_data = []
            
            for i, photo_path in enumerate(photo_list):
                filename = os.path.basename(photo_path)
                name_without_ext = os.path.splitext(filename)[0]
                
                # parsing logic: DT2024-016-101-0-1-600A_R04
                line_no = ""
                size = ""
                joint_no = ""
                
                if "_" in name_without_ext:
                    parts = name_without_ext.rsplit("_", 1)
                    line_no = parts[0] # DT2024-016-101-0-1-600A
                    joint_no = parts[1] # R04
                    
                    if "-" in line_no:
                        size = line_no.rsplit("-", 1)[1] # 600A
                else:
                    line_no = name_without_ext
                    
                if i == 0:
                    messagebox.showinfo("첫번째 사진 파싱 확인", f"원본 파일명: {filename}\n추출된 Line No: {line_no}\n추출된 Size: {size}\n추출된 Joint No: {joint_no}")

                    
                item = {
                    'selected': True,
                    'Line No.': line_no,
                    'Joint No.': joint_no,
                    'Size': size,
                    'Th\'k(mm)': "",
                    'Acc': "",
                    'Rej': "",
                    'Start': "",
                    'End': "",
                    'Length(mm)': "",
                    'Upper': "",
                    'Lower': "",
                    'Height(mm)': "",
                    'Type of Flaw': "",
                    'a/l': "", 'a/t': "",
                    'Evaluation': "",
                    'Remarks': ""
                }
                
                item_full = dict(item)
                for k in self.paut_column_keys:
                    if k not in item_full and k != "selected":
                        item_full[k] = ""
                self.paut_extracted_data.append(item_full)
                
            self.log(f"✅ 사진 목록에서 PAUT 데이터 생성 완료: {len(self.paut_extracted_data)} 건")
            getattr(self, '_last_auto_material_sig', {}).pop('PAUT', None)
            self.populate_preview(self.paut_extracted_data, mode="PAUT")
            self.update_date_listbox("PAUT")
            
        except Exception as e:
            self.log(f"❌ 사진에서 PAUT 데이터 추출 중 오류: {e}")
            messagebox.showerror("오류", f"오류가 발생했습니다:\n{e}")

    def _extract_paut_data(self):
        file_path = self.paut_target_file_path.get()
        if not file_path:
            messagebox.showwarning("파일 미선택", "PAUT 데이터 엑셀 파일을 먼저 선택해주세요.")
            return
            
        try:
            self.log(f"📂 PAUT 데이터 불러오는 중: {os.path.basename(file_path)}")
            self.progress['value'] = 20
            
            df = pd.read_excel(file_path)
            cols = df.columns.tolist()
            
            mapping = {
                "date": next((c for c in cols if any(x in str(c).upper() for x in ["DATE", "일자", "날짜"])), None),
                "t": next((c for c in cols if any(x in c.upper() for x in ["THK", "두께", "TH'K", "THICKNESS"])), None),
                "h": next((c for c in cols if any(x in str(c).upper() for x in ["HEIGHT", "높이", "FLAW_HEIGHT"]) or str(c).strip().upper() == "H"), None),
                # [FIX] Prevent 'l' from matching '검사길이' or 'PAUT길이' (Tested Length)
                "l": next((c for c in cols if (any(x in str(c).upper() for x in ["LENGTH", "길이", "FLAW_LENGTH", "LENTH"]) or str(c).strip().upper() == "L") and not any(exclude in str(c).upper() for exclude in ["검사", "PAUT"])), None),
                "d": next((c for c in cols if any(x in str(c).upper() for x in ["DEPTH", "깊이", "FLAW_DEPTH"]) or str(c).strip().upper() == "D"), None),
                "nature": next((c for c in cols if any(x in str(c).upper() for x in ["NAT", "종류", "FLAW_NATURE", "TYPE"])), None),
                "line": next((c for c in cols if any(x in c.upper() for x in ["LINE", "ISO", "DWG", "DRAWING"])), None),
                "joint": next((c for c in cols if any(x in c.upper() for x in ["JOINT", "WELD"])), None),
                "start": next((c for c in cols if any(x in c.upper() for x in ["START", "L1"])), None),
                "end": next((c for c in cols if any(x in c.upper() for x in ["END", "L2"])), None),
                "upper": next((c for c in cols if any(x in c.upper() for x in ["UPPER", "UP", "DEPTH"])), None),
                "lower": next((c for c in cols if any(x in c.upper() for x in ["LOWER", "LOW"])), None),
                "acc": next((c for c in cols if any(x in str(c).upper() for x in ["ACC", "합격"])), None),
                "rej": next((c for c in cols if any(x in str(c).upper() for x in ["REJ", "불합격"])), None),
                "eval": next((c for c in cols if any(x in str(c).upper() for x in ["합부", "EVAL", "판정"])), None),
                "welder": next((c for c in cols if any(x in str(c).upper() for x in ["용접사", "WELDER"])), None),
                "tested_len": next((c for c in cols if any(x in str(c).upper().replace('\n', '').replace(' ', '') for x in ["검사길이", "PAUT길이", "TESTED", "TESTLEN"])), None),
                "size": next((c for c in cols if any(x in str(c).upper() for x in ["SIZE", "크기", "관경"])), None),
                "remarks": next((c for c in cols if any(x in c.upper() for x in ["REMARK", "REMAKER", "비고"])), None)
            }
            
            if not all([mapping["t"], mapping["h"], mapping["l"], mapping["d"]]):
                self.log("⚠️ 필수 컬럼을 모두 찾을 수 없어 자동 매핑에 일부 실패했습니다.")
                
            self.paut_extracted_data = []
            for _, row in df.iterrows():
                acc_raw = str(row.get(mapping["acc"], "")) if mapping["acc"] else ""
                rej_raw = str(row.get(mapping["rej"], "")) if mapping["rej"] else ""
                eval_raw = str(row.get(mapping["eval"], "")) if mapping["eval"] else ""
                
                # Derive ACC/REJ from '합부' if explicit ACC/REJ columns are missing
                if not acc_raw and not rej_raw and eval_raw:
                    if "합격" in eval_raw and "불합격" not in eval_raw:
                        acc_raw = "√"
                    elif "불합격" in eval_raw or "REJ" in eval_raw.upper() or "NG" in eval_raw.upper():
                        rej_raw = "√"
                    elif "ACC" in eval_raw.upper() or "OK" in eval_raw.upper():
                        acc_raw = "√"

                line_val = str(row.get(mapping["line"], "")) if mapping["line"] else ""
                if line_val.lower() == "nan" or not line_val.strip():
                    continue

                # [NEW] Format Date cleanly (remove time)
                raw_date = row.get(mapping["date"], "") if mapping["date"] else ""
                date_str = ""
                if pd.notna(raw_date) and raw_date:
                    try:
                        date_str = pd.to_datetime(raw_date).strftime('%Y-%m-%d')
                    except:
                        date_str = str(raw_date).split(' ')[0]

                # [NEW] Format Size (Append 'A' if missing)
                size_val = str(row.get(mapping["size"], "")) if mapping["size"] else ""
                size_val = str(size_val).replace('.0', '').strip() if pd.notna(size_val) else ""
                
                # [NEW] Auto-map Thickness based on Size if Th'k is missing
                thk_val = str(row.get(mapping["t"], "")) if mapping["t"] else ""
                if not thk_val or str(thk_val).lower() == "nan":
                    # Official GS Power Standard
                    thk_mapping = {
                        "1100": "11.1", "1000": "11.1",
                        "900": "10.3", "850": "10.3",
                        "800": "9.5",
                        "750": "8.7", "700": "8.7", "650": "8.7",
                        "600": "9.5", "550": "9.5",
                        "500": "6.4", "450": "6.4", "400": "6.4", "350": "6.4", "300": "6.4", "250": "6.4", "200": "6.4"
                    }
                    lookup_key = size_val.replace("A", "").strip()
                    thk_val = thk_mapping.get(lookup_key, "")
                
                if size_val and size_val[-1].isdigit():
                    size_val += "A"

                # [NEW] Format Joint No (Pad single digits to 2 digits, handling suffixes like R3-1 -> R03-1)
                import re
                joint_val = str(row.get(mapping["joint"], "")) if mapping["joint"] else ""
                joint_val = joint_val.replace('.0', '').strip() if pd.notna(joint_val) else ""
                joint_val = re.sub(r'^([A-Za-z]*)(\d)(?!\d)', r'\g<1>0\2', joint_val)

                item = {
                    'selected': True,
                    'Date': date_str,
                    'Line No.': line_val,
                    'Joint No.': joint_val,
                    'Size': size_val,
                    'Th\'k(mm)': thk_val,
                    'Acc': acc_raw,
                    'Rej': rej_raw,
                    'Start': str(row.get(mapping["start"], "")) if mapping["start"] else "",
                    'End': str(row.get(mapping["end"], "")) if mapping["end"] else "",
                    'Length(mm)': str(row.get(mapping["l"], "")) if mapping["l"] else "",
                    'Upper': str(row.get(mapping["upper"], "")) if mapping["upper"] else "",
                    'Lower': str(row.get(mapping["lower"], "")) if mapping["lower"] else "",
                    'Height(mm)': str(row.get(mapping["h"], "")) if mapping["h"] else "",
                    'Type of Flaw': "" if acc_raw == "√" else (str(row.get(mapping["nature"], "Slag")) if mapping["nature"] else "Slag"),
                    'a/l': "", 'a/t': "",
                    'Welder': self._normalize_daily_photo_welder(
                        row.get(mapping["welder"], "")
                    ) if mapping["welder"] else "",
                    'Tested Length': str(row.get(mapping["tested_len"], "")) if mapping["tested_len"] else "",
                    'Evaluation': eval_raw,
                    'Remarks': str(row.get(mapping["remarks"], "")) if mapping["remarks"] else ""
                }
                # [NEW] 원본 엑셀의 모든 컬럼 데이터를 유지하여 '컬럼 관리'에서 활용 가능하게 함
                item_full = row.to_dict()
                item_full.update(item)
                from site_apps.central.src.reports.paut_fields import normalize_tcg
                normalize_tcg(item_full)
                # [NEW] 수동 추가된 컬럼(Remark 등)이 추출 시 증발하지 않도록 초기화
                for k in self.paut_column_keys:
                    if k not in item_full and k != "selected":
                        item_full[k] = ""
                self.paut_extracted_data.append(item_full)
            
            self.progress['value'] = 100
            getattr(self, '_last_auto_material_sig', {}).pop('PAUT', None)
            self.sort_by_column("ISO/DWG", mode="PAUT") # Auto sort for PAUT
            self.update_date_listbox("PAUT") # [FIX] Populate date filter
            self.log(f"✅ PAUT 데이터 추출 완료: {len(self.paut_extracted_data)} 건")
            
        except Exception as e:
            self.log(f"❌ PAUT 데이터 추출 오류: {e}")

    def _run_batch_paut_eval(self):
        if not self.paut_extracted_data:
            messagebox.showwarning("데이터 없음", "먼저 데이터를 추출해주세요.")
            return
            
        self.log("🚀 PAUT 일괄 판정 시작...")
        count_ok, count_ng = 0, 0
        
        for item in self.paut_extracted_data:
            t = self.to_float(item.get('Th\'k(mm)', 0))
            h = self.to_float(item.get('Height(mm)', 0))
            l = self.to_float(item.get('Length(mm)', 0))
            d = self.to_float(item.get('Upper', 0))
            nat = item.get('Type of Flaw', 'Slag')
            eval_mode = self.paut_eval_mode.get()
            
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
            item['Evaluation'] = f"{res} ({loc})"
            # a/l, a/t calculation
            if l > 0 and h > 0: item['a/l'] = f"{h/l:.3f}"
            else: item['a/l'] = ""
            if t > 0 and h > 0: item['a/t'] = f"{h/t:.3f}"
            else: item['a/t'] = ""
            
            if res == "Accept": 
                count_ok += 1
                # [NEW] 합격 시 결함 정보 지우기
                item['Type of Flaw'] = ""
                item['a/l'] = ""
                item['a/t'] = ""
            else: 
                count_ng += 1
            
        self.populate_preview(self.paut_extracted_data, mode="PAUT")
        self.log(f"✅ 판정 완료: 총 {len(self.paut_extracted_data)} 건 (합격: {count_ok}, 불합격: {count_ng})")

    def _generate_paut_report(self):
        # PAUT 기준 성적서 값은 실행 중 메모리에 남아 있는 이전 설정보다
        # 우선한다. save_settings() 전에 UI 변수까지 갱신하여 설정 파일이
        # 다시 예전 값으로 덮어써지지 않게 한다.
        reference_settings = {
            'probe': '5L64-A2 / 5L16A10',
            'wedge': 'SA2-N55S-IHC / SA10-N55S',
            'gain': '11 + 6dB / 21.4 + 6dB',
            'offset': '±17 mm',
            'tcg1_p1': '0', 'tcg1_g1': '4.2dB',
            'tcg1_p2': '8.96mm', 'tcg1_g2': '4.5dB',
            'tcg1_p3': '17.13mm', 'tcg1_g3': '7.0dB',
            'tcg1_p4': '24.79mm', 'tcg1_g4': '12.7dB',
            'tcg2_p1': '0', 'tcg2_g1': '3.2dB',
            'tcg2_p2': '7.61mm', 'tcg2_g2': '2.7dB',
            'tcg2_p3': '15.35mm', 'tcg2_g3': '5.3dB',
            'tcg2_p4': '23.17mm', 'tcg2_g4': '8.1dB',
            'set1_qty': '16', 'set1_first': '1', 'set1_last': '16',
            'set2_qty': '16', 'set2_first': '49', 'set2_last': '64',
        }
        if hasattr(self, 'paut_equip_vars'):
            for key, value in reference_settings.items():
                if key in self.paut_equip_vars:
                    self.paut_equip_vars[key].set(value)
                self.config[f'PAUT_EQUIP_{key.upper()}'] = value

        self.save_settings() # Ensure UI -> config sync
        template_path = self.paut_template_file_path.get()
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("파일 미선택", "PAUT 성적서 양식 파일을 선택해주세요.")
            return

        final_list = [
            d for d in self.paut_extracted_data
            if d.get('selected', True)
            and d.get('date_filtered', True)
            and d.get('line_filtered', True)
            and d.get('size_filtered', True)
        ]
        if not final_list:
            messagebox.showwarning("항목 미선택", "선택된 데이터가 없습니다.")
            return

        contractors = {str(item.get('_source_contractor', '')).strip() for item in final_list
                       if str(item.get('_source_contractor', '')).strip()}
        if len(contractors) > 1:
            messagebox.showwarning('시공사 확인', '여러 시공사가 포함되어 있습니다. 같은 시공사 자료만 선택해 주세요.\n' + ', '.join(sorted(contractors)))
            return
        if contractors:
            self.paut_contractor.set(next(iter(contractors)))
            self.config['PAUT_CONTRACTOR'] = self.paut_contractor.get()
            self.save_settings()
        self.log(f"🚀 PAUT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        
        try:
            # keep_vba=True forces openpyxl to write a macro-enabled workbook
            # content type.  When the selected template is a normal .xlsx file,
            # saving that workbook with an .xlsx extension makes Excel reject it
            # as an extension/format mismatch.
            template_ext = os.path.splitext(template_path)[1].lower()
            keep_vba = template_ext == '.xlsm'
            wb = openpyxl.load_workbook(template_path, keep_vba=keep_vba)
            
            # [NEW] Find "을" sheet for data entry
            target_ws = None
            for sheet in wb.worksheets:
                if "을" in sheet.title:
                    target_ws = sheet
                    break
                    
            if not target_ws:
                target_ws = wb.worksheets[0]
                
            ws = target_ws
            ws.title = f"PAUT_Report_001"
            self._prepare_paut_data_stamp(ws)
            
            # [NEW] Write Cover (Gapji) metadata if there is a cover sheet
            if len(wb.worksheets) > 1:
                first_item = final_list[0] if final_list else None
                ws0 = wb.worksheets[0]
                self._write_gapji_metadata(ws0, mode="PAUT", first_item=first_item)

                # 리포트 기본정보의 검사일자를 우선 사용한다. 입력값이 없을 때만
                # 선택 데이터의 실제 검사기간을 대신 표시한다.
                if not self.gapji_exam_date.get().strip():
                    selected_dates = []
                    for report_item in final_list:
                        date_text = str(report_item.get("Date", "")).strip()
                        date_match = re.search(r"\d{4}-\d{2}-\d{2}", date_text)
                        if date_match:
                            selected_dates.append(date_match.group())
                    if selected_dates:
                        start_date = min(selected_dates)
                        end_date = max(selected_dates)
                        exam_period = (
                            start_date if start_date == end_date
                            else f"{start_date}~{end_date}"
                        )
                        self.safe_set_value(ws0, "AA6", exam_period)

                if not str(ws0["AA5"].value or "").strip():
                    self.safe_set_value(ws0, "AA5", "SIT/GI-JY-SRJ-PAUT-001")
                self._prepare_paut_cover_identity(ws0)
                
                # 갑지의 AA4, AA5, AA6 내용을 을지의 Y4, Y5, Y6에 복사
                try:
                    self.safe_set_value(ws, 'Y4', ws0['AA4'].value)
                    self.safe_set_value(ws, 'Y5', ws0['AA5'].value)
                    self.safe_set_value(ws, 'Y6', ws0['AA6'].value)
                except Exception as e:
                    self.log(f"[WARNING] 갑지 데이터 복사 실패: {e}")
            
            self.force_print_settings(ws, context="DATA")
            
            start_row = int(self.config.get('PAUT_START_ROW', 11))
            end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
            
            curr_row = start_row
            data_font = Font(size=9)
            
            try:
                active_tab_id = self.paut_col_nb.select()
                active_template = self.paut_col_nb.tab(active_tab_id, "text")
            except:
                active_template = "지역난방"

            for i, item in enumerate(final_list):
                if curr_row > end_row:
                    ws = self.prepare_next_paut_sheet(wb, ws.title, i//(end_row-start_row+1))
                    curr_row = start_row
                
                def _get_col(key, default):
                    localized_key = key.replace("PAUT_", f"PAUT_{active_template}_")
                    val = str(self.config.get(localized_key, '')).strip()
                    if not val:
                        val = str(self.config.get(key, '')).strip()
                    return val if val else default

                # [NEW] Combine ACC and REJ to prevent overwriting if they map to the same column
                acc_val = 'ACC' if item.get('Acc', '') in ['√', 'ACC'] else ''
                rej_val = 'REJ' if item.get('Rej', '') in ['√', 'REJ'] else ''
                acc_rej_combined = acc_val if acc_val else rej_val

                # [NEW] Check if there is no defect
                from site_apps.central.src.reports.paut_fields import normalize_tcg
                normalize_tcg(item)
                is_no_defect = (not str(item.get('Type of Flaw', '')).strip()
                                and not str(item.get('Height(mm)', '')).strip())
                if active_template == '지역난방':
                    from site_apps.central.src.reports.paut_report_layout import prepare_indication_row
                    prepare_indication_row(ws, curr_row, not is_no_defect,
                                           has_tcg=str(item.get('TCG MAX(%)', '')).strip() != '')

                # [NEW] Respect user's UI column configuration while providing standard defaults
                mapping = {
                    _get_col('PAUT_COL_NO', 'A'): i+1, 
                    _get_col('PAUT_COL_LINE', 'B'): item.get('Line No.', ''), 
                    _get_col('PAUT_COL_JOINT', 'C'): item.get('Joint No.', ''),
                    _get_col('PAUT_COL_SIZE', 'D'): item.get('Size', ''),
                    _get_col('PAUT_COL_THK', 'E'): item.get('Th\'k(mm)', ''), 
                    _get_col('PAUT_COL_ACC', 'F'): acc_rej_combined,
                    _get_col('PAUT_COL_START', 'G'): item.get('Start', ''),
                    _get_col('PAUT_COL_END', 'H'): item.get('End', ''),
                    _get_col('PAUT_COL_L', 'I'): item.get('Length(mm)', ''),
                    _get_col('PAUT_COL_UP', 'J'): item.get('Upper', ''), 
                    _get_col('PAUT_COL_H', 'K'): item.get('Height(mm)', ''), 
                    _get_col('PAUT_COL_TCG', 'W' if active_template == '지역난방' else 'L'): item.get('TCG MAX(%)', ''),
                    _get_col('PAUT_COL_WELDER', 'M'): item.get('Welder', ''),
                    _get_col('PAUT_COL_TESTED_LEN', 'N'): item.get('Tested Length', ''),
                    _get_col('PAUT_COL_NAT', 'O'): item.get('Type of Flaw', '') or (
                        item.get('Remarks', '') if active_template == '지역난방' else ''),
                    _get_col('PAUT_COL_EVAL', '0'): item.get('Evaluation', ''),
                    _get_col('PAUT_COL_REM', '0'): item.get('Remarks', '')
                }
                
                # [NEW] Explicitly write "NO RECORDABLE INDICATION" to column N if there is no defect
                if is_no_defect:
                    mapping[self.col_to_num('N')] = "NO RECORDABLE INDICATION"
                
                for col_key, val in mapping.items():
                    try:
                        col_idx = self.col_to_num(col_key)
                        if col_idx < 1: continue # Skip if 0 or invalid
                        
                        cell = ws.cell(row=curr_row, column=col_idx, value=val)
                        cell.font = data_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    except Exception:
                        continue
                
                curr_row += 1
                self.progress['value'] = (i+1) / len(final_list) * 100
                
            # [NEW] Add "- 이 하 여 백 -" at the row following the last data row
            if curr_row <= end_row:
                try:
                    ws.merge_cells(start_row=curr_row, start_column=14, end_row=curr_row, end_column=24)
                    blank_cell = ws.cell(row=curr_row, column=14, value="- 이 하 여 백 -")
                    blank_cell.font = Font(name='맑은 고딕', size=10, bold=True)
                    blank_cell.alignment = Alignment(horizontal='center', vertical='center')
                except Exception as e:
                    self.log(f"이하 여백 삽입 오류: {e}")
                
            # [NEW] Add Tested Length Summary at the end
            total_org_m = 0.0
            total_rep_m = 0.0
            processed_joints = set()
            
            for item in final_list:
                try:
                    joint_no = str(item.get('Joint No.', '')).strip().upper()
                    if not joint_no or joint_no in processed_joints:
                        continue
                    
                    val_str = str(item.get('Tested Length', '0')).replace(',', '').strip()
                    num_match = re.search(r"[\d\.]+", val_str)
                    if num_match:
                        val = float(num_match.group())
                        processed_joints.add(joint_no) # Mark as processed only if we successfully parsed length
                        
                        acc_val = 'ACC' if item.get('Acc', '') in ['√', 'ACC'] else ''
                        rej_val = 'REJ' if item.get('Rej', '') in ['√', 'REJ'] else ''
                        acc_rej_combined = acc_val if acc_val else rej_val
                        
                        # Use ACC/REJ to distinguish Original vs Repair
                        if acc_rej_combined == 'REJ' or '불합격' in str(item.get('Rej', '')):
                            total_rep_m += val
                        else:
                            total_org_m += val
                except:
                    pass
            
            total_m = total_org_m + total_rep_m

            org_val = f"{total_org_m:.4f} M" if total_org_m > 0 else " M"
            rep_val = f"{total_rep_m:.4f} M" if total_rep_m > 0 else " M"
            tot_val = f"{total_m:.4f} M" if total_m > 0 else " M"

            from openpyxl.styles import Border, Side
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            summary_data = [
                ("Original", org_val),
                ("Repair", rep_val),
                ("Total", tot_val)
            ]
            
            aa_col = self.col_to_num('AA')
            ae_col = self.col_to_num('AE')
            
            # [NEW] Check if there is enough space for the 3-row summary.
            # If not, create a new sheet for the summary to prevent overwriting footers.
            if curr_row + 2 > end_row:
                ws = self.prepare_next_paut_sheet(wb, ws.title, len(final_list)//(end_row-start_row+1) + 1)
                curr_row = start_row

            for r_idx, (label, val) in enumerate(summary_data):
                r = curr_row + r_idx
                
                c_label = ws.cell(row=r, column=aa_col, value=label)
                c_label.border = thin_border
                c_label.font = data_font
                c_label.alignment = Alignment(horizontal='center', vertical='center')
                
                c_val = ws.cell(row=r, column=ae_col, value=val)
                c_val.border = thin_border
                c_val.font = data_font
                c_val.alignment = Alignment(horizontal='center', vertical='center')
            
            # [NEW] PAUT Page Numbering Logic (Gapji: AE3, Eulji: AF3)
            photo_pages = 0
            try:
                if hasattr(self, 'paut_photo_pages_var'):
                    photo_pages = int(self.paut_photo_pages_var.get())
            except:
                pass
            
            total_p = len(wb.worksheets) + photo_pages

            # PAUT 보고서 뒤에 붙는 사진대장은 전체 보고서의 다음 번호부터
            # 이어지도록 사진대장 탭의 시작/전체 페이지를 자동 동기화한다.
            # 예: PAUT 2페이지 + 사진대장 3페이지 -> 사진대장 3~5 / 전체 5.
            if photo_pages > 0:
                photo_start = len(wb.worksheets) + 1
                if hasattr(self, "photo_start_page"):
                    self.photo_start_page.set(str(photo_start))
                if hasattr(self, "photo_total_pages"):
                    self.photo_total_pages.set(str(total_p))
                self.log(
                    f"[PAUT] 사진대장 페이지 자동 설정: "
                    f"시작 {photo_start}, 전체 {total_p}"
                )

            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                p_text = f"Page    {page_num}    of    {total_p}"
                
                # If there are multiple sheets, the first one is the Cover (Gapji)
                if p_idx == 0 and len(wb.worksheets) > 1:
                    self.safe_set_value(s, 'AC3', p_text)
                    try: 
                        if s['AC3'].font: s['AC3'].font = s['AC3'].font.copy(name='바탕', size=9, bold=False)
                        else: s['AC3'].font = Font(name='바탕', size=9, bold=False)
                    except: pass

                    # 장비/TCG 및 ISO 시험 등급 기준값은 갑지에만 적용한다.
                    if hasattr(self, 'paut_equip_vars'):
                        for k, (lbl, cell, default_val) in self.paut_equip_map.items():
                            val = self.paut_equip_vars[k].get()
                            self.safe_set_value(s, cell, val)

                    grade_values = {
                        'A38': 'C',
                        'F38': 'A-Technique',
                        'K38': 'Level 3',
                        'R38': '11 + 6dB / 21.4 + 6dB',
                    }
                    for cell_ref, value in grade_values.items():
                        self.safe_set_value(s, cell_ref, value)

                else:
                    self.safe_set_value(s, 'AA3', p_text)
                    try: 
                        if s['AA3'].font: s['AA3'].font = s['AA3'].font.copy(name='바탕', size=9, bold=False)
                        else: s['AA3'].font = Font(name='바탕', size=9, bold=False)
                    except: pass

                # Fix right border for AJ4:AJ45 (Draw solid thin line)
                try:
                    from site_apps.central.src.reports.paut_report_layout import solid_right_border
                    solid_right_border(s)
                except Exception as e:
                    self.log(f"AJ border error: {e}")

                # 갑지 검사자는 박광복으로 표시하고 등록된 서명을 함께 넣는다.
                self._apply_paut_inspector_signature(s)
                self._make_header_stamp_transparent(s)

            # [NEW] 최종 저장 직전 사용자의 행/열 커스텀 설정을 모든 시트에 적용
            from site_apps.central.src.reports.paut_report_layout import align_report_heights, measure_excel_print_widths
            from site_apps.central.src.reports.paut_report_layout import fit_height_heading
            for p_idx, s in enumerate(wb.worksheets):
                ctx = "COVER" if p_idx == 0 else "DATA"
                self.apply_custom_dimensions(s, ctx)
                self.force_print_settings(s, ctx)
                if p_idx > 0:
                    fit_height_heading(s)

            output_ext = '.xlsm' if keep_vba else '.xlsx'
            output_dir = os.path.join(BASE_DIR, "output") if os.path.exists(os.path.join(BASE_DIR, "output")) else BASE_DIR
            out_path = self._report_save_path(output_dir, output_ext, mode='PAUT')
            if not out_path:
                return
            out_name = os.path.basename(out_path)
            
            if not os.path.exists(os.path.dirname(out_path)): os.makedirs(os.path.dirname(out_path))
            wb.save(out_path)
            adjusted_path = None
            adjusted_wb = None
            try:
                widths, heights = measure_excel_print_widths(
                    out_path, wb.sheetnames, include_heights=True)
                # Saving an image-bearing openpyxl workbook closes its image
                # streams. Reload before saving again, and replace atomically.
                adjusted_wb = openpyxl.load_workbook(out_path, keep_vba=keep_vba)
                align_report_heights(adjusted_wb.worksheets, measured_widths=widths,
                                     measured_heights=heights)
                # Width measurement uses a temporary fixed scale to align page
                # heights. Restore PAUT's one-page fit after that calculation so
                # Excel does not leave the last column on a page to the right.
                for page_index, adjusted_sheet in enumerate(adjusted_wb.worksheets):
                    adjusted_context = "COVER" if page_index == 0 else "DATA"
                    self.force_print_settings(adjusted_sheet, adjusted_context)
                fd, adjusted_path = tempfile.mkstemp(
                    prefix='paut_print_', suffix=output_ext,
                    dir=os.path.dirname(out_path))
                os.close(fd)
                adjusted_wb.save(adjusted_path)
                os.replace(adjusted_path, out_path)
                adjusted_path = None
            except Exception as exc:
                self.log(f"[PAUT] 인쇄 높이 정렬 실패 — 보정 전 정상 파일 유지: {exc}")
            finally:
                if adjusted_wb is not None:
                    adjusted_wb.close()
                if adjusted_path and os.path.exists(adjusted_path):
                    os.remove(adjusted_path)
            self._record_report_issue('PAUT', wb, out_path, final_list)
            self.log(f"✅ PAUT 성적서 생성 완료: {out_name}")
            os.startfile(os.path.dirname(out_path))
            
        except Exception as e:
            self.log(f"❌ PAUT 성적서 생성 오류: {e}")

    def _sync_paut_photo_page_numbers(self):
        """PAUT 첨부 장수로 사진대장의 시작/전체 페이지를 미리 계산한다."""
        if "photo_start_page" not in self.__dict__ or "photo_total_pages" not in self.__dict__:
            return
        try:
            raw_pages = self.__dict__["paut_photo_pages_var"].get() if "paut_photo_pages_var" in self.__dict__ else self.config.get("PAUT_PHOTO_PAGES", "0")
            photo_pages = int(str(raw_pages).strip() or "0")
        except (TypeError, ValueError):
            return
        if photo_pages <= 0:
            return

        selected_count = sum(
            1 for item in getattr(self, "paut_extracted_data", [])
            if item.get("selected", True)
            and item.get("date_filtered", True)
            and item.get("line_filtered", True)
            and item.get("size_filtered", True)
        )
        start_row = int(self.config.get("PAUT_START_ROW", 11))
        end_row = int(self.config.get("PAUT_DATA_END_ROW", 40))
        rows_per_page = max(1, end_row - start_row + 1)
        data_pages = max(1, math.ceil(selected_count / rows_per_page))
        paut_pages = 1 + data_pages  # 갑지 1장 + 을지

        self.photo_start_page.set(str(paut_pages + 1))
        self.photo_total_pages.set(str(paut_pages + photo_pages))

    def _export_paut_session(self):
        """[NEW] 현재 PAUT 미리보기 데이터를 CSV로 내보내기"""
        if not self.paut_extracted_data:
            messagebox.showwarning("데이터 없음", "내보낼 데이터가 없습니다.")
            return
        
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="PAUT 세션 저장",
            initialfile=f"PAUT_Session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not path: return

        try:
            df = pd.DataFrame(self.paut_extracted_data)
            # UTF-8-SIG for Excel compatibility
            df.to_csv(path, index=False, encoding='utf-8-sig')
            self.log(f"💾 PAUT 세션 저장 완료: {os.path.basename(path)}")
            messagebox.showinfo("저장 완료", f"현재 세션이 저장되었습니다.\n{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"세션 저장 중 오류 발생: {e}")

    def _import_paut_session(self):
        """[NEW] 저장된 PAUT 세션 CSV 불러오기"""
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="PAUT 세션 불러오기"
        )
        if not path: return

        try:
            df = pd.read_csv(path)
            # Ensure mandatory columns or basic structure
            new_data = df.to_dict('records')
            
            # Simple normalization (handle NaNs from CSV)
            normalized_data = []
            for i, d in enumerate(new_data):
                clean_dict = {}
                for k, v in d.items():
                    if pd.isna(v): clean_dict[k] = ""
                    else: clean_dict[k] = v
                
                # [FIX] Ensure mandatory attributes for UI filtering
                if 'date_filtered' not in clean_dict: clean_dict['date_filtered'] = True
                if 'line_filtered' not in clean_dict: clean_dict['line_filtered'] = True
                if 'size_filtered' not in clean_dict: clean_dict['size_filtered'] = True
                if 'selected' not in clean_dict: clean_dict['selected'] = True
                if 'order_index' not in clean_dict: clean_dict['order_index'] = i
                
                normalized_data.append(clean_dict)
            
            self.paut_extracted_data = normalized_data
            self.file_info_vars['PAUT'].set(f"📂 로드됨: {os.path.basename(path)} (총 {len(normalized_data)}건)")
            self.update_date_listbox("PAUT")
            self.populate_preview(self.paut_extracted_data, mode="PAUT")
            self.log(f"📂 PAUT 세션 불러오기 완료: {os.path.basename(path)} ({len(normalized_data)}건)")
        except Exception as e:
            messagebox.showerror("불러오기 오류", f"세션 불러오기 중 오류 발생: {e}")

    def _reset_paut_manual_inputs(self):
        """[NEW] 개별 판정 입력 필드 초기화 (두께 T 등 공통 설정값 제외)"""
        for k in ['h', 'l', 'd', 'z1', 'z2', 'L1', 'L2', 'D_Upper', 'D_Lower']:
            if k in self.paut_manual_vars:
                self.paut_manual_vars[k].set("")
        self.paut_manual_vars['nature'].set("")
        self.paut_manual_vars['loc'].set("-")
        self.paut_res_label.config(text="결과 대기", fg="black", bg="#f3f4f6")

    def prepare_next_paut_sheet(self, wb, prev_title, page_num):
        source_sheet = wb[prev_title]
        new_sheet = wb.copy_worksheet(source_sheet)
        new_sheet.title = f"PAUT_Report_{page_num+1:03d}"
        
        # [FIX] openpyxl's copy_worksheet does not copy images. Manually copy them to preserve Logos.
        try:
            import io
            import copy
            from openpyxl.drawing.image import Image as OpenpyxlImage
            for idx, img in enumerate(source_sheet._images):
                try:
                    img_bytes = img._data()
                    
                    # Fix original image (since _data() closes its file pointer)
                    fixed_orig = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): fixed_orig.anchor = copy.copy(img.anchor)
                    source_sheet._images[idx] = fixed_orig
                    
                    # Add to new sheet
                    new_img = OpenpyxlImage(io.BytesIO(img_bytes))
                    if hasattr(img, 'anchor'): new_img.anchor = copy.copy(img.anchor)
                    new_sheet.add_image(new_img)
                except Exception as inner_e:
                    self.log(f"개별 로고 복사 오류: {inner_e}")
        except Exception as e:
            self.log(f"PAUT 로고 복사 오류: {e}")
        
        start_row = int(self.config.get('PAUT_START_ROW', 11))
        end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
        for r in range(start_row, end_row + 1):
            for c in range(1, 51): 
                try:
                    new_sheet.cell(row=r, column=c).value = None
                except Exception:
                    pass
        
        return new_sheet

    def _apply_paut_inspector_signature(self, ws):
        """PAUT 갑지의 검사자 이름과 서명을 박광복으로 교체한다."""
        inspector_name = str(ws["D45"].value or "").replace(" ", "")
        if inspector_name not in {"박원준", "박광복"}:
            return False

        signature_path = os.path.join(SRC_ROOT or "", "signs", "박광복.png")
        if not os.path.exists(signature_path):
            self.log(f"[WARNING] 검사자 서명 파일을 찾을 수 없습니다: {signature_path}")
            return False

        ws["D45"] = "박 광 복"
        ws["D45"].alignment = Alignment(
            horizontal="center", vertical="center", shrink_to_fit=True
        )

        # 기존 검사자 서명(I44:K45)의 위치와 크기는 양식 그대로 유지한다.
        inspector_anchor = None
        retained_images = []
        for image in ws._images:
            anchor = getattr(image, "anchor", None)
            marker = getattr(anchor, "_from", None)
            if (
                marker is not None
                and 7 <= marker.col <= 10
                and 43 <= marker.row <= 44
            ):
                inspector_anchor = anchor
                continue
            retained_images.append(image)

        if inspector_anchor is None:
            self.log("[WARNING] 기존 검사자 서명 위치를 찾을 수 없습니다.")
            return False

        import copy
        ws._images = retained_images
        try:
            transparent_stream = self._white_background_to_transparent_png(
                signature_path
            )
            signature = XLImage(transparent_stream)
            # workbook 저장 시까지 BytesIO가 닫히지 않도록 참조를 유지한다.
            signature._pmi_source_stream = transparent_stream
        except Exception as exc:
            self.log(f"[WARNING] 검사자 서명 투명 처리 실패: {exc}")
            signature = XLImage(signature_path)
        signature.anchor = copy.copy(inspector_anchor)
        ws.add_image(signature)
        return True

    def _make_header_stamp_transparent(self, ws):
        """보고서 머리글의 붉은 도장에서 흰 배경을 제거한다."""
        import copy

        replaced_images = []
        changed = False
        for image in ws._images:
            anchor = getattr(image, "anchor", None)
            marker = getattr(anchor, "_from", None)
            is_header_stamp = (
                marker is not None
                and 12 <= marker.col <= 22
                and 0 <= marker.row <= 7
                and 0.65 <= (float(image.width) / max(float(image.height), 1.0)) <= 1.35
            )
            if not is_header_stamp:
                replaced_images.append(image)
                continue

            try:
                source_bytes = image._data()
                with PILImage.open(io.BytesIO(source_bytes)) as source:
                    rgba = source.convert("RGBA")
                    transparent_pixels = []
                    for red, green, blue, alpha in rgba.getdata():
                        # 붉은 인영만 남기고 JPEG 압축으로 생긴 흰색/회색 배경도 제거한다.
                        if red > 100 and red - green > 12 and red - blue > 12:
                            transparent_pixels.append((red, green, blue, alpha))
                        else:
                            transparent_pixels.append((red, green, blue, 0))
                    rgba.putdata(transparent_pixels)
                    output = io.BytesIO()
                    rgba.save(output, format="PNG")
                    output.seek(0)

                stamp = XLImage(output)
                stamp.anchor = copy.copy(anchor)
                replaced_images.append(stamp)
                changed = True
            except Exception as exc:
                self.log(f"[WARNING] 도장 배경 투명 처리 실패: {exc}")
                replaced_images.append(image)

        if changed:
            ws._images = replaced_images
        return changed

    def _get_photo_log_stamp(self):
        """PAUT 양식의 도장을 추출해 투명 PNG 데이터로 반환한다."""
        template_path = self.paut_template_file_path.get().strip()
        if not template_path or not os.path.exists(template_path):
            return None

        template_wb = None
        try:
            template_wb = openpyxl.load_workbook(template_path, keep_vba=True)
            for sheet in template_wb.worksheets:
                for image in sheet._images:
                    anchor = getattr(image, "anchor", None)
                    marker = getattr(anchor, "_from", None)
                    is_header_stamp = (
                        marker is not None
                        and 12 <= marker.col <= 22
                        and 0 <= marker.row <= 7
                        and 0.65 <= (float(image.width) / max(float(image.height), 1.0)) <= 1.35
                    )
                    if not is_header_stamp:
                        continue

                    with PILImage.open(io.BytesIO(image._data())) as source:
                        rgba = source.convert("RGBA")
                        pixels = []
                        for red, green, blue, alpha in rgba.getdata():
                            if red > 100 and red - green > 12 and red - blue > 12:
                                pixels.append((red, green, blue, alpha))
                            else:
                                pixels.append((red, green, blue, 0))
                        rgba.putdata(pixels)
                        output = io.BytesIO()
                        rgba.save(output, format="PNG")
                        return output.getvalue()
        except Exception as exc:
            self.log(f"[PhotoLog] 도장 이미지 준비 실패: {exc}")
        finally:
            if template_wb is not None:
                template_wb.close()
        return None
