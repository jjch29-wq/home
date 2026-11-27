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
from site_apps.central.src.reports.photo_log_manager import PhotoLogMixin
from site_apps.central.src.reports.report_sheet_styler import ReportSheetStylerMixin
from site_apps.central.src.reports.ndt_report_runners import NDTReportRunnersMixin
from site_apps.central.src.reports.report_preview_view import ReportPreviewViewMixin
from site_apps.central.src.reports.report_table_controller import ReportTableControllerMixin

from site_apps.central.src.reports.report_mode_ui_mixin import ReportModeUIMixin
from site_apps.central.src.reports.report_extraction_mixin import ReportExtractionMixin



try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

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
    CONFIG_DIR = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.dirname(sys.executable),
        "PMI", "central", "report_settings",
    )
    LEGACY_SETTINGS_FILE = os.path.join(os.path.dirname(sys.executable), "logo_settings_unified.json")
    RESOURCE_DIR = os.path.join(base_path, "resources")
else:
    SCRIPT_HOME = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = find_project_home(SCRIPT_HOME)
    CONFIG_DIR = os.path.abspath(os.path.join(SCRIPT_HOME, "..", "..", "data", "report_settings"))
    LEGACY_SETTINGS_FILE = os.path.join(BASE_DIR, "config", "logo_settings_unified.json")
    RESOURCE_DIR = os.path.join(BASE_DIR, "resources")

SETTINGS_FILE = os.path.join(CONFIG_DIR, "logo_settings_unified.json")
PHOTO_LOG_CONFIG_FILE = os.path.join(CONFIG_DIR, "photolog_config.json")
LEGACY_PHOTO_LOG_CONFIG_FILE = os.path.join(
    os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else BASE_DIR,
    "photolog_config.json",
)


def _existing_report_settings(primary, *legacy_paths):
    return next((path for path in (primary, *legacy_paths) if os.path.isfile(path)), None)
SRC_ROOT = str(Path(__file__).resolve().parents[4]) if 'Path' in globals() else os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))



# Embedded Logo Fallback
DEFAULT_SITCO_LOGO_B64 = "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"

class PMIReportApp(PhotoLogMixin, ReportSheetStylerMixin, NDTReportRunnersMixin, ReportPreviewViewMixin, ReportTableControllerMixin, ReportModeUIMixin, ReportExtractionMixin):
    def __init__(self, root, central_app=None, is_embedded=False):
        self._is_ready = False
        self.root = root
        self.central_app = central_app
        self.is_embedded = is_embedded
        if not is_embedded and hasattr(self.root, "title"):
            try:
                self.root.title(f"SITCO 통합 성적서 자동 생성기 ({APP_VERSION})")
                self.root.geometry("950x850")
                self.root.configure(background="#f9fafb")
            except Exception:
                pass
        
        # 1. Initialize Configuration first (needed for state variables)
        self.config = {
            # UI State (Persistence)
            'PMI_SASH_RATIO': 0.5, 'RT_SASH_RATIO': 0.5, 'PT_SASH_RATIO': 0.5, 'MT_SASH_RATIO': 0.5, 'PAUT_SASH_RATIO': 0.5,
            'PMI_COL_WIDTHS': {}, 'RT_COL_WIDTHS': {}, 'KOGAS_COL_WIDTHS': {}, 'PT_COL_WIDTHS': {}, 'MT_COL_WIDTHS': {}, 'PAUT_COL_WIDTHS': {},
            'SIDEBAR_HIDDEN': {},
            # 갑지 (Cover)
            'SEOUL_COVER_ANCHOR': "E5", 'SEOUL_COVER_W': 200.0, 'SEOUL_COVER_H': 18.0, 'SEOUL_COVER_X': 30.0, 'SEOUL_COVER_Y': 15.0,
            'SITCO_COVER_ANCHOR': "A6", 'SITCO_COVER_W': 80.0, 'SITCO_COVER_H': 40.0, 'SITCO_COVER_X': 10.0, 'SITCO_COVER_Y': 10.0,
            'FOOTER_COVER_ANCHOR': "P50", 'FOOTER_COVER_W': 80.0, 'FOOTER_COVER_H': 20.0, 'FOOTER_COVER_X': 0.0, 'FOOTER_COVER_Y': 10.0,
            'FOOTER_PT_COVER_ANCHOR': "A50", 'FOOTER_PT_COVER_W': 100.0, 'FOOTER_PT_COVER_H': 25.0, 'FOOTER_PT_COVER_X': 5.0, 'FOOTER_PT_COVER_Y': 10.0,
            # 을지 (Data)
            'SEOUL_DATA_ANCHOR': "F5", 'SEOUL_DATA_W': 200.0, 'SEOUL_DATA_H': 18.0, 'SEOUL_DATA_X': 30.0, 'SEOUL_DATA_Y': 15.0,
            'SITCO_DATA_ANCHOR': "A6", 'SITCO_DATA_W': 100.0, 'SITCO_DATA_H': 50.0, 'SITCO_DATA_X': 20.0, 'SITCO_DATA_Y': 5.0,
            'FOOTER_DATA_ANCHOR': "M46", 'FOOTER_DATA_W': 100.0, 'FOOTER_DATA_H': 15.0, 'FOOTER_DATA_X': 25.0, 'FOOTER_DATA_Y': 10.0,
            'FOOTER_PT_DATA_ANCHOR': "A46", 'FOOTER_PT_DATA_W': 100.0, 'FOOTER_PT_DATA_H': 25.0, 'FOOTER_PT_DATA_X': 3.0, 'FOOTER_PT_DATA_Y': 7.0,
            'MARGIN_DATA_TOP': 0.2, 'MARGIN_DATA_BOTTOM': 0.0, 'MARGIN_DATA_LEFT': 0.5, 'MARGIN_DATA_RIGHT': 0.3, 'PRINT_SCALE_DATA': 95,
            'START_ROW': 19, 'DATA_END_ROW': 45, 'PRINT_END_ROW': 47,
            # [NEW] 선택적 행/열 조절용 설정
            'CUSTOM_ROWS_COVER': '23-38', 'CUSTOM_ROW_HEIGHT_COVER': 21.0,
            'CUSTOM_COLS_COVER': '', 'CUSTOM_COL_WIDTH_COVER': 10.0,
            'CUSTOM_ROWS_DATA': '', 'CUSTOM_ROW_HEIGHT_DATA': 20.55,
            'CUSTOM_COLS_DATA': '', 'CUSTOM_COL_WIDTH_DATA': 10.0,
            
            # PAUT Specific (B31.1)
            'PAUT_START_ROW': 11, 'PAUT_DATA_END_ROW': 40, 'PAUT_PRINT_END_ROW': 45,
            
            # RT Specific (Radiographic Testing)
            'RT_START_ROW': 11, 'RT_DATA_END_ROW': 45, 'RT_PRINT_END_ROW': 47,
            
            # Photo Log Sash
            'PHOTO_SASH_RATIO': 0.45
        }

        # 2. State Variables
        self.setting_vars = {} # [CRITICAL FIX] Container for all dynamic UI variables
        self.logo_folder_path = tk.StringVar(value="")
        self.target_file_path = tk.StringVar(value=self.config.get('PMI_TARGET_PATH', ""))
        pmi_tmpl = self.config.get('PMI_TEMPLATE_PATH', "")
        if pmi_tmpl and not os.path.exists(pmi_tmpl):
            for candidate in ("home/data/SIT-PMI-001.xlsx", "home/data/SIT-PMI-K3 롯데 템플릿보고서.xlsm"):
                full_c = os.path.abspath(candidate)
                if os.path.exists(full_c):
                    pmi_tmpl = full_c
                    break
        self.template_file_path = tk.StringVar(value=pmi_tmpl)
        self.sequence_filter = tk.StringVar()
        self.rt_extract_keyword = tk.StringVar()  # [NEW] 추출 키워드 필터
        self.extraction_mode = tk.StringVar(value="전체")
        self.auto_verify = tk.BooleanVar(value=True)
        self.pmi_layout_mode = tk.StringVar(value=self.config.get('PMI_LAYOUT_MODE', 'STANDARD'))
        self.pmi_pane_ratio = self.config.get('PMI_SASH_RATIO', 0.5)
        self.current_mode = "PMI" # [NEW] Track active report mode
        
        self.show_selected_only = tk.BooleanVar(value=False)
        self.extracted_data = [] 
        self.pmi_sort_col = "" 
        self.pmi_sort_rev = False 
        self.rt_sort_col = ""
        self.rt_sort_rev = False
        self.pt_sort_col = ""
        self.mt_sort_col = ""
        self.pt_sort_rev = False
        self.mt_sort_rev = False
        self.paut_sort_col = ""
        self.paut_sort_rev = False
        
        # [NEW] PMI Preview Search & Filter
        self.pmi_search_loc = tk.StringVar()
        self.pmi_show_deficiency_only = tk.BooleanVar(value=False)
        
        # [NEW] 원소 함량 필터링용 상태 변수 (PMI 전용)
        self.element_filters = [] # list of dict: {'key': StringVar, 'min': StringVar, 'max': StringVar}
        # 기본 필터 추가 (Cr, Ni, Mo)
        for k in ["Cr", "Ni", "Mo"]:
            self.element_filters.append({
                'key': tk.StringVar(value=k), 
                'min': tk.StringVar(value=''), 
                'max': tk.StringVar(value='')
            })
            
        # --- PAUT State Variables ---
        self.paut_target_file_path = tk.StringVar(value=self.config.get('PAUT_TARGET_PATH', ""))
        
        default_paut_path = os.path.join(os.path.expanduser("~"), "Desktop", "지역난방 PAUT.xlsx")
        saved_paut_path = self.config.get('PAUT_TEMPLATE_PATH', "")
        if not saved_paut_path and os.path.exists(default_paut_path):
            saved_paut_path = default_paut_path
            
        self.paut_template_file_path = tk.StringVar(value=saved_paut_path)
        self.paut_manual_vars = {
            't': tk.StringVar(), 'h': tk.StringVar(), 'l': tk.StringVar(), 'd': tk.StringVar(),
            'nature': tk.StringVar(value=""), 'loc': tk.StringVar(value="-"),
            'db': tk.StringVar(value="6"), 'peak': tk.StringVar(value="80"), 
            'z1': tk.StringVar(), 'z2': tk.StringVar(), 
            'L1': tk.StringVar(), 'L2': tk.StringVar(),
            'D_Upper': tk.StringVar(), 'D_Lower': tk.StringVar(),
            'target_fsh': tk.StringVar(value="40.0%")
        }
        self.paut_eval_mode = tk.StringVar(value="ECA") # [NEW] Standard / ECA
        self.paut_manual_vars['peak'].trace_add("write", lambda *args: self._update_paut_target_fsh())
        self.paut_manual_vars['db'].trace_add("write", lambda *args: self._update_paut_target_fsh())
        self.paut_extracted_data = []
        
        self.rt_target_file_path = tk.StringVar(value=self.config.get('RT_TARGET_PATH', ""))
        self.rt_template_file_path = tk.StringVar(value=self.config.get('RT_TEMPLATE_PATH', ""))
        self.kogas_target_file_path = tk.StringVar(value=self.config.get('KOGAS_TARGET_PATH', ""))
        self.kogas_template_file_path = tk.StringVar(value=self.config.get('KOGAS_TEMPLATE_PATH', ""))
        self.rt_kogas_mode = tk.BooleanVar(value=False) 
        self.rt_kogas_mode.trace_add("write", lambda *a: self._on_rt_kogas_mode_change())
        self.extracted_data = []
        self.rt_extracted_data = []
        self.kogas_extracted_data = []
        
        # --- Column Keys Initialization ---
        self.column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Result"]
        self.rt_column_keys = ["selected", "No", "Date", "Report No", "Sec", "Dwg", "Joint", "Loc", "T", "Mat", "Size", "Deg", "Acc", "Rej", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "Result", "Welder", "Remarks"]
        self.kogas_column_keys = list(self.rt_column_keys)
        self.pt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "T", "Mat", "Deg", "Result", "Welder", "Remarks"]
        self.mt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "T", "Mat", "Deg", "Result", "Welder", "Remarks"]
        self.paut_column_keys = ["selected", "No", "Date", "Report No", "ISO", "Joint", "Size", "Loc", "T", "Acc", "Rej", "Mat", "Grade", "Nature", "Type", "a/l", "a/t", "Evaluation", "Remarks"]
        
        # --- PT State Variables ---
        self.pt_target_file_path = tk.StringVar(value=self.config.get('PT_TARGET_PATH', ""))
        self.pt_template_file_path = tk.StringVar(value=self.config.get('PT_TEMPLATE_PATH', ""))
        self.pt_extracted_data = []
        self.item_idx_map = []      # [NEW] Missing for PMI
        self.rt_item_idx_map = []   # [NEW] Missing for RT
        self.pt_item_idx_map = []
        
        self.mt_target_file_path = tk.StringVar(value=self.config.get('MT_TARGET_PATH', ""))
        self.mt_template_file_path = tk.StringVar(value=self.config.get('MT_TEMPLATE_PATH', ""))
        self.mt_extracted_data = []
        self.mt_item_idx_map = []
        
        self.paut_item_idx_map = [] 
        self.kogas_item_idx_map = []
        
        # [REFINED] Column Keys Mapping (Must match Treeview column count and order)
        self.column_keys = ["_status", "selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Grade"]
        self.rt_column_keys = ["selected", "No", "Date", "Report No", "Sec", "Dwg", "Joint", "Loc", "T", "Mat", "Size", "Acc", "Rej", "Deg", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "Welder", "Remarks"]
        self.kogas_column_keys = list(self.rt_column_keys)
        self.pt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "NPS", "Thk.", "Material", "TestItem", "Result", "Welder", "Remarks"]
        self.mt_column_keys = ["selected", "No", "Date", "Report No", "Dwg", "Joint", "Material", "TestItem", "Result", "Welder", "Remarks"]
        self.paut_column_keys = ["selected", "No", "Date", "Report No", "Line No.", "Joint No.", "Size", "Th'k(mm)", "Acc", "Rej", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "Type of Flaw", "a/l", "a/t", "Welder", "Tested Length", "Evaluation", "Remarks"]
        
        self.date_listbox = None
    
        # [NEW] Handle Application Closing for Final State Capture
        if not self.is_embedded and hasattr(self.root, "protocol"):
            try: self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            except Exception: pass
        self.rt_date_listbox = None
        self.mt_date_listbox = None
        self.paut_date_listbox = None
        self.paut_line_listbox = None
        self.paut_size_listbox = None
        self.kogas_date_listbox = None

        # [NEW] Add traces for Template-Linked Config Auto-Load
        self.rt_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.rt_template_file_path.get(), "RT"))
        self.kogas_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.kogas_template_file_path.get(), "KOGAS"))
        self.pt_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.pt_template_file_path.get(), "PT"))
        self.mt_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.mt_template_file_path.get(), "MT"))
        self.paut_template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.paut_template_file_path.get(), "PAUT"))
        self.template_file_path.trace_add("write", lambda *args: self.load_template_specific_config(self.template_file_path.get(), "PMI"))

        # Keys are now verified to match Treeview column indices 1-N.

        # 3. UI Initialization
        # [NEW] File Info Strings for Preview Headers
        self.file_info_vars = {
            'PMI': tk.StringVar(value="📄 파일을 선택해주세요."),
            'RT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'KOGAS': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'MT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PAUT': tk.StringVar(value="📄 파일을 선택해주세요."),
            'PHOTO': tk.StringVar(value="📸 사진 리스트를 구성해주세요.")
        }

        # --- Gapji (Cover) Metadata Variables ---
        self.gapji_project = tk.StringVar(value=self.config.get(
            'GAPJI_PROJECT', "2026년 중앙지사 열수송관 비파괴검사 단가계약"
        ))
        self.gapji_customer = tk.StringVar(value=self.config.get(
            'GAPJI_CUSTOMER', "한국지역난방공사 중앙지사"
        ))
        self.gapji_item = tk.StringVar(value=self.config.get('GAPJI_ITEM', "PIPE"))
        self.gapji_material = tk.StringVar(value=self.config.get('GAPJI_MATERIAL', ""))
        self.paut_sec = tk.StringVar(value=self.config.get('PAUT_SEC', ''))
        self.paut_contractor = tk.StringVar(value=self.config.get('PAUT_CONTRACTOR', 'GS네오텍'))
        self.mt_surface_temp = tk.StringVar(value=self.config.get('MT_SURFACE_TEMP', ''))
        self.gapji_report_no = tk.StringVar(value=self.config.get('GAPJI_REPORT_NO', ""))
        self.gapji_exam_date = tk.StringVar(value=self.config.get('GAPJI_EXAM_DATE', datetime.datetime.now().strftime("%Y-%m-%d")))
        self.pmi_chem_elements = [
            tk.StringVar(value=self.config.get(f'PMI_CHEM_ELEMENT_{idx}', default))
            for idx, default in enumerate(("Ni", "Cr", "Mo", "-", "-"), start=1)
        ]
        self.pmi_chem_specs = [
            tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{idx}', ""))
            for idx in (1, 2, 3)
        ]
        self.pmi_spec1_mat = tk.StringVar(value=self.config.get('PMI_SPEC1_MAT', ''))
        self.pmi_spec2_mat = tk.StringVar(value=self.config.get('PMI_SPEC2_MAT', ''))
        self.pmi_spec3_mat = tk.StringVar(value=self.config.get('PMI_SPEC3_MAT', ''))
        self.rt_isotope = tk.StringVar(value=self.config.get('RT_ISOTOPE', 'Ir-192'))
        self.rt_activity = tk.StringVar(value=self.config.get('RT_ACTIVITY', '17'))
        self.rt_activity_unit = tk.StringVar(value=self.config.get('RT_ACTIVITY_UNIT', 'Ci'))
        self.rt_film_brand = tk.StringVar(value=self.config.get('RT_FILM_BRAND', 'FUJI#80'))
        self.rt_film_type = tk.StringVar(value=self.config.get('RT_FILM_TYPE', 'I'))
        self.rt_film_size = tk.StringVar(value=self.config.get('RT_FILM_SIZE', '3⅓ X 12'))
        self.rt_exposure_time = tk.StringVar(value=self.config.get('RT_EXPOSURE_TIME', '2 MIN 40 SEC'))
        
        # [NEW] Trace Gapji metadata for real-time preview updates
        for var in [self.gapji_project, self.gapji_customer, self.gapji_item, 
                    self.gapji_material, self.gapji_report_no, self.gapji_exam_date,
                    self.rt_isotope, self.rt_activity, self.rt_activity_unit,
                    self.rt_film_brand, self.rt_film_type, self.rt_film_size, self.rt_exposure_time]:
            var.trace_add("write", lambda *a: self._update_gapji_preview(getattr(self, 'current_mode', 'PMI')))

        # --- Photo Log State Variables ---
        self.photo_header_map = {
            "PAUT": "REPORT OF PHASED ARRAY UT EXAMINATION (위 상 배 열 초 음 파 탐 상 검 사 보 고 서)",
            "RT": "REPORT OF RADIOGRAPHIC EXAMINATION (방 사 선 투 과 검 사 보 고 서)",
            "PT": "REPORT OF LIQUID PENETRANT EXAMINATION (침 투 탐 상 검 사 보 고 서)",
            "MT": "REPORT OF MAGNETIC PARTICLE EXAMINATION (자 분 탐 상 검 사 보 고 서)",
            "PMI": "REPORT OF POSITIVE MATERIAL IDENTIFICATION (재 질 성 분 분 석 검 사 보 고 서)",
            "UT": "REPORT OF ULTRASONIC EXAMINATION (초 음 파 탐 상 검 사 보 고 서)",
            "NDT": "REPORT OF NON-DESTRUCTIVE EXAMINATION (비 파 괴 검 사 보 고 서)",
            "기타 (직접 입력)": ""
        }
        self.photo_orderer = tk.StringVar(value="서울에너지공사")
        self.photo_inspect_date = tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d"))
        self.photo_report_no = tk.StringVar(value="SIT/GI-SE-PAUT-TNTFJPWJ001")
        self.photo_inspect_type = tk.StringVar(value="PAUT")
        self.photo_report_title = tk.StringVar(value=self.photo_header_map["PAUT"])
        self.photo_cols_per_row = tk.StringVar(value="2")
        self.photo_rows_per_page = tk.StringVar(value="4")
        self.photo_keep_aspect = tk.BooleanVar(value=True)
        self.photo_output_name = tk.StringVar(value="NDT_Photo_Log_Final.xlsx")
        _def_logo = os.path.join(RESOURCE_DIR, "logo.png")
        self.photo_logo_path = tk.StringVar(value=_def_logo if os.path.exists(_def_logo) else "")
        self.photo_logo_width_var = tk.StringVar(value="80")
        self.photo_logo_height_var = tk.StringVar(value="") # [NEW] Custom Logo Height
        self.photo_logo_x_var = tk.StringVar(value="2")
        self.photo_logo_y_var = tk.StringVar(value="0")
        self.photo_cell_width_var = tk.StringVar(value="53.0")
        self.photo_cell_height_var = tk.StringVar(value="178.0")
        self.photo_margin_top_var = tk.StringVar(value="0.5")
        self.photo_margin_bottom_var = tk.StringVar(value="0.5")
        self.photo_margin_left_var = tk.StringVar(value="0.4")
        self.photo_margin_right_var = tk.StringVar(value="0.4")
        self.photo_print_scale_var = tk.StringVar(value="100")
        self.photo_desc_height_var = tk.StringVar(value="20.0")
        self.photo_align_var = tk.StringVar(value="중앙 정렬")
        self.photo_fit_width_var = tk.BooleanVar(value=False)
        self.photo_auto_fit_page_var = tk.BooleanVar(value=True) # [NEW] Auto-fit to A4 page
        self.photo_auto_rotate_var = tk.BooleanVar(value=False)
        self.photo_width_pct_var = tk.StringVar(value="100.0")
        self.photo_height_pct_var = tk.StringVar(value="140.0")
        self.photo_width_pixel_adj_var = tk.StringVar(value="0")
        self.photo_height_pixel_adj_var = tk.StringVar(value="10") # [NEW] Height padding inside cell
        self.photo_shift_x_var = tk.StringVar(value="0")
        self.photo_shift_y_var = tk.StringVar(value="0")
        self.photo_dpi_var = tk.StringVar(value="96") # [NEW] User Configurable DPI
        self.photo_selected_files = []
        self.photo_item_metadata = {}

        # 검사방법별 기본정보를 저장하고 사진대장 정보에도 즉시 반영한다.
        for var in (
            self.gapji_project, self.gapji_customer, self.gapji_item,
            self.gapji_material, self.gapji_report_no, self.gapji_exam_date,
        ):
            var.trace_add("write", lambda *_args: self._on_report_info_changed())
        
        self.load_settings()
        self._initialize_report_info_by_mode()
        self._load_report_info_for_mode("PMI")
        self._sync_photo_report_info()
        self.photo_report_no.trace_add('write', lambda *_args: self._sync_photo_output_name())
        self._sync_photo_output_name()
        
        # [MIGRATION] Add missing keys from loaded config
        if "Size" not in self.paut_column_keys:
            try:
                idx = self.paut_column_keys.index("Joint No.")
                self.paut_column_keys.insert(idx + 1, "Size")
            except:
                self.paut_column_keys.append("Size")
        if "Acc" not in self.paut_column_keys:
            try:
                idx = self.paut_column_keys.index("Th'k(mm)")
                self.paut_column_keys.insert(idx + 1, "Acc")
                self.paut_column_keys.insert(idx + 2, "Rej")
            except:
                self.paut_column_keys.extend(["Acc", "Rej"])
                
        self.last_photo_save_dir = self.config.get('PHOTO_LOG_SETTINGS', {}).get('last_save_dir', "")

        self.create_widgets()
        
        # [NEW] Sync initial file info from loaded config
        self._sync_all_file_infos()
        
        # [NEW] Auto-save on exit
        if not self.is_embedded and hasattr(self.root, "protocol"):
            try: self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            except Exception: pass
        
        self._is_ready = True
        self.log("[INFO] 통합 버전을 시작했습니다.")
        
    def _safe_update_scrollregion(self):
        """Safely update canvas scrollregion, handling None bbox."""
        try:
            bbox = self.canvas.bbox("all")
            if bbox:
                self.canvas.configure(scrollregion=bbox)
        except Exception:
            pass

    def log(self, message):
        line = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}"
        if hasattr(self, 'status_log'):
            self.status_log.config(state='normal')
            self.status_log.insert(tk.END, line + '\n')
            self.status_log.see(tk.END)
            self.status_log.config(state='disabled')
        else:
            try:
                print(line)
            except UnicodeEncodeError as error:
                # Keep the log readable when stdout cannot encode an emoji or filename.
                print(line.encode(error.encoding, errors='backslashreplace').decode(error.encoding))

    def fix_material_name(self, t):
        """재질명을 표준 포맷으로 변환"""
        t_str = str(t).upper()
        if t_str == 'NAN': return ""
        t_str = t_str.replace('A312-TP304', 'S/S').replace('A312-304L', 'S/S').replace('A312-305L', 'S/S').replace('A53-B', 'C/S').replace('A106-B', 'C/S')
        return t_str.replace('C2','C/S').replace('C4','C/S').replace('CS','C/S').replace('S99','S/S').replace('SS','S/S')

    def force_two_digit(self, val):
        """숫자 값을 2자리 문자열로 변환 (예: 1 -> 01)"""
        try:
            s = str(val).strip()
            f = float(s)
            if f.is_integer(): return f"{int(f):02d}"
            return s
        except: return str(val).strip()

    def load_settings(self):
        # 1. 내장 설정 파일 (실행파일 내부에 번들링된 기본값)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            bundle_settings = os.path.join(sys._MEIPASS, "logo_settings_unified.json")
            if os.path.exists(bundle_settings):
                try:
                    with open(bundle_settings, 'r', encoding='utf-8') as f:
                        bundle_data = json.load(f)
                        self.config.update(bundle_data)
                    print("SUCCESS: 내장된 기본 설정을 로드했습니다.")
                except Exception as e:
                    print(f"WARNING: 내장 설정 불러오기 실패: {e}")

        # 2. 외부 설정 파일 (사용자 저장값 - 내장값보다 우선순위 높음)
        settings_path = _existing_report_settings(SETTINGS_FILE, LEGACY_SETTINGS_FILE)
        if settings_path:
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                    
                    # [REFINED] Case-insensitive Sanitization for all modules
                    sanit_map = {
                        'PMI_NAME_DWG': (['ISO/DWG', 'ISO/DWG.'], 'Drawing No.'),
                        'PMI_NAME_JOINT': (['JOINT NO', 'JOINT NO.'], 'Joint No.'),
                        'PMI_NAME_LOC': (['TEST LOCATION', 'LOC'], 'Location'),
                        'PMI_NAME_RES': (['GRADE', 'RESULT'], 'Result'),
                        'RT_NAME_DWG': (['ISO/DWG', 'ISO/DWG.', 'DWG'], 'Drawing No.'),
                        'RT_NAME_JOINT': (['JOINT NO', 'FILM IDENT', 'JOINT'], 'Film Ident. No.'),
                        'RT_NAME_LOC': (['LOC', 'LOC.'], 'Film Location'),
                        'RT_NAME_WELDER': (['WELDER'], 'Welder No'),
                        'RT_NAME_DEG': (['물성', '물성(DEG)', 'DEG'], 'Deg')
                    }
                    for k, (olds, new) in sanit_map.items():
                        if k in saved_data:
                            val = str(saved_data[k]).upper().strip()
                            if any(val == o.upper() for o in olds):
                                saved_data[k] = new

                    # [NEW] Sanitize old "Element-" prefixes and Mn exclusion
                    for k in ['PMI_NAME_NI', 'PMI_NAME_CR', 'PMI_NAME_MO']:
                        if k in saved_data and isinstance(saved_data[k], str) and (saved_data[k].startswith("Element-") or "ELEMENT" in saved_data[k].upper()):
                            saved_data[k] = saved_data[k].replace("Element-", "").replace("ELEMENT-", "").replace("Element", "").strip()
                    
                    if 'column_keys' in saved_data and isinstance(saved_data['column_keys'], list):
                        while 'Mn' in saved_data['column_keys']: saved_data['column_keys'].remove('Mn')

                    # [NEW] Defect Names Sanitization (D1-D15 to circled labels)
                    defect_standards = {
                        1: "① Crack", 2: "② IP", 3: "③ LF", 4: "④ Slag", 5: "⑤ Por",
                        6: "⑥ U/C", 7: "⑦ RUC", 8: "⑧ BT", 9: "⑨ TI", 10: "⑩ CP",
                        11: "⑪ RC", 12: "⑫ Mis", 13: "⑬ EP", 14: "⑭ SD", 15: "⑮ Oth"
                    }
                    for d_i, std_name in defect_standards.items():
                        d_k = f'RT_NAME_D{d_i}'
                        if d_k in saved_data:
                            curr_val = str(saved_data[d_k]).upper().strip()
                            if curr_val == f"D{d_i}" or not curr_val or curr_val == "NAN":
                                saved_data[d_k] = std_name
                    
                    # [NEW] Restore Photo Log Variables
                    photo_vars = {
                        'orderer': self.photo_orderer, 'inspect_date': self.photo_inspect_date,
                        'inspect_type': self.photo_inspect_type, 'report_title': self.photo_report_title,
                        'report_no': self.photo_report_no,
                        'cols_per_row': self.photo_cols_per_row, 'rows_per_page': self.photo_rows_per_page, 'keep_aspect': self.photo_keep_aspect,
                        'output_name': self.photo_output_name, 'logo_path': self.photo_logo_path,
                        'logo_width': self.photo_logo_width_var, 'logo_height': self.photo_logo_height_var,
                        'logo_x': self.photo_logo_x_var, 'logo_y': self.photo_logo_y_var,
                        'cell_width': self.photo_cell_width_var, 'cell_height': self.photo_cell_height_var,
                        'm_top': self.photo_margin_top_var, 'm_bottom': self.photo_margin_bottom_var,
                        'm_left': self.photo_margin_left_var, 'm_right': self.photo_margin_right_var,
                        'print_scale': self.photo_print_scale_var, 'desc_height': self.photo_desc_height_var,
                        'photo_align': self.photo_align_var, 'auto_fit_page': self.photo_auto_fit_page_var, 'fit_width': self.photo_fit_width_var,
                        'auto_rotate': self.photo_auto_rotate_var, 'width_pct': self.photo_width_pct_var,
                        'height_pct': self.photo_height_pct_var,
                        'pixel_adj': self.photo_width_pixel_adj_var, 'shift_x': self.photo_shift_x_var,
                        'shift_y': self.photo_shift_y_var, 'dpi': self.photo_dpi_var
                    }
                    if 'PHOTO_LOG_SETTINGS' in saved_data:
                        plist = saved_data['PHOTO_LOG_SETTINGS']
                        for pk, pvar in photo_vars.items():
                            if pk in plist:
                                if isinstance(pvar, tk.BooleanVar): pvar.set(bool(plist[pk]))
                                else: pvar.set(str(plist[pk]))
                        # Log success for key fine-tune variables
                        self.log(f"[Restore] 사진대장 설정 복구: 비율({self.photo_width_pct_var.get()}%), 추가({self.photo_width_pixel_adj_var.get()}px), 이동({self.photo_shift_x_var.get()}px)")
                        if 'selected_files' in plist:
                            self.photo_selected_files = plist['selected_files']
                            if hasattr(self, 'photo_listbox'):
                                self.photo_listbox.delete(0, tk.END)
                                for f_path in self.photo_selected_files:
                                    self.photo_listbox.insert(tk.END, f_path)
                        if isinstance(plist.get('item_metadata'), dict):
                            self.photo_item_metadata = plist['item_metadata']

                    # [NEW] Restore custom column layouts if they exist
                    if 'column_keys' in saved_data and isinstance(saved_data['column_keys'], list):
                        self.column_keys = list(saved_data['column_keys'])
                        if "Report No" not in self.column_keys:
                            try: self.column_keys.insert(self.column_keys.index("Date") + 1, "Report No")
                            except: self.column_keys.insert(3, "Report No")
                    if 'rt_column_keys' in saved_data and isinstance(saved_data['rt_column_keys'], list):
                        self.rt_column_keys = list(saved_data['rt_column_keys'])
                        if "Report No" not in self.rt_column_keys:
                            try: self.rt_column_keys.insert(self.rt_column_keys.index("Date") + 1, "Report No")
                            except: self.rt_column_keys.insert(3, "Report No")
                        if "Sec" not in self.rt_column_keys:
                            insert_at = self.rt_column_keys.index("Date") + 2 if "Date" in self.rt_column_keys else 4
                            self.rt_column_keys.insert(insert_at, "Sec")
                    if 'kogas_column_keys' in saved_data and isinstance(saved_data['kogas_column_keys'], list):
                        self.kogas_column_keys = list(saved_data['kogas_column_keys'])
                        if "Report No" not in self.kogas_column_keys:
                            try: self.kogas_column_keys.insert(self.kogas_column_keys.index("Date") + 1, "Report No")
                            except: self.kogas_column_keys.insert(3, "Report No")
                    if 'pt_column_keys' in saved_data and isinstance(saved_data['pt_column_keys'], list):
                        self.pt_column_keys = list(saved_data['pt_column_keys'])
                        if "Report No" not in self.pt_column_keys:
                            try: self.pt_column_keys.insert(self.pt_column_keys.index("Date") + 1, "Report No")
                            except: self.pt_column_keys.insert(3, "Report No")
                        # [마이그레이션] 구버전 config에 Thk.가 없으면 Joint 뒤에 삽입
                        if "Thk." not in self.pt_column_keys:
                            try:
                                insert_at = self.pt_column_keys.index("Joint") + 1
                            except ValueError:
                                insert_at = 5
                            self.pt_column_keys.insert(insert_at, "Thk.")
                    
                    if "No" in self.pt_column_keys:
                        self.pt_column_keys.remove("No")
                    self.pt_column_keys.insert(1, "No")

                    if "NPS" in self.pt_column_keys:
                        self.pt_column_keys.remove("NPS")
                    try:
                        self.pt_column_keys.insert(self.pt_column_keys.index("Joint") + 1, "NPS")
                    except ValueError:
                        self.pt_column_keys.insert(6, "NPS")
                    if 'mt_column_keys' in saved_data:
                        self.mt_column_keys = list(saved_data['mt_column_keys'])
                        if "Report No" not in self.mt_column_keys:
                            try: self.mt_column_keys.insert(self.mt_column_keys.index("Date") + 1, "Report No")
                            except: self.mt_column_keys.insert(3, "Report No")
                    if 'paut_column_keys' in saved_data and isinstance(saved_data['paut_column_keys'], list):
                        loaded_paut_keys = list(saved_data['paut_column_keys'])
                        # [FIX] Ensure new columns are added even if loading from old config
                        default_paut_keys = ["selected", "No", "Date", "Report No", "Line No.", "Joint No.", "Size", "Th'k(mm)", "Acc", "Rej", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "Type of Flaw", "a/l", "a/t", "Welder", "Tested Length", "Evaluation", "Remarks"]
                        for dk in default_paut_keys:
                            if dk not in loaded_paut_keys:
                                insert_idx = default_paut_keys.index(dk)
                                if insert_idx <= len(loaded_paut_keys):
                                    loaded_paut_keys.insert(insert_idx, dk)
                                else:
                                    loaded_paut_keys.append(dk)
                        self.paut_column_keys = loaded_paut_keys
                    self.config.update(saved_data)
                    
                    # Restore Path StringVars in UI
                    pmi_tp = self.config.get('PMI_TEMPLATE_PATH', "")
                    if pmi_tp and not os.path.exists(pmi_tp):
                        for candidate in ("home/data/PMI-템플릿보고서.xlsx", "home/data/SIT-PMI-001.xlsx", "home/data/SIT-PMI-K3 롯데 템플릿보고서.xlsm"):
                            full_c = os.path.abspath(candidate)
                            if os.path.exists(full_c):
                                pmi_tp = full_c
                                break
                    self.target_file_path.set(self.config.get('PMI_TARGET_PATH', ""))
                    self.template_file_path.set(pmi_tp)
                    self.rt_target_file_path.set(self.config.get('RT_TARGET_PATH', ""))
                    self.rt_template_file_path.set(self.config.get('RT_TEMPLATE_PATH', ""))
                    self.kogas_target_file_path.set(self.config.get('KOGAS_TARGET_PATH', ""))
                    self.kogas_template_file_path.set(self.config.get('KOGAS_TEMPLATE_PATH', ""))
                    self.pt_target_file_path.set(self.config.get('PT_TARGET_PATH', ""))
                    self.pt_template_file_path.set(self.config.get('PT_TEMPLATE_PATH', ""))
                    self.mt_target_file_path.set(self.config.get('MT_TARGET_PATH', ''))
                    self.mt_template_file_path.set(self.config.get('MT_TEMPLATE_PATH', ''))
                    self.mt_surface_temp.set(str(self.config.get('MT_SURFACE_TEMP', '')))
                    self.paut_target_file_path.set(self.config.get('PAUT_TARGET_PATH', ""))
                    self.paut_template_file_path.set(self.config.get('PAUT_TEMPLATE_PATH', ""))
                self.log("[SUCCESS] 사용자 저장 설정을 모두 불러왔습니다.")
                
                # Restore Window Geometry/State
                if 'WINDOW_GEOMETRY' in saved_data and not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                    try: self.root.geometry(saved_data['WINDOW_GEOMETRY'])
                    except: pass
                if 'WINDOW_STATE' in saved_data and not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                    try: self.root.state(saved_data['WINDOW_STATE'])
                    except: pass
                if 'ACTIVE_TAB' in saved_data:
                    self._initial_active_tab = saved_data['ACTIVE_TAB']
            except Exception as e:
                print(f"WARNING: 사용자 설정 불러오기 실패: {e}")

        # PAUT 마이그레이션 (UT_* -> PAUT_*)
        for k in list(self.config.keys()):
            if k.startswith("UT_"):
                new_k = k.replace("UT_", "PAUT_")
                if new_k not in self.config:
                    self.config[new_k] = self.config[k]
        
        # [NEW] Load Photo Log specific config
        self.load_photo_log_config()

    def capture_ui_state(self):
        """UI 요소들의 현재 상태(컬럼 너비, 분할선 위치 등)를 config에 반영"""
        try:
            d = self.__dict__
            trees = [
                ('PMI', d.get('preview_tree')), 
                ('RT', d.get('rt_preview_tree')), 
                ('PT', d.get('pt_preview_tree')), 
                ('MT', d.get('mt_preview_tree')),
                ('PAUT', d.get('paut_preview_tree')),
                ('KOGAS', d.get('kogas_preview_tree'))
            ]
            for mode, tree in trees:
                if tree:
                    try:
                        if tree.winfo_ismapped() and tree.winfo_width() > 100:
                            self.config[f"{mode}_COL_WIDTHS"] = {col: tree.column(col, "width") for col in tree["columns"]}
                    except: pass

            # 2. 패널 분할선(Sash) 비율 및 사이드바 접기 상태 캡처 (화면에 표시 중이고 초기 비율 적용이 끝난 패널만 갱신)
            paned_windows = [
                ('PMI', d.get('pmi_paned')),
                ('RT', d.get('rt_paned')),
                ('PT', d.get('pt_paned')),
                ('MT', d.get('mt_paned')),
                ('PAUT', d.get('paut_paned')),
                ('PHOTO', d.get('photo_paned')),
            ]
            hidden_map = self.config.setdefault('SIDEBAR_HIDDEN', {})
            if not isinstance(hidden_map, dict):
                hidden_map = {}
                self.config['SIDEBAR_HIDDEN'] = hidden_map
            for mode, pw in paned_windows:
                if pw and pw.winfo_exists():
                    try:
                        panes = pw.panes()
                        if not panes:
                            continue
                        is_hidden = bool(pw.tk.getboolean(pw.panecget(panes[0], 'hide')))
                        hidden_map[mode] = is_hidden
                        if is_hidden or not getattr(pw, '_sash_initialized', False) or not pw.winfo_ismapped():
                            continue
                        total_w = pw.winfo_width()
                        if total_w > 100:
                            ratio = max(0.08, min(0.92, pw.sash_coord(0)[0] / total_w))
                            self.config[f'{mode}_SASH_RATIO'] = ratio
                            setattr(self, f'{mode.lower()}_pane_ratio', ratio)
                    except Exception:
                        pass

            # 3. 윈도우 및 활성 탭 상태 캡처
            try:
                if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                    self.config['WINDOW_GEOMETRY'] = self.root.geometry()
                if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                    self.config['WINDOW_STATE'] = self.root.state()
                if 'mode_notebook' in d:
                    self.config['ACTIVE_TAB'] = self.mode_notebook.index("current")
                
                # 4. 파일 경로 관성 저장
                self.config['PMI_TARGET_PATH'] = self.target_file_path.get()
                self.config['PMI_TEMPLATE_PATH'] = self.template_file_path.get()
                self.config['RT_TARGET_PATH'] = self.rt_target_file_path.get()
                self.config['RT_TEMPLATE_PATH'] = self.rt_template_file_path.get()
                self.config['PT_TARGET_PATH'] = self.pt_target_file_path.get()
                self.config['PT_TEMPLATE_PATH'] = self.pt_template_file_path.get()
                self.config['MT_TARGET_PATH'] = self.mt_target_file_path.get()
                self.config['MT_TEMPLATE_PATH'] = self.mt_template_file_path.get()
                self.config['PAUT_TARGET_PATH'] = self.paut_target_file_path.get()
                self.config['PAUT_TEMPLATE_PATH'] = self.paut_template_file_path.get()
                self.config['KOGAS_TARGET_PATH'] = self.kogas_target_file_path.get()
                self.config['KOGAS_TEMPLATE_PATH'] = self.kogas_template_file_path.get()
                if 'pmi_layout_mode' in d:
                    self.config['PMI_LAYOUT_MODE'] = self.pmi_layout_mode.get()
                
                if 'paut_equip_vars' in d:
                    for k, var in self.paut_equip_vars.items():
                        self.config[f'PAUT_EQUIP_{k.upper()}'] = var.get()
                
                if 'paut_photo_pages_var' in d:
                    self.config['PAUT_PHOTO_PAGES'] = str(self.paut_photo_pages_var.get())
                # 5. Photo Log Settings Capture
                prev_photo_cfg = self.config.get('PHOTO_LOG_SETTINGS', {}) if isinstance(self.config.get('PHOTO_LOG_SETTINGS'), dict) else {}
                self.config['PHOTO_LOG_SETTINGS'] = {
                    'orderer': self.photo_orderer.get(), 'inspect_date': self.photo_inspect_date.get(),
                    'inspect_type': self.photo_inspect_type.get(), 'report_title': self.photo_report_title.get(),
                    'report_no': self.photo_report_no.get(),
                    'start_page': d['photo_start_page'].get() if 'photo_start_page' in d else prev_photo_cfg.get('start_page', str(self.config.get('PHOTO_START_PAGE', ''))),
                    'total_pages': d['photo_total_pages'].get() if 'photo_total_pages' in d else prev_photo_cfg.get('total_pages', str(self.config.get('PHOTO_TOTAL_PAGES', ''))),
                    'cols_per_row': self.photo_cols_per_row.get(), 'rows_per_page': self.photo_rows_per_page.get(), 'keep_aspect': self.photo_keep_aspect.get(),
                    'output_name': self.photo_output_name.get(), 'logo_path': self.photo_logo_path.get(),
                    'logo_width': self.photo_logo_width_var.get(), 'logo_height': self.photo_logo_height_var.get(),
                    'logo_x': self.photo_logo_x_var.get(), 'logo_y': self.photo_logo_y_var.get(),
                    'cell_width': self.photo_cell_width_var.get(), 'cell_height': self.photo_cell_height_var.get(),
                    'm_top': self.photo_margin_top_var.get(), 'm_bottom': self.photo_margin_bottom_var.get(),
                    'm_left': self.photo_margin_left_var.get(), 'm_right': self.photo_margin_right_var.get(),
                    'print_scale': self.photo_print_scale_var.get(), 'desc_height': self.photo_desc_height_var.get(),
                    'photo_align': self.photo_align_var.get(), 'auto_fit_page': self.photo_auto_fit_page_var.get(), 'fit_width': self.photo_fit_width_var.get(),
                    'auto_rotate': self.photo_auto_rotate_var.get(), 'width_pct': self.photo_width_pct_var.get(),
                    'height_pct': self.photo_height_pct_var.get(),
                    'pixel_adj': self.photo_width_pixel_adj_var.get(), 'shift_x': self.photo_shift_x_var.get(),
                    'shift_y': self.photo_shift_y_var.get(), 'dpi': self.photo_dpi_var.get(),
                    'selected_files': self.photo_selected_files,
                    'item_metadata': self.photo_item_metadata,
                    'last_save_dir': d.get('last_photo_save_dir', "")
                }
                
                # [NEW] Save custom column list configurations
                self.config['column_keys'] = list(self.column_keys)
                self.config['rt_column_keys'] = list(self.rt_column_keys)
                self.config['kogas_column_keys'] = list(self.kogas_column_keys)
                self.config['pt_column_keys'] = list(self.pt_column_keys)
                self.config['mt_column_keys'] = list(self.mt_column_keys)
                self.config['paut_column_keys'] = list(self.paut_column_keys)
            except Exception as e:
                self.log(f"[ERROR] UI 상태 캡처 실패: {e}")
        except Exception as e:
            self.log(f"[ERROR] 전체 상태 캡처 실패: {e}")

    def save_photo_log_config(self):
        """Save Photo Log specific layout settings to a separate JSON."""
        self.capture_ui_state()
        plist = self.config.get('PHOTO_LOG_SETTINGS', {})
        try:
            os.makedirs(CONFIG_DIR, exist_ok=True)
            with open(PHOTO_LOG_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(plist, f, indent=4, ensure_ascii=False)
            self.log("💾 사진대장 전용 설정이 저장되었습니다.")
            messagebox.showinfo("저장 완료", "사진대장 레이아웃 설정이 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    def load_photo_log_config(self):
        """Load Photo Log specific layout settings."""
        config_path = _existing_report_settings(
            PHOTO_LOG_CONFIG_FILE, LEGACY_PHOTO_LOG_CONFIG_FILE
        )
        if not config_path:
            return
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                plist = json.load(f)
            
            # Update UI vars (using same logic as in load_settings)
            photo_vars = {
                'orderer': self.photo_orderer, 'inspect_date': self.photo_inspect_date,
                'inspect_type': self.photo_inspect_type, 'report_title': self.photo_report_title,
                'report_no': self.photo_report_no,
                'start_page': self.__dict__.get('photo_start_page'),
                'total_pages': self.__dict__.get('photo_total_pages'),
                'cols_per_row': self.photo_cols_per_row, 'rows_per_page': self.photo_rows_per_page, 'keep_aspect': self.photo_keep_aspect,
                'output_name': self.photo_output_name, 'logo_path': self.photo_logo_path,
                'logo_width': self.photo_logo_width_var, 'logo_height': self.photo_logo_height_var,
                'logo_x': self.photo_logo_x_var, 'logo_y': self.photo_logo_y_var,
                'cell_width': self.photo_cell_width_var, 'cell_height': self.photo_cell_height_var,
                'm_top': self.photo_margin_top_var, 'm_bottom': self.photo_margin_bottom_var,
                'm_left': self.photo_margin_left_var, 'm_right': self.photo_margin_right_var,
                'print_scale': self.photo_print_scale_var, 'desc_height': self.photo_desc_height_var,
                'photo_align': self.photo_align_var, 'auto_fit_page': self.photo_auto_fit_page_var, 'fit_width': self.photo_fit_width_var,
                'auto_rotate': self.photo_auto_rotate_var, 'width_pct': self.photo_width_pct_var,
                'height_pct': self.photo_height_pct_var,
                'pixel_adj': self.photo_width_pixel_adj_var, 'shift_x': self.photo_shift_x_var,
                'shift_y': self.photo_shift_y_var, 'dpi': self.photo_dpi_var
            }
            for pk, pvar in photo_vars.items():
                if pk in plist and pvar is not None:
                    if isinstance(pvar, tk.BooleanVar): pvar.set(bool(plist[pk]))
                    else: pvar.set(str(plist[pk]))
            
            if 'selected_files' in plist:
                self.photo_selected_files = plist['selected_files']
                if 'photo_listbox' in self.__dict__:
                    self.photo_listbox.delete(0, tk.END)
                    for f_path in self.photo_selected_files:
                        self.photo_listbox.insert(tk.END, f_path)
            if isinstance(plist.get('item_metadata'), dict):
                self.photo_item_metadata = plist['item_metadata']
            
            self.log("📂 사진대장 전용 설정을 불러왔습니다.")
        except Exception as e:
            print(f"Error loading photolog config: {e}")

    def load_template_specific_config(self, template_path, mode="RT", force=False):
        """템플릿 파일과 연동된 JSON 설정 파일을 로드하여 UI에 반영"""
        if not template_path or not os.path.exists(template_path): return
        
        # [NEW] Template-name based heuristics (e.g. K1/K3/롯데/플랜트 배관용)
        if mode == "PMI":
            if any(k in os.path.basename(template_path).upper() for k in ["LOTTE", "PROVIDENCE", "롯데", "모재", "3단"]):
                if hasattr(self, 'pmi_layout_mode'):
                    self.pmi_layout_mode.set("PLANT_TRIPLE")
            else:
                if hasattr(self, 'pmi_layout_mode'):
                    self.pmi_layout_mode.set("STANDARD")

        config_path = os.path.splitext(template_path)[0] + ".json"
        try:
            if not os.path.exists(config_path) or not os.access(config_path, os.R_OK):
                return
            mtime = os.path.getmtime(config_path)
            cache = getattr(self, '_loaded_template_configs', None)
            if cache is None:
                cache = {}
                self._loaded_template_configs = cache
            cache_key = (mode, os.path.abspath(config_path))
            if not force and cache.get(cache_key) == mtime:
                return
            with open(config_path, 'r', encoding='utf-8') as f:
                t_config = json.load(f)
            cache[cache_key] = mtime
            
            self.log(f"📂 [Auto-Load] '{os.path.basename(template_path)}' 전용 설정을 불러왔습니다.")
            
            # Update self.config and UI variables
            for key, value in t_config.items():
                self.config[key] = value
                if key in self.setting_vars:
                    self.setting_vars[key].set(str(value))
            
            # Mode-specific UI update (Force refresh logic if needed)
            if mode in ["RT", "KOGAS"]:
                self._update_rt_preview_columns(mode)
            elif mode == "PT":
                try: self._update_pt_preview_columns()
                except: pass
            elif mode == "PAUT":
                try: self._update_paut_target_fsh()
                except: pass
            elif mode == "PMI":
                try:
                    if "PMI_LAYOUT_MODE" in t_config and hasattr(self, 'pmi_layout_mode'):
                        self.pmi_layout_mode.set(t_config["PMI_LAYOUT_MODE"])
                    elif any(k in os.path.basename(template_path).upper() for k in ["LOTTE", "PROVIDENCE", "롯데", "모재", "3단"]):
                        if hasattr(self, 'pmi_layout_mode'):
                            self.pmi_layout_mode.set("PLANT_TRIPLE")
                    else:
                        if hasattr(self, 'pmi_layout_mode'):
                            self.pmi_layout_mode.set("STANDARD")
                    self._update_pmi_filter_ui()
                except: pass

        except Exception as e:
            self.log(f"[ERROR] 템플릿 전용 설정 로드 실패: {e}")

    def save_template_specific_config(self, mode="RT"):
        """현재 설정을 선택된 템플릿 전용 JSON 파일로 저장"""
        self.save_settings()
        if mode == "RT": template_path = self.rt_template_file_path.get()
        elif mode == "KOGAS": template_path = self.kogas_template_file_path.get()
        elif mode == "PT": template_path = self.pt_template_file_path.get()
        elif mode == "MT": template_path = self.mt_template_file_path.get()
        elif mode == "PAUT": template_path = self.paut_template_file_path.get()
        else: template_path = self.template_file_path.get()
        
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("경고", "먼저 템플릿 파일을 선택해주세요.")
            return
            
        config_path = os.path.splitext(template_path)[0] + ".json"
        self.capture_ui_state() # Ensure current UI is in self.config
        
        # Determine relevant keys to save for this template
        # 1. Logos & Print settings (Template-dependent)
        keys_to_save = []
        for ctx in ["COVER", "DATA"]:
            # Correct Prefix mapping: SITCO_RT_COVER, etc.
            # Actually, mode is like "RT", ctx is "COVER" -> prefix is "RT_COVER"
            # In UI: self._create_setting_grid(t_cover, "RT_COVER")
            # Inside grid: keys are SITCO_RT_COVER_X etc.
            for prefix in ["SITCO", "SEOUL", "FOOTER", "FOOTER_PT"]:
                for suffix in ["_PATH", "_ANCHOR", "_W", "_H", "_X", "_Y"]:
                    keys_to_save.append(f"{prefix}_{mode}_{ctx}{suffix}")
            for suffix in ["_TOP", "_BOTTOM", "_LEFT", "_RIGHT"]:
                keys_to_save.append(f"MARGIN_{mode}_{ctx}{suffix}")
            keys_to_save.append(f"PRINT_SCALE_{mode}_{ctx}")
            keys_to_save.append(f"PRINT_AREA_{mode}_{ctx}")
            
        # 2. Mode-specific boundaries and column mappings
        if mode == "RT":
            keys_to_save += ["RT_START_ROW", "RT_END_ROW"]
            keys_to_save += ["RT_FILM_BRAND", "RT_FILM_TYPE", "RT_FILM_SIZE",
                             "RT_ISOTOPE", "RT_ACTIVITY", "RT_ACTIVITY_UNIT"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("RT_COL_") or k.startswith("RT_NAME_")]
        elif mode == "KOGAS":
            keys_to_save += ["KOGAS_START_ROW", "KOGAS_DATA_END_ROW", "RT_KOGAS_D_START_COL"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("KOGAS_")]
        elif mode == "PT":
            keys_to_save += ["PT_START_ROW", "PT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PT_COL_") or k.startswith("PT_NAME_")]
        elif mode == "MT":
            keys_to_save += [k for k in self.config if k.startswith("MT_")]
        elif mode == "PAUT":
            keys_to_save += ["PAUT_START_ROW", "PAUT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PAUT_COL_") or k.startswith("PAUT_NAME_")]
        elif mode == "PMI":
            keys_to_save += ["PMI_START_ROW", "PMI_DATA_END_ROW", "PMI_PRINT_END_ROW", "PMI_LAYOUT_MODE"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PMI_COL_") or k.startswith("PMI_NAME_")]

        # Build template-specific dict
        t_data = {k: self.config[k] for k in keys_to_save if k in self.config}
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(t_data, f, indent=4, ensure_ascii=False)
            self.log(f"💾 [SUCCESS] '{os.path.basename(template_path)}' 전용 설정이 저장되었습니다.")
            messagebox.showinfo("저장 완료", f"템플릿 전용 설정이 저장되었습니다.\n{os.path.basename(config_path)}")
        except Exception as e:
            messagebox.showerror("저장 실패", f"설정 저장 중 오류 발생: {e}")


    def save_template_specific_config_silent(self, mode="RT"):
        """현재 설정을 선택된 템플릿 전용 JSON 파일로 조용히(팝업 없이) 저장"""
        if mode == "RT": template_path = self.rt_template_file_path.get()
        elif mode == "KOGAS": template_path = self.kogas_template_file_path.get()
        elif mode == "PT": template_path = self.pt_template_file_path.get()
        elif mode == "MT": template_path = self.mt_template_file_path.get()
        elif mode == "PAUT": template_path = self.paut_template_file_path.get()
        else: template_path = self.template_file_path.get()
        
        if not template_path or not os.path.exists(template_path):
            return
            
        config_path = os.path.splitext(template_path)[0] + ".json"
        
        # Determine relevant keys to save for this template
        keys_to_save = []
        for ctx in ["COVER", "DATA"]:
            for prefix in ["SITCO", "SEOUL", "FOOTER", "FOOTER_PT"]:
                for suffix in ["_PATH", "_ANCHOR", "_W", "_H", "_X", "_Y"]:
                    keys_to_save.append(f"{prefix}_{mode}_{ctx}{suffix}")
            for suffix in ["_TOP", "_BOTTOM", "_LEFT", "_RIGHT"]:
                keys_to_save.append(f"MARGIN_{mode}_{ctx}{suffix}")
            keys_to_save.append(f"PRINT_SCALE_{mode}_{ctx}")
            keys_to_save.append(f"PRINT_AREA_{mode}_{ctx}")
            
        # 2. Mode-specific boundaries and column mappings
        if mode == "RT":
            keys_to_save += ["RT_START_ROW", "RT_END_ROW"]
            keys_to_save += ["RT_FILM_BRAND", "RT_FILM_TYPE", "RT_FILM_SIZE",
                             "RT_ISOTOPE", "RT_ACTIVITY", "RT_ACTIVITY_UNIT"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("RT_COL_") or k.startswith("RT_NAME_")]
        elif mode == "KOGAS":
            keys_to_save += ["KOGAS_START_ROW", "KOGAS_DATA_END_ROW", "RT_KOGAS_D_START_COL"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("KOGAS_")]
        elif mode == "PT":
            keys_to_save += ["PT_START_ROW", "PT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PT_COL_") or k.startswith("PT_NAME_")]
        elif mode == "MT":
            keys_to_save += [k for k in self.config if k.startswith("MT_")]
        elif mode == "PAUT":
            keys_to_save += ["PAUT_START_ROW", "PAUT_END_ROW"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PAUT_COL_") or k.startswith("PAUT_NAME_")]
        elif mode == "PMI":
            keys_to_save += ["PMI_START_ROW", "PMI_DATA_END_ROW", "PMI_PRINT_END_ROW", "PMI_LAYOUT_MODE"]
            keys_to_save += [k for k in self.config.keys() if k.startswith("PMI_COL_") or k.startswith("PMI_NAME_")]

        # Build template-specific dict
        t_data = {k: self.config[k] for k in keys_to_save if k in self.config}
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(t_data, f, indent=4, ensure_ascii=False)
            cache = getattr(self, '_loaded_template_configs', None)
            if isinstance(cache, dict):
                cache[(mode, os.path.abspath(config_path))] = os.path.getmtime(config_path)
        except Exception as e:
            print(f"[ERROR] 템플릿 전용 설정 자동 저장 실패: {e}")


    def save_settings(self, silent=True):
        """현재 설정을 파일(JSON)에 저장"""
        if not getattr(self, '_is_ready', True):
            return True, "초기화 중 저장 생략"
        self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        self.capture_ui_state()
        try:
            if hasattr(self, 'setting_vars'):
                for key, var in self.setting_vars.items():
                    val = var.get()
                    try:
                        # [FIX] Handle various key types correctly for persistence
                        if "AREA" in key or "_PATH" in key or "_NAME" in key or ("_COL_" in key and "WIDTH" not in key):
                            self.config[key] = str(val).strip()
                        elif any(key.endswith(x) for x in ['_X', '_Y', '_W', '_H']) or any(x in key for x in ['MARGIN', 'HEIGHT', 'WIDTH']):
                            try: self.config[key] = float(val) if str(val).strip() else 0.0
                            except: self.config[key] = 0.0
                        elif 'SCALE' in key or key.endswith(('_ROW', '_COL', '_IDX')):
                            # Column indexes and row numbers should be handled as strings if they are like 'A', 
                            # but for safety in logic we often use col_to_num. 
                            # Here we store them as strings to preserve 'A', 'B' etc.
                            self.config[key] = str(val).strip()
                        else:
                            self.config[key] = str(val)
                    except Exception as e:
                        self.config[key] = str(val)

            # [NEW] 필터 설정 저장 (PMI)
            if hasattr(self, 'element_filters'):
                filter_data = []
                for f_item in self.element_filters:
                    filter_data.append({
                        'key': f_item['key'].get(),
                        'min': f_item['min'].get(),
                        'max': f_item['max'].get()
                    })
                self.config["PMI_FILTERS"] = filter_data

            # [NEW] 탭별 전용 행 설정 강제 저장
            self.config["PMI_START_ROW"] = self.config.get("START_ROW", 19)
            self.config["PMI_DATA_END_ROW"] = self.config.get("DATA_END_ROW", 47)
            self.config["PMI_PRINT_END_ROW"] = self.config.get("PRINT_END_ROW", 47)

            # [NEW] Window State Capture
            if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'geometry'):
                self.config['WINDOW_GEOMETRY'] = self.root.geometry()
            if not getattr(self, 'is_embedded', False) and hasattr(self.root, 'state'):
                self.config['WINDOW_STATE'] = self.root.state()
            try: self.config['ACTIVE_TAB'] = self.mode_notebook.index("current")
            except: pass
            # [FIX] Ensure the config directory exists (Prevents Errno 2)
            os.makedirs(CONFIG_DIR, exist_ok=True)

            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, default=str)

            # [NEW] 모든 모드의 템플릿 전용 설정 파일도 함께 동기화(자동 저장)하여 값 덮어쓰기 방지
            for m in ["PMI", "RT", "KOGAS", "PT", "MT", "PAUT"]:
                self.save_template_specific_config_silent(m)

            if not silent:
                self.log("[SUCCESS] 모든 설정이 파일에 안전하게 저장되었습니다.")
            return True, "성공"
            return True, "성공"
        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            self.log(f"[ERROR] 설정 저장 실패 상세:\n{err_msg}")
            return False, str(e)

    def manual_save_settings(self):
        """사용자가 수동으로 버튼을 눌러 저장할 때 알림 표시"""
        success, msg = self.save_settings(silent=False)
        if success:
            messagebox.showinfo("성공", "현재 설정이 안전하게 저장되었습니다.")
        else:
            messagebox.showerror("오류", f"설정 저장 중 문제가 발생했습니다.\n\n원인: {msg}")


    def load_daily_usage_from_central(self):
        """중앙지사 앱의 일보 데이터(daily_usage_df)로부터 데이터를 추출하여 현재 보고서에 채움"""
        import pandas as pd
        
        target_date = None
        if hasattr(self, 'central_app') and self.central_app:
            if hasattr(self.central_app, 'selected_date') and self.central_app.selected_date:
                target_date = str(self.central_app.selected_date).strip()
            elif hasattr(self.central_app, 'date_entry'):
                try:
                    target_date = str(self.central_app.date_entry.get()).strip()
                except Exception:
                    pass

        if not target_date:
            target_date = datetime.datetime.now().strftime("%Y-%m-%d")

        input_date = simpledialog.askstring(
            "중앙지사 일보 데이터 불러오기",
            "불러올 작업 일자를 입력하세요 (YYYY-MM-DD):",
            initialvalue=target_date,
            parent=self.root
        )
        if not input_date:
            return
        target_date = input_date.strip()

        df = None
        if hasattr(self, 'central_app') and self.central_app and hasattr(self.central_app, 'daily_usage_df'):
            df = self.central_app.daily_usage_df
        else:
            # Fallback: Material_Inventory.xlsx 직접 로드
            try:
                from ..sync_google_sheet import get_sync_inventory_path
                inv_path = get_sync_inventory_path()
                if os.path.exists(inv_path):
                    df = pd.read_excel(inv_path, sheet_name="DailyUsage")
            except Exception as e:
                messagebox.showerror("오류", f"재고 데이터 파일을 읽을 수 없습니다: {e}", parent=self.root)
                return

        if df is None or df.empty:
            messagebox.showwarning("알림", "불러올 일보 데이터가 없습니다.", parent=self.root)
            return

        date_col = 'Date' if 'Date' in df.columns else df.columns[0]
        # 날짜 필터링 (다양한 날짜 포맷 허용)
        df_filtered = df[df[date_col].astype(str).str.contains(target_date, na=False)]
        if df_filtered.empty:
            # YYYYMMDD 형태도 확인
            compact_date = target_date.replace('-', '').replace('.', '')
            df_filtered = df[df[date_col].astype(str).str.contains(compact_date, na=False)]

        if df_filtered.empty:
            messagebox.showinfo("알림", f"{target_date} 일자의 일보 데이터가 없습니다.", parent=self.root)
            return

        # 기본 정보 설정
        first_row = df_filtered.iloc[0]
        self.gapji_exam_date.set(target_date)
        cust = str(first_row.get('업체명') or first_row.get('Site') or '한국지역난방공사 중앙지사')
        if cust and cust != 'nan':
            self.gapji_customer.set(cust)
        proj = str(first_row.get('구분') or first_row.get('Site') or '열수송관 개체공사')
        if proj and proj != 'nan':
            self.gapji_project.set(proj)
        item = str(first_row.get('검사품명') or 'Pipe')
        if item and item != 'nan':
            self.gapji_item.set(item)

        # 현재 활성화된 모드에 데이터 추가
        current_mode = getattr(self, 'current_mode', 'PMI')
        count = 0

        for idx, row in df_filtered.iterrows():
            m_method = str(row.get('검사방법') or '').strip().upper()
            welder = str(row.get('검사자') or row.get('User') or '').strip()
            size = str(row.get('관경(Inch)') or '').strip()
            joints = row.get('조인트수') or 1
            try:
                num_joints = int(joints)
            except:
                num_joints = 1

            for j_idx in range(num_joints):
                joint_name = f"W-{j_idx+1:02d}" if num_joints > 1 else "W-01"
                count += 1
                item_dict = {
                    "selected": True,
                    "No": str(count),
                    "Date": target_date,
                    "Report No": str(row.get('성적서번호') or ''),
                    "Dwg": str(row.get('구분') or '-'),
                    "Joint": joint_name,
                    "Loc": str(row.get('Site') or '중앙지사'),
                    "Size": size,
                    "T": "8.0",
                    "Mat": str(row.get('Material') or 'SPW'),
                    "Result": "ACC",
                    "Welder": welder,
                    "Remarks": str(row.get('Remark') or '')
                }

                if current_mode == 'RT':
                    self.rt_extracted_data.append(item_dict)
                elif current_mode == 'MT':
                    self.mt_extracted_data.append(item_dict)
                elif current_mode == 'PT':
                    self.pt_extracted_data.append(item_dict)
                elif current_mode == 'PAUT':
                    item_dict.update({
                        "Line No.": str(row.get('구분') or '-'),
                        "Joint No.": joint_name,
                        "Th'k(mm)": "8.0",
                        "Acc": "O",
                        "Rej": "-",
                        "Evaluation": "ACC"
                    })
                    self.paut_extracted_data.append(item_dict)
                else:
                    self.extracted_data.append(item_dict)

        # 뷰 새로고침
        if hasattr(self, '_update_preview_table'):
            try:
                self._update_preview_table(current_mode)
            except Exception:
                pass
        
        self.log(f"[INFO] {target_date} 일보 데이터 {count}건 불러오기 완료 ({current_mode} 모드)")
        messagebox.showinfo("완료", f"{target_date} 일자의 일보 데이터 {count}건을 성공적으로 불러왔습니다!\n현재 모드: {current_mode}", parent=self.root)

    def on_closing(self):
        """애플리케이션 종료 시 최종 상태 저장"""
        try:
            self.capture_ui_state()
            self.save_settings()
        except: pass
        if not getattr(self, "is_embedded", False):
            try:
                self.root.destroy()
            except: pass

    def _apply_sash_ratio(self, mode):
        """Helper to apply saved sash ratio with a slight delay to ensure UI mapping."""
        pw = getattr(self, f"{mode.lower()}_paned", None)
        default_ratio = 0.45 if mode == "PHOTO" else 0.5
        ratio = self.config.get(f"{mode}_SASH_RATIO", default_ratio)
        try:
            ratio = max(0.08, min(0.92, float(ratio)))
        except (TypeError, ValueError):
            ratio = default_ratio
        setattr(self, f"{mode.lower()}_pane_ratio", ratio)
        if not pw:
            return

        def apply(attempts=0):
            if not pw.winfo_exists():
                return
            try:
                panes = pw.panes()
                if panes and pw.tk.getboolean(pw.panecget(panes[0], 'hide')):
                    pw._sash_initialized = True
                    return
                total_w = pw.winfo_width()
                if pw.winfo_ismapped() and total_w > 100:
                    pw.sash_place(0, int(total_w * getattr(self, f"{mode.lower()}_pane_ratio", ratio)), 0)
                    pw._sash_initialized = True
                    return
            except Exception:
                return
            if attempts < 20:
                try:
                    if self.root.winfo_exists():
                        self.root.after(100, lambda: apply(attempts + 1))
                except tk.TclError:
                    pass

        try:
            if self.root.winfo_exists():
                self.root.after(100, apply)
        except tk.TclError:
            pass

    def _infer_paned_mode(self, paned):
        for mode in ("PMI", "RT", "PT", "MT", "PAUT", "PHOTO"):
            if getattr(self, f"{mode.lower()}_paned", None) is paned:
                return mode
        return None

    def _bind_report_paned(self, paned, mode, default_ratio=0.5):
        """보고서 서브탭 PanedWindow의 분할선 비율 유지 및 드래그 즉시 저장 바인딩을 설정한다."""
        ratio_attr = f"{mode.lower()}_pane_ratio"
        try:
            saved = float(self.config.get(f"{mode}_SASH_RATIO", default_ratio))
            saved = max(0.08, min(0.92, saved))
        except (TypeError, ValueError):
            saved = default_ratio
        setattr(self, ratio_attr, saved)

        def on_configure(event=None):
            if event is not None and event.widget is not paned:
                return
            if not paned.winfo_exists() or not getattr(paned, '_sash_initialized', False):
                return
            try:
                panes = paned.panes()
                if panes and paned.tk.getboolean(paned.panecget(panes[0], 'hide')):
                    return
                total_w = paned.winfo_width()
                if total_w > 100:
                    cur_ratio = float(getattr(self, ratio_attr, saved))
                    paned.sash_place(0, int(total_w * cur_ratio), 0)
            except Exception:
                pass

        def on_press(_event=None):
            try:
                panes = paned.panes()
                if panes and paned.tk.getboolean(paned.panecget(panes[0], 'hide')):
                    paned._sash_press_x = None
                    return
                paned._sash_press_x = paned.sash_coord(0)[0]
            except Exception:
                paned._sash_press_x = None

        def on_release(_event=None):
            prev_x = getattr(paned, '_sash_press_x', None)
            paned._sash_press_x = None
            if prev_x is None or not getattr(paned, '_sash_initialized', False):
                return

            def commit():
                if not paned.winfo_exists():
                    return
                try:
                    panes = paned.panes()
                    if panes and paned.tk.getboolean(paned.panecget(panes[0], 'hide')):
                        return
                    total_w = paned.winfo_width()
                    cur_x = paned.sash_coord(0)[0]
                    if total_w > 100 and abs(cur_x - prev_x) >= 2:
                        new_ratio = max(0.08, min(0.92, cur_x / total_w))
                        setattr(self, ratio_attr, new_ratio)
                        self.config[f"{mode}_SASH_RATIO"] = new_ratio
                        self.save_settings()
                except Exception:
                    pass

            self.root.after(10, commit)

        paned.bind("<Configure>", on_configure, add="+")
        paned.bind("<ButtonPress-1>", on_press, add="+")
        paned.bind("<ButtonRelease-1>", on_release, add="+")

    def evaluate_paut_flaw(self, t, h, l, depth, flaw_nature, mode="ECA"):
        """
        PAUT Acceptance Criteria Logic
        - Standard (ASME B31.1): Fatal flaws (Crack, LOF, IP) are rejected immediately.
        - ECA (Fracture Mechanics): Only Crack is fatal. LOF/IP are evaluated by size.
        """
        try:
            t_val = float(t)
            h_val = float(h)
            l_val = float(l)
            d_val = float(depth)
        except (ValueError, TypeError):
            return "Error (Invalid Dimension)", "Unknown"

        # 0. Automatic Location Determination
        s_top = d_val
        s_bottom = t_val - (d_val + h_val)
        s = min(s_top, s_bottom)
        s_limit = 0.4 * (h_val / 2)
        
        loc = "Surface" if s <= s_limit else "Subsurface"

        # 1. Immediate Rejection (Crack)
        nature_str = str(flaw_nature).strip().lower()
        if 'crack' in nature_str or '균열' in nature_str:
            return "Reject (Crack)", loc
            
        # [NEW] Mode-based Rejection for LOF (LF) and IP
        if mode == "Standard":
            unacceptable_types = ['lof', 'lack of fusion', 'ip', 'incomplete penetration', 'lf']
            if any(x in nature_str for x in unacceptable_types):
                return f"Reject ({flaw_nature})", loc
        
        if t_val <= 0:
            return "Error (Zero/Negative Th'k)", "-"
            
        if l_val <= 0 and h_val <= 0:
            return "Accept", "-"
            
        if l_val <= 0 or h_val <= 0:
            return "Error (Zero/Negative Value)", loc

        # 1.1 Special Rules for 6mm <= t < 13mm
        if 6 <= t_val < 13:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            if t_val < 10: h_surf_max, h_sub_max = 0.95, 0.96
            elif t_val < 12: h_surf_max, h_sub_max = 1.04, 1.04
            else: h_surf_max, h_sub_max = 1.13, 1.14
                
            limit = h_surf_max if loc == "Surface" else h_sub_max
            if h_val > limit: return f"Reject ({loc} h: {h_val} > {limit}mm)", loc
            return f"Accept (h: {h_val} <= {limit}mm)", loc

        # 1.2 Special Rules for 13mm <= t < 25.4mm
        if 13 <= t_val < 25.4:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            actual_h_t = h_val / t_val
            allowed_h_t = 0.087 if loc == "Surface" else 0.143 
            if actual_h_t > allowed_h_t: return f"Reject ({loc} h/t: {actual_h_t:.3f} > {allowed_h_t:.3f})", loc
            return f"Accept (h/t: {actual_h_t:.3f} <= {allowed_h_t:.3f})", loc
                
        # 2. Aspect Ratio (a/l) logic for t >= 25.4mm
        a_val = h_val if loc == "Surface" else h_val / 2
        aspect_ratio_a_l = a_val / l_val
        
        master_table = [
            (0.00, 0.031, 0.034), (0.05, 0.033, 0.038), (0.10, 0.036, 0.043),
            (0.15, 0.041, 0.054), (0.20, 0.047, 0.066), (0.25, 0.055, 0.078),
            (0.30, 0.064, 0.090), (0.35, 0.074, 0.103), (0.40, 0.083, 0.116),
            (0.45, 0.085, 0.129), (0.50, 0.087, 0.143)
        ]
        
        allowed_a_t = 0
        for ar_limit, surf_a_t, sub_a_t in master_table:
            if aspect_ratio_a_l <= ar_limit:
                allowed_a_t = surf_a_t if loc == 'Surface' else sub_a_t
                break
        if allowed_a_t == 0: allowed_a_t = master_table[-1][1] if loc == 'Surface' else master_table[-1][2]

        actual_a_t = a_val / t_val
        if actual_a_t <= allowed_a_t: return f"Accept (a/t: {actual_a_t:.3f} <= {allowed_a_t:.3f})", loc
        return f"Reject ({loc} a/t: {actual_a_t:.3f} > {allowed_a_t:.3f})", loc

    _LAZY_SUBTAB_ATTRS = {
        'preview_tree': 0, 'pmi_paned': 0, 'pmi_tab_notebook': 0, 'filter_container': 0,
        'rt_preview_tree': 1, 'kogas_preview_tree': 1, 'rt_paned': 1, 'rt_tab_nb': 1, 'rt_preview_nb': 1, 'rt_sub_mode': 1,
        'pt_preview_tree': 2, 'pt_paned': 2, 'pt_tab_notebook': 2,
        'mt_preview_tree': 3, 'mt_paned': 3, 'mt_tab_notebook': 3,
        'paut_preview_tree': 4, 'paut_paned': 4, 'paut_tab_notebook': 4, 'paut_equip_vars': 4, 'paut_photo_pages_var': 4,
        'photo_listbox': 5, 'photo_paned': 5, 'photo_preview_canvas': 5, 'photo_start_page': 5, 'photo_total_pages': 5,
    }

    def __getattr__(self, name):
        tab_idx = self._LAZY_SUBTAB_ATTRS.get(name)
        if tab_idx is not None and '_subtab_builders' in self.__dict__:
            if tab_idx not in self.__dict__.get('_built_subtabs', set()):
                self._ensure_subtab_built(tab_idx)
                if name in self.__dict__:
                    return self.__dict__[name]
        raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")

    def _ensure_subtab_built(self, tab_idx):
        if not hasattr(self, '_subtab_builders') or tab_idx in self._built_subtabs:
            return
        builder = self._subtab_builders.get(tab_idx)
        if not builder:
            return
        self._built_subtabs.add(tab_idx)
        prev_ready = getattr(self, '_is_ready', False)
        self._is_ready = False
        try:
            builder()
        finally:
            self._is_ready = prev_ready

    def _on_main_tab_changed(self, event):
        try:
            # 탭 전환 시 메모리 내 모드별 리포트 정보만 즉시 동기화 (불필요한 8개 JSON 디스크 쓰기 생략)
            self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        except: pass
        try:
            selected_tab = self.mode_notebook.select()
            tab_idx = self.mode_notebook.index(selected_tab)
            self._ensure_subtab_built(tab_idx)
            tab_text = self.mode_notebook.tab(selected_tab, "text")
            if "PMI" in tab_text: self.current_mode = "PMI"
            elif "RT" in tab_text:
                # RT 탭 진입 시 현재 설정된 서브 모드(표준/가스공사) 확인
                self.current_mode = self.rt_sub_mode.get()
            elif "PT" in tab_text: self.current_mode = "PT"
            elif "MT" in tab_text: self.current_mode = "MT"
            elif "PAUT" in tab_text: self.current_mode = "PAUT"
            elif "사진" in tab_text:
                self.current_mode = "PHOTO"
                self._sync_photo_report_info()

            if self.current_mode in self._report_info_modes():
                self._load_report_info_for_mode(self.current_mode)
            
            # [NEW] 모드 변경 시 해당 모드의 설정을 다시 로드
            if self.current_mode in ["RT", "KOGAS"]:
                t_path = self.rt_template_file_path.get() if self.current_mode == "RT" else self.kogas_template_file_path.get()
                self.load_template_specific_config(t_path, self.current_mode)
        except: pass

    @staticmethod
    def _report_info_modes():
        return ("PMI", "RT", "KOGAS", "PT", "MT", "PAUT")

    def _initialize_report_info_by_mode(self):
        """기존 공통정보를 마이그레이션하여 검사방법별 정보로 준비한다."""
        common_values = {
            'project': str(self.config.get('GAPJI_PROJECT', '')).strip(),
            'customer': str(self.config.get('GAPJI_CUSTOMER', '')).strip(),
            'item': str(self.config.get('GAPJI_ITEM', '')).strip(),
            'material': str(self.config.get('GAPJI_MATERIAL', '')).strip(),
            'report_no': str(self.config.get('GAPJI_REPORT_NO', '')).strip(),
            'exam_date': str(self.config.get('GAPJI_EXAM_DATE', '')).strip(),
        }
        self.report_info_by_mode = {}
        for mode in self._report_info_modes():
            self.report_info_by_mode[mode] = {
                key: str(self.config.get(
                    f'{mode}_{key.upper()}', common_value
                )).strip()
                for key, common_value in common_values.items()
            }
        if self.report_info_by_mode['MT']['item'].upper() == 'HOT TEPPING':
            self.report_info_by_mode['MT']['item'] = 'HOT TAPPING'
            self.config['MT_ITEM'] = 'HOT TAPPING'
        mt_item_upper = self.report_info_by_mode['MT']['item'].upper()
        cur_mt_mat = self.report_info_by_mode['MT']['material'].strip()
        if mt_item_upper == 'HOT TAPPING':
            self.report_info_by_mode['MT']['material'] = 'SPPS250'
            self.config['MT_MATERIAL'] = 'SPPS250'
        elif mt_item_upper == 'PIPE' and cur_mt_mat in ('', 'SPPS250'):
            self.report_info_by_mode['MT']['material'] = 'SPPS380'
            self.config['MT_MATERIAL'] = 'SPPS380'
        self._last_auto_material_sig = {}
        self._report_info_loading = False

    def _store_report_info_for_mode(self, mode):
        if (
            mode not in self._report_info_modes()
            or not hasattr(self, 'report_info_by_mode')
            or getattr(self, '_report_info_loading', False)
        ):
            return
        info = {
            'project': self.gapji_project.get().strip(),
            'customer': self.gapji_customer.get().strip(),
            'item': self.gapji_item.get().strip(),
            'material': self.gapji_material.get().strip(),
            'report_no': self.gapji_report_no.get().strip(),
            'exam_date': self.gapji_exam_date.get().strip(),
        }
        self.report_info_by_mode[mode] = info
        for key, value in info.items():
            self.config[f'{mode}_{key.upper()}'] = value
        # 현재 탭의 성적서 생성 함수가 사용하는 공통 키도 함께 맞춘다.
        for key, value in info.items():
            self.config[f'GAPJI_{key.upper()}'] = value

    def _load_report_info_for_mode(self, mode):
        if mode not in self._report_info_modes() or not hasattr(self, 'report_info_by_mode'):
            return
        info = self.report_info_by_mode[mode]
        self._report_info_loading = True
        try:
            self.gapji_project.set(info['project'])
            self.gapji_customer.set(info['customer'])
            self.gapji_item.set(info['item'])
            self.gapji_material.set(info['material'])
            self.gapji_report_no.set(info['report_no'])
            self.gapji_exam_date.set(info['exam_date'])
            for key, value in info.items():
                self.config[f'GAPJI_{key.upper()}'] = value
        finally:
            self._report_info_loading = False

    def _on_report_info_changed(self):
        if getattr(self, '_report_info_loading', False):
            return
        self._store_report_info_for_mode(getattr(self, 'current_mode', ''))
        self._sync_photo_report_info()

    def _mode_item_name(self, mode):
        if getattr(self, 'current_mode', '') == mode and hasattr(self, 'gapji_item'):
            return self.gapji_item.get().strip()
        if hasattr(self, 'report_info_by_mode') and mode in self.report_info_by_mode:
            return str(self.report_info_by_mode[mode].get('item', '')).strip()
        return self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else ''

    def _restore_gapji_input_focus(self, widget):
        """Keep the insertion cursor visible after a cover combobox selection."""
        if widget is None:
            return

        def restore():
            try:
                if not widget.winfo_exists():
                    return
                widget.focus_set()
                widget.icursor(tk.END)
                widget.selection_clear()
            except (tk.TclError, AttributeError):
                pass

        try:
            self.root.after_idle(restore)
        except (tk.TclError, AttributeError):
            pass

    def _on_gapji_material_selected(self, mode, from_focus_out=False, source_widget=None):
        mat = self.gapji_material.get().strip()
        if hasattr(self, 'report_info_by_mode') and mode in self.report_info_by_mode:
            self.report_info_by_mode[mode]['material'] = mat
        self.config['GAPJI_MATERIAL'] = mat
        self.config[f'{mode}_MATERIAL'] = mat
        if mode == 'MT' and getattr(self, 'mt_extracted_data', None):
            from site_apps.central.src.reports.pipe_material_spec import (
                infer_row_pipe_material, is_row_active_for_report,
            )
            item_name = self._mode_item_name('MT')
            changed = False
            for row in self.mt_extracted_data:
                if is_row_active_for_report(row, mode='MT'):
                    target_mat = (
                        infer_row_pipe_material(row, item_name) or mat
                        if mat == 'SPPS380 / SPW400' else mat
                    )
                    if target_mat and row.get('Material') != target_mat:
                        row['Material'] = target_mat
                        changed = True
            if changed and not from_focus_out:
                self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')
        self.save_settings()
        if not from_focus_out:
            self._restore_gapji_input_focus(source_widget)

    def _sync_gapji_material_from_items(self, mode, data_list, force=False):
        """현재 활성 행들의 관경(20A~600A -> SPPS380, 650A~900A -> SPW400)에 맞춰 갑지 재질을 자동 동기화한다."""
        if mode not in ('RT', 'KOGAS', 'PT', 'MT', 'PAUT') or not data_list:
            return
        from site_apps.central.src.reports.pipe_material_spec import (
            KDHC_PIPE_MATERIALS, infer_kdhc_pipe_material,
            infer_row_pipe_material, is_row_active_for_report,
        )
        item_name = self._mode_item_name(mode)
        inferred = infer_kdhc_pipe_material(data_list, mode=mode, item_name=item_name)
        if not inferred:
            return
        if not hasattr(self, '_last_auto_material_sig') or not isinstance(self._last_auto_material_sig, dict):
            self._last_auto_material_sig = {}
        sig = (item_name.upper(), inferred)
        prev_sig = self._last_auto_material_sig.get(mode)
        if not force and prev_sig == sig:
            return
        self._last_auto_material_sig[mode] = sig

        if hasattr(self, 'report_info_by_mode') and mode in self.report_info_by_mode:
            self.report_info_by_mode[mode]['material'] = inferred
        self.config[f'{mode}_MATERIAL'] = inferred

        cur_mode = getattr(self, 'current_mode', '')
        if cur_mode == mode or (not cur_mode and hasattr(self, 'gapji_material')):
            if self.gapji_material.get().strip() != inferred:
                self.gapji_material.set(inferred)
            self.config['GAPJI_MATERIAL'] = inferred

        if mode == 'MT':
            auto_set = set(KDHC_PIPE_MATERIALS) | {''}
            for row in data_list:
                if is_row_active_for_report(row, mode='MT'):
                    cur_row_mat = str(row.get('Material', '') or '').strip()
                    if force or cur_row_mat in auto_set:
                        row['Material'] = infer_row_pipe_material(row, item_name) or inferred

    def _apply_mt_item_selection(self):
        from site_apps.central.src.reports.pipe_material_spec import (
            infer_kdhc_pipe_material, infer_row_pipe_material, is_row_active_for_report,
        )
        item_name = self.gapji_item.get().strip()
        if item_name.upper() == 'HOT TAPPING':
            material = 'SPPS250'
        elif item_name.upper() == 'PIPE':
            material = infer_kdhc_pipe_material(self.mt_extracted_data, mode='MT', item_name='PIPE') or 'SPPS380'
        else:
            return
        if not hasattr(self, '_last_auto_material_sig') or not isinstance(self._last_auto_material_sig, dict):
            self._last_auto_material_sig = {}
        self._last_auto_material_sig['MT'] = (item_name.upper(), material)
        self.gapji_material.set(material)
        for item in self.mt_extracted_data:
            if is_row_active_for_report(item, mode='MT'):
                item['TestItem'] = item_name
                item['Material'] = infer_row_pipe_material(item, item_name) or material
        self.save_settings()
        self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')

    def _sync_photo_report_info(self):
        """사진대장에서 선택한 검사방법의 저장정보를 불러온다."""
        if not hasattr(self, 'photo_orderer'):
            return
        self.photo_orderer.set("한국지역난방공사 중앙지사")
        mode = self.photo_inspect_type.get().strip().upper()
        info = getattr(self, 'report_info_by_mode', {}).get(mode)
        if info:
            report_no = info['report_no'].strip()
            self.photo_report_no.set(f"Photo {report_no}" if report_no else "")
            self.photo_inspect_date.set(info['exam_date'])

    def _sync_photo_output_name(self):
        number = self.photo_report_no.get().strip()
        if not number:
            self.photo_output_name.set('Photo.xlsx')
            return
        if not re.match(r'^Photo(?:\s|$)', number, re.I):
            number = 'Photo ' + number
        number = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '-', number).rstrip(' .')
        self.photo_output_name.set(number[:180].rstrip(' .') + '.xlsx')

    def create_widgets(self):
        style = ttk.Style()
        if not getattr(self, 'is_embedded', False):
            style.theme_use('clam')
            style.configure("TFrame", background="#f9fafb")
            style.configure("TLabel", background="#f9fafb", font=("Malgun Gothic", 10))
            style.configure("TLabelframe", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
            style.configure("TLabelframe.Label", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
            style.configure("Action.TButton", font=("Malgun Gothic", 11, "bold"), padding=10)
            style.map("TEntry", 
                      selectbackground=[('focus', '#3b82f6'), ('!focus', '#3b82f6')],
                      selectforeground=[('focus', 'white'), ('!focus', 'white')])
            style.configure("Treeview", rowheight=28, font=("Malgun Gothic", 10))
            style.map("Treeview", background=[('selected', '#3b82f6')], foreground=[('selected', 'white')])
            style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"))
        
        # [NEW] 탭 스타일 설정
        style.configure("Main.TNotebook.Tab", font=("Malgun Gothic", 11, "bold"), padding=[10, 3])

        # --- Main Scrollable Container ---
        self.canvas = tk.Canvas(self.root, background="#f9fafb", highlightthickness=0, yscrollincrement=40)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, background="#f9fafb")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self._safe_update_scrollregion()
        )

        self.canvas_frame_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        def _on_canvas_configure(event):
            # [ROOT CAUSE FIX] Always force width, handle None bbox gracefully
            self.canvas.itemconfigure(self.canvas_frame_window, width=event.width)
            self._safe_update_scrollregion()

        self.canvas.bind("<Configure>", _on_canvas_configure)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.root.option_add('*selectBackground', '#3b82f6')
        self.root.option_add('*selectForeground', 'white')

        self._create_entry_context_menu()
        self._create_tree_context_menu()

        # Mouse Wheel Support
        def _on_mousewheel(event):
            try:
                focused = self.root.focus_get()
                curr = focused
                while curr:
                    try:
                        c_name = curr.winfo_class()
                        if c_name in ('Treeview', 'Text', 'Listbox'):
                            curr.yview_scroll(int(-1*(event.delta/120)*4.0), "units")
                            return "break"
                    except: pass
                    curr = curr.master
            except Exception: pass
            self.canvas.yview_scroll(int(-1*(event.delta/120)*2.0), "units")
            return "break"
        self.root.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        def _on_root_click(event):
            try:
                if "treeview" in str(event.widget).lower(): return
                w_class = event.widget.winfo_class()
                if w_class in ('Frame', 'TFrame', 'Label', 'TLabel', 'Canvas', 'Labelframe', 'TLabelframe'):
                    self.root.focus_set()
            except: pass
        self.root.bind_all("<Button-1>", _on_root_click, add="+")

        # [ROOT CAUSE FIX] Place Notebook DIRECTLY in root, NOT in Canvas scrollable_frame.
        # Canvas windows do not reliably propagate width to children.
        # Since PMI uses splitscreen (no scrolling needed), direct packing guarantees full width.
        
        # Hide Canvas & Scrollbar (no longer needed for layout)
        self.canvas.pack_forget()
        self.scrollbar.pack_forget()
        
        # ===== PACK ORDER: bottom-most items first =====
        
        # [UX] ESC Key Binding: Unfocus only (preserving content)
        self.root.bind_class("TEntry", "<Escape>", self._on_entry_esc)
        self.root.bind_class("Entry", "<Escape>", self._on_entry_esc)
        
        # 1. Status bar (very bottom)
        self.status_bar = tk.Frame(self.root, background="#e2e8f0", height=22)
        self.status_bar.pack(fill='x', side='bottom')
        self.status_bar.pack_propagate(False)
        tk.Label(self.status_bar, text="준비됨", font=("Malgun Gothic", 8), background="#e2e8f0", foreground="#64748b").pack(side='left', padx=10)
        tk.Label(self.status_bar, text=f"Build Version: v{APP_VERSION}", font=("Arial", 8), background="#e2e8f0", foreground="#94a3b8").pack(side='right', padx=10)

        # 2. Progress bar + Log (above status bar, compact)
        bottom_frame = tk.Frame(self.root, background="#f9fafb")
        bottom_frame.pack(fill='x', side='bottom', padx=5)

        self.progress = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill='x', pady=(5, 2))

        log_frame = tk.Frame(bottom_frame, background="#f9fafb")
        log_frame.pack(fill='x')

        self.status_log = tk.Text(log_frame, height=4, font=("Consolas", 8), state='disabled', background="#1e1e2e", foreground="#10b981", padx=5, pady=3)
        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.status_log.yview)
        self.status_log.configure(yscrollcommand=vsb.set)
        self.status_log.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        

        
        # 3. THEN pack Notebook (fills all remaining space)
        self.mode_notebook = ttk.Notebook(self.root, style="Main.TNotebook")
        self.mode_notebook.pack(fill='both', expand=True, padx=5, pady=(5, 0))
        
        # [NEW] Track mode change
        self.mode_notebook.bind("<<NotebookTabChanged>>", self._on_main_tab_changed)
        
        self.pmi_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.rt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.pt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.paut_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.photo_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")

        self.mode_notebook.add(self.pmi_mode_frame, text=" 🔬 PMI (OES) ")
        self.mode_notebook.add(self.rt_mode_frame, text=" 🔬 RT (Standard) ")
        

        
        self.pt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.mode_notebook.add(self.pt_mode_frame, text=" 🔬 Penetrant (PT) ")
        self.mt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.mode_notebook.add(self.mt_mode_frame, text=" 🔬 Magnetic (MT) ")
        self.mode_notebook.add(self.paut_mode_frame, text=" 🔬 Phased Array (PAUT) ")
        self.mode_notebook.add(self.photo_mode_frame, text=" 📸 사진대장 (Photo Log) ")

        # Setup sub-tabs lazily: only build the initially active tab upfront
        self._built_subtabs = set()
        self._subtab_builders = {
            0: lambda: self._setup_pmi_ui(self.pmi_mode_frame),
            1: lambda: self._setup_rt_ui(self.rt_mode_frame),
            2: lambda: self._setup_pt_ui(self.pt_mode_frame),
            3: lambda: self._setup_mt_ui(self.mt_mode_frame),
            4: lambda: self._setup_paut_ui(self.paut_mode_frame),
            5: lambda: self._setup_photo_log_ui(self.photo_mode_frame),
        }
        init_tab = getattr(self, '_initial_active_tab', 0)
        if not isinstance(init_tab, int) or init_tab not in self._subtab_builders:
            init_tab = 0
        self._ensure_subtab_built(init_tab)
        init_mode_map = {0: "PMI", 1: "RT", 2: "PT", 3: "MT", 4: "PAUT", 5: "PHOTO"}
        self.current_mode = init_mode_map.get(init_tab, "PMI")
        if self.current_mode in self._report_info_modes():
            self._load_report_info_for_mode(self.current_mode)
        elif self.current_mode == "PHOTO":
            self._sync_photo_report_info()
        if init_tab != 0:
            try:
                self.mode_notebook.select(init_tab)
            except Exception:
                pass

    def _add_sidebar_toggle(self, paned, sidebar, mode=None):
        if not getattr(self, 'is_embedded', False):
            return
        if mode is None and hasattr(self, '_infer_paned_mode'):
            mode = self._infer_paned_mode(paned)
        cfg = getattr(self, 'config', None)
        hidden_map = cfg.get('SIDEBAR_HIDDEN', {}) if isinstance(cfg, dict) else {}
        initially_hidden = bool(mode and isinstance(hidden_map, dict) and hidden_map.get(mode, False))

        previous_width = 425
        if mode and isinstance(cfg, dict):
            try:
                saved_ratio = float(cfg.get(f"{mode}_SASH_RATIO", 0.5))
                if 0.08 <= saved_ratio <= 0.92:
                    setattr(self, f"{mode.lower()}_pane_ratio", saved_ratio)
            except (TypeError, ValueError):
                pass

        toolbar = ttk.Frame(paned.master)
        toolbar.pack(fill='x', before=paned)

        def restore_width():
            if not paned.winfo_exists() or paned.tk.getboolean(paned.panecget(sidebar, 'hide')):
                return
            target_w = previous_width
            if mode and isinstance(cfg, dict):
                total_w = paned.winfo_width()
                ratio = getattr(self, f"{mode.lower()}_pane_ratio", cfg.get(f"{mode}_SASH_RATIO"))
                if total_w > 100 and ratio is not None:
                    try:
                        target_w = max(120, min(int(total_w * float(ratio)), total_w - 120))
                    except (TypeError, ValueError):
                        pass
            paned.sash_place(0, target_w, 0)
            paned._sash_initialized = True

        def toggle():
            nonlocal previous_width
            if paned.tk.getboolean(paned.panecget(sidebar, 'hide')):
                paned.paneconfigure(sidebar, hide=False)
                button.configure(text='◀ 왼쪽 패널 접기')
                if mode and isinstance(cfg, dict):
                    cfg.setdefault('SIDEBAR_HIDDEN', {})[mode] = False
                def _restore_and_save():
                    restore_width()
                    if callable(getattr(self, 'save_settings', None)):
                        self.save_settings()
                paned.after_idle(_restore_and_save)
            else:
                try:
                    previous_width = paned.sash_coord(0)[0]
                except Exception:
                    pass
                total_w = paned.winfo_width()
                if mode and isinstance(cfg, dict):
                    if total_w > 100 and previous_width > 0:
                        ratio = max(0.08, min(0.92, previous_width / total_w))
                        setattr(self, f"{mode.lower()}_pane_ratio", ratio)
                        cfg[f"{mode}_SASH_RATIO"] = ratio
                    cfg.setdefault('SIDEBAR_HIDDEN', {})[mode] = True
                paned.paneconfigure(sidebar, hide=True)
                button.configure(text='▶ 왼쪽 패널 펼치기')
                if callable(getattr(self, 'save_settings', None)):
                    self.save_settings()

        init_text = '▶ 왼쪽 패널 펼치기' if initially_hidden else '◀ 왼쪽 패널 접기'
        button = ttk.Button(toolbar, text=init_text, command=toggle)
        button.pack(side='left', padx=6, pady=3)
        if initially_hidden:
            paned.paneconfigure(sidebar, hide=True)
            paned._sash_initialized = True

    def col_to_num(self, col_str):
        """[NEW] 엑셀 열 이름(A, B, AA...)을 숫자 인덱스(1, 2, 27...)로 변환"""
        if not col_str: return 0
        s = str(col_str).strip().upper().replace("열", "").replace("행", "").strip()
        if s.isdigit(): return int(s)
        
        res = 0
        for i, char in enumerate(reversed(s)):
            if ord('A') <= ord(char) <= ord('Z'):
                res += (ord(char) - ord('A') + 1) * (26 ** i)
            else:
                return 0 # Invalid
        return res

    def _create_tree_context_menu(self):
        """미리보기 트리뷰용 공통 컨텍스트 메뉴 초기화"""
        self.ctx_menu = tk.Menu(self.root, tearoff=0)
        # 기본 항목 기입 (show_context_menu에서 command 동적 바인딩)
        self.ctx_menu.add_command(label="선택 항목 체크/해제 토글 (Toggle Check)")
        self.ctx_menu.add_command(label="현재 셀 내용만 복사 (Copy Cell)")
        self.ctx_menu.add_command(label="선택 행 전체 복사 (Copy Row)")
        self.ctx_menu.add_command(label="이 컬럼 전체 복사 (Copy Column)")
        self.ctx_menu.add_command(label="셀 내용 붙여넣기 (Paste)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="위로 이동 (Move Up)")
        self.ctx_menu.add_command(label="아래로 이동 (Move Down)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="행 추가 (Add Row)")
        self.ctx_menu.add_command(label="선택 삭제 (Delete Selected)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="선택 항목 ISO 병합 (Merge ISO)")
        self.ctx_menu.add_command(label="선택 항목 Joint 병합 (Merge Joint)")
        self.ctx_menu.add_command(label="선택 항목 ISO 병합 해제 (Ungroup ISO)")
        self.ctx_menu.add_command(label="선택 항목 Joint 병합 해제 (Ungroup Joint)")
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="선택 항목 일괄 변경 (Bulk Update)")
        # PAUT 전용 (기본은 숨김)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="컬럼 설정 (Column Manager)")

    def _create_entry_context_menu(self):
        self.entry_popup = tk.Menu(self.root, tearoff=0)
        self.entry_popup.add_command(label="복사 (Copy)", command=self._copy_entry)
        self.entry_popup.add_command(label="붙여넣기 (Paste)", command=self._paste_entry)
        self.entry_popup.add_separator()
        self.entry_popup.add_command(label="잘라내기 (Cut)", command=self._cut_entry)
        
        self.root.bind_class("Entry", "<Button-3>", self._show_entry_popup)
        self.root.bind_class("TEntry", "<Button-3>", self._show_entry_popup)

    def _show_entry_popup(self, event):
        self.current_entry = event.widget
        self._popup_selection_range = None
        if self.current_entry.select_present():
            self._popup_selection_range = (self.current_entry.index("sel.first"), self.current_entry.index("sel.last"))
        self.entry_popup.tk_popup(event.x_root, event.y_root)

    def _copy_entry(self):
        try:
            if getattr(self, 'current_entry', None):
                if self.current_entry.select_present():
                    first, last = self.current_entry.index("sel.first"), self.current_entry.index("sel.last")
                elif getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                else:
                    return
                text = self.current_entry.get()[first:last]
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                self.root.update()
        except Exception: pass

    def _paste_entry(self):
        try:
            text = self.root.clipboard_get()
            if getattr(self, 'current_entry', None):
                if getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                    self.current_entry.delete(first, last)
                    self.current_entry.insert(first, text)
                elif self.current_entry.select_present():
                    self.current_entry.delete("sel.first", "sel.last")
                    self.current_entry.insert(tk.INSERT, text)
                else:
                    self.current_entry.insert(tk.INSERT, text)
        except Exception: pass

    def _cut_entry(self):
        self._copy_entry()
        try:
            if getattr(self, 'current_entry', None):
                if getattr(self, '_popup_selection_range', None):
                    first, last = self._popup_selection_range
                    self.current_entry.delete(first, last)
                elif self.current_entry.select_present():
                    self.current_entry.delete("sel.first", "sel.last")
        except Exception: pass

    def _excel_anchor_to_px(self, anchor, scale):
        """Roughly map Excel cell to A4 pixel coordinates."""
        import re
        m = re.match(r"([A-Z]+)(\d+)", anchor.upper())
        if not m: return 0, 0
        
        col_str, row_str = m.groups()
        # Col to Num (A=0, B=1...)
        col = 0
        for char in col_str: col = col * 26 + (ord(char) - ord('A'))
        row = int(row_str) - 1
        
        # Rough estimates for A4 grid (210mm / 10 cols, 297mm / 50 rows)
        # Adjusting to feel more like typical report templates
        x = (col * 15) * scale # 15mm per col
        y = (row * 6) * scale  # 6mm per row
        return x, y

    def _create_setting_grid(self, parent, context):
        # [REFINED] More organized layout for logo settings
        items = [
            ("SITCO 로고", f"SITCO_{context}"), 
            ("서울검사 로고", f"SEOUL_{context}"), 
            ("바닥글 우측 (Footer Right)", f"FOOTER_{context}"), 
            ("바닥글 좌측 (Footer Left)", f"FOOTER_PT_{context}")
        ]
        
        next_row = 0
        for i, (label, key_prefix) in enumerate(items):
            # [HYPER-REACTIVE] Force columnspan=1 to match weighted col 0 of parent
            short_label = label.split(" ")[0].replace("로고", "")
            block = tk.LabelFrame(parent, text=f" {short_label} ", padx=2, pady=1, background="#ffffff", font=("Malgun Gothic", 8, "bold"))
            # [FIX] Offset row by 1 to leave room for Gapji Info at row 0 if present
            block.grid(row=i + 1, column=0, sticky='ew', pady=1, padx=2)
            
            frame = tk.Frame(block, background="#ffffff")
            frame.pack(fill='x')
            frame.columnconfigure(1, weight=5, minsize=0) # Path entry weight
            for c in [4, 6, 8, 10]: frame.columnconfigure(c, weight=1, minsize=0) # Coords weight
            
            tk.Label(frame, text="P:", width=2, anchor='e', background="#ffffff", font=("Malgun Gothic", 7)).grid(row=0, column=0)
            v_path = tk.StringVar(value=self.config.get(f"{key_prefix}_PATH", ""))
            ttk.Entry(frame, textvariable=v_path, width=15, exportselection=False).grid(row=0, column=1, sticky='ew', padx=1) # [FIX] Width 1 -> 15
            ttk.Button(frame, text="..", width=2, command=lambda v=v_path: self._browse_file(v, [("Images", "*.png;*.jpg;*.jpeg")])).grid(row=0, column=2, padx=1)
            self.setting_vars[f"{key_prefix}_PATH"] = v_path
            
            for idx, (coord, key_suffix) in enumerate([("X", "X"), ("Y", "Y"), ("W", "W"), ("H", "H")]):
                tk.Label(frame, text=f"{coord}:", width=1, anchor='e', background="#ffffff", font=("Malgun Gothic", 7)).grid(row=0, column=3+idx*2)
                v = tk.StringVar(value=str(self.config.get(f"{key_prefix}_{key_suffix}", "0.0")))
                ttk.Entry(frame, textvariable=v, width=6, exportselection=False).grid(row=0, column=4+idx*2, sticky='ew', padx=1) # [FIX] Width 1 -> 6
                self.setting_vars[f"{key_prefix}_{key_suffix}"] = v
            next_row = i + 1
        return next_row
            
    def _create_margin_settings(self, parent, context, use_pack=True):
        # [FIX] Adaptive layout to avoid TclError (grid vs pack conflict)
        frame = ttk.LabelFrame(parent, text=" 인쇄 및 여백 (Margins) ", padding=2)
        if use_pack:
            frame.pack(fill='x', pady=(5, 0))
        else:
            # Inside Notebook tabs, we usually use grid. 
            # We assume it's placed at the bottom of the existing grid.
            frame.grid(row=100, column=0, columnspan=10, sticky='ew', pady=(10, 0))
        
        # Force 0 minsize for elasticity
        for col in range(10): frame.columnconfigure(col, minsize=0)
        # [FIX] Distribute expansion only to entries
        for col in [0, 2, 4, 6, 8]: frame.columnconfigure(col, weight=0) # Labels
        for col in [1, 3, 5, 7, 9]: frame.columnconfigure(col, weight=1) # Entries

        m_items = [("T", "TOP"), ("B", "BOTTOM"), ("L", "LEFT"), ("R", "RIGHT")]
        for i, (sh_lbl, key) in enumerate(m_items):
            ttk.Label(frame, text=f"{sh_lbl}:", font=("Arial", 7)).grid(row=0, column=i*2, sticky='e', padx=1)
            v = tk.StringVar(value=str(self.config.get(f"MARGIN_{context}_{key}", "0.2")))
            ent = ttk.Entry(frame, textvariable=v, width=1)
            ent.grid(row=0, column=i*2+1, sticky='ew', padx=1)
            ent.bind("<Return>", lambda e: self.root.focus_set())
            self.setting_vars[f"MARGIN_{context}_{key}"] = v
            
        ttk.Label(frame, text="S:", font=("Arial", 7)).grid(row=0, column=8, sticky='e', padx=1)
        v_s = tk.StringVar(value=str(self.config.get(f"PRINT_SCALE_{context}", "95")))
        ent_s = ttk.Entry(frame, textvariable=v_s, width=1)
        ent_s.grid(row=0, column=9, sticky='ew', padx=1)
        ent_s.bind("<Return>", lambda e: self.root.focus_set())
        self.setting_vars[f"PRINT_SCALE_{context}"] = v_s

        # [NEW] Print Area Row
        area_f = ttk.Frame(frame)
        area_f.grid(row=1, column=0, columnspan=10, sticky='ew', pady=(2, 0))
        ttk.Label(area_f, text="Area:", font=("Arial", 7)).pack(side='left', padx=1)
        v_a = tk.StringVar(value=str(self.config.get(f"PRINT_AREA_{context}", "")))
        ent_a = ttk.Entry(area_f, textvariable=v_a)
        ent_a.pack(side='left', fill='x', expand=True, padx=1)
        self.setting_vars[f"PRINT_AREA_{context}"] = v_a
        
        # Adjusters Row
        sub = ttk.Frame(frame)
        sub.grid(row=2, column=0, columnspan=10, sticky='ew', pady=2)
        for c in [1, 3, 5, 7]: sub.columnconfigure(c, weight=1, minsize=0)
        for c in [0, 2, 4, 6]: sub.columnconfigure(c, weight=0, minsize=0)
        
        labels = [("Rows:", f"CUSTOM_ROWS_{context}"), ("H:", f"CUSTOM_ROW_HEIGHT_{context}"), 
                  ("Cols:", f"CUSTOM_COLS_{context}"), ("W:", f"CUSTOM_COL_WIDTH_{context}")]
        
        for i, (lbl, key) in enumerate(labels):
            ttk.Label(sub, text=lbl, font=("Arial", 7)).grid(row=0, column=i*2, sticky='e')
            defaultv = "16.5" if "HEIGHT" in key else ("10.0" if "WIDTH" in key else "")
            v = tk.StringVar(value=str(self.config.get(key, defaultv)))
            ent = ttk.Entry(sub, textvariable=v, width=1)
            ent.grid(row=0, column=i*2+1, sticky='ew', padx=1)
            ent.bind("<FocusOut>", lambda e: self.save_settings())
            ent.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key] = v

    def _create_column_mapping_ui(self, parent, mode, mapping_items):
        """[NEW] 전 모듈 공용 엑셀 컬럼 매핑 및 이름 설정 UI (스크롤바 추가)"""
        # Outer Frame for Scrollbar
        outer_frame = tk.Frame(parent, background="#f9fafb")
        outer_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(outer_frame, background="#f9fafb", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, background="#f9fafb")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        container = ttk.LabelFrame(scrollable_frame, text=f" {mode} 성적서 및 미리보기 컬럼 설정 ", padding=10)
        container.pack(fill='both', expand=True, pady=5, padx=10)
        
        # Grid settings
        container.columnconfigure(2, weight=1)
        container.columnconfigure(5, weight=1)
        
        # Header labels
        ttk.Label(container, text="항목", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=0, sticky='w')
        ttk.Label(container, text="엑셀열 | 표시이름", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=1, columnspan=2, sticky='w')
        ttk.Label(container, text="항목", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=3, sticky='w')
        ttk.Label(container, text="엑셀열 | 표시이름", font=("Malgun Gothic", 8, "bold")).grid(row=0, column=4, columnspan=2, sticky='w')
        
        last_row = 0
        for i, (label, key_idx, def_idx, key_name, def_name, internal_id) in enumerate(mapping_items):
            row = (i // 2) + 1
            last_row = max(last_row, row)
            col_offset = (i % 2) * 3
            
            # 1. Label
            ttk.Label(container, text=label, font=("Malgun Gothic", 8)).grid(row=row, column=col_offset, sticky='e', padx=2, pady=2)
            
            # 2. Excel Col #
            v_idx = tk.StringVar(value=str(self.config.get(key_idx, def_idx)))
            ent_idx = ttk.Entry(container, textvariable=v_idx, width=4)
            ent_idx.grid(row=row, column=col_offset+1, sticky='w', padx=2, pady=2)
            ent_idx.bind("<FocusOut>", lambda e: self.save_settings())
            ent_idx.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key_idx] = v_idx
            
            # [NEW] Trace to update config in real-time so changes are active immediately
            v_idx.trace_add("write", lambda *args, var=v_idx, k=key_idx: self.config.update({k: var.get()}))
            
            # 3. Display Name
            v_name = tk.StringVar(value=str(self.config.get(key_name, def_name)))
            ent_name = ttk.Entry(container, textvariable=v_name, width=12)
            ent_name.grid(row=row, column=col_offset+2, sticky='ew', padx=2, pady=2)
            ent_name.bind("<FocusOut>", lambda e: self.save_settings())
            ent_name.bind("<Return>", lambda e: [self.root.focus_set(), self.save_settings()])
            self.setting_vars[key_name] = v_name
            
            # Trace to update config and preview tree heading in real-time
            v_name.trace_add("write", lambda *args, tid=internal_id, m=mode, var=v_name, kn=key_name: self._update_mode_heading(m, tid, var, kn))

        ttk.Label(container, text="* 엑셀열: 알파벳(A, B...) 또는 숫자(1, 2...) 입력 (0=제외)", foreground="gray", font=("Malgun Gothic", 8)).grid(row=last_row+1, column=0, columnspan=6, pady=(10, 2))
        
        # [NEW] Save Button
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=last_row+2, column=0, columnspan=6, pady=5)
        ttk.Button(btn_frame, text="💾 현재 컬럼 설정 저장", command=self.manual_save_settings).pack()

    def _update_mode_heading(self, mode, internal_id, var, key_name):
        """[NEW] 특정 모드의 미리보기 헤더 이름을 실시간으로 업데이트 및 설정 저장"""
        try:
            new_text = var.get()
            # 1. Update config immediately so it survives refreshes
            self.config[key_name] = new_text
            
            # 2. Update existing tree UI
            tree = None
            if mode == "PMI": tree = self.preview_tree
            elif mode == "RT": tree = self.rt_preview_tree
            elif mode == "KOGAS": tree = self.kogas_preview_tree
            elif mode == "PT": tree = self.pt_preview_tree
            
            if tree:
                try: tree.heading(internal_id, text=new_text)
                except: pass
        except Exception:
            pass

    def _show_paut_check_menu(self):
        tree = self.paut_preview_tree
        try:
            self.paut_check_menu.tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self.paut_check_menu.grab_release()

    def _show_mt_check_menu(self):
        tree = self.mt_preview_tree
        try:
            self.mt_check_menu.tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self.mt_check_menu.grab_release()


    def _show_rt_check_menu(self, mode):
        tree = self._get_mode_info(mode)[0]
        try:
            self._rt_check_menus[mode].tk_popup(tree.winfo_pointerx(), tree.winfo_rooty() + 32)
        finally:
            self._rt_check_menus[mode].grab_release()


    def _refresh_rt_filter_headers(self, mode):
        tree = self._get_mode_info(mode)[0]
        columns = list(tree['columns'])
        for key, flag, candidates in [('Date', 'date_filtered', ('Date',)),
                                       ('Dwg', 'drawing_filtered', ('Dwg', 'ISO/Dwg', 'ISO/DWG'))]:
            column = next((value for value in candidates if value in columns), None)
            if column is None:
                continue
            values = {str(item.get(key, '') or '').strip() for item in self._get_mode_info(mode)[2]}
            active = {str(item.get(key, '') or '').strip() for item in self._get_mode_info(mode)[2] if item.get(flag, True)}
            suffix = f' ({len(active)}개)' if values != active else ''
            tree.heading(column, text=('Drawing No.' if key == 'Dwg' else key) + ' ▾' + suffix,
                         command=lambda k=key, f=flag, c=column: self._choose_rt_filter(mode, k, f, c))
        if mode in getattr(self, '_rt_check_menus', {}) and 'selected' in columns:
            tree.heading('selected', text='V ▾', command=lambda: self._show_rt_check_menu(mode))


    def _choose_rt_filter(self, mode, key, flag, column):
        states = {}
        for item in self._get_mode_info(mode)[2]:
            states.setdefault(str(item.get(key, '') or '').strip(), bool(item.get(flag, True)))
        states = dict(sorted(states.items()))
        panel = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(panel, column, 560 if key == 'Dwg' else 420, 420,
                                            tree=self._get_mode_info(mode)[0])
        ttk.Label(panel, text=f'{key} — 여러 항목 체크 후 적용').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        ttk.Entry(panel, textvariable=search).pack(fill='x', padx=10)
        frame = ttk.Frame(panel)
        frame.pack(fill='both', expand=True, padx=10, pady=6)
        choices = tk.Listbox(frame, font=('Malgun Gothic', 10), exportselection=False)
        choices.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(frame, command=choices.yview)
        scrollbar.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scrollbar.set)
        visible = []
        def refresh(*_args):
            visible[:] = [value for value in states if search.get().strip().casefold() in value.casefold()]
            choices.delete(0, tk.END)
            for value in visible:
                choices.insert(tk.END, ('[v] ' if states[value] else '[ ] ') + (value or '(비어 있음)'))
        def toggle(event):
            index = choices.nearest(event.y) if event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                value = visible[index]
                states[value] = not states[value]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for value in states:
                states[value] = state
            refresh()
        def apply():
            for item in self._get_mode_info(mode)[2]:
                item[flag] = states.get(str(item.get(key, '') or '').strip(), False)
            self.update_date_listbox(mode=mode)
            self.populate_preview(self._get_mode_info(mode)[2], switch_tab=False, mode=mode)
            panel.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', padx=8, pady=6)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', panel.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(panel, text='정렬 ↑ / ↓', command=lambda: self.sort_by_column(key, mode=mode)).pack(fill='x', padx=10, pady=4)


    def on_tree_select(self, event):
        """항목 선택 시 체크박스 상태와 동기화 (현재는 드래그 안정성을 위해 드래그 중에는 무시)"""
        pass

    def toggle_item_selection(self, item_id, mode="PMI"):
        """단일 항목의 체크 상태를 토글"""
        tree, idx_map, data, _ = self._get_mode_info(mode)
        
        view_idx = tree.index(item_id)
        if 0 <= view_idx < len(idx_map):
            actual_idx = idx_map[view_idx]
            data[actual_idx]['selected'] = not data[actual_idx].get('selected', True)
            self.populate_preview(data, switch_tab=False, mode=mode)
            
            # 선택 상태 복구
            num_items = len(tree.get_children())
            if view_idx < num_items:
                new_item = tree.get_children()[view_idx]
                tree.selection_set(new_item)
                tree.focus(new_item)

    def toggle_selected_items(self, mode="PMI"):
        """선택된 모든 항목의 체크 상태를 토글"""
        tree, idx_map, data, _ = self._get_mode_info(mode)
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        for item_id in selected_ids:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['selected'] = not data[actual_idx].get('selected', True)
        
        self.populate_preview(data, switch_tab=False, mode=mode)
        # 선택 보존
        for sid in selected_ids:
            try: tree.selection_add(sid)
            except: pass

    def on_tree_click_checkbox(self, event):
        """Reserved for future coordinate-based checkbox toggle if needed without blocking drag."""
        pass

    def on_tree_double_click(self, event, mode="PMI"):
        """셀(Cell) 더블클릭 시 엑셀처럼 그 자리에서 바로 수동 편집 가능하게 함"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, col_keys = mode_info

        region = tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y)
        if not item_id or not column: return
        
        # [DYNAMIC] 컬럼 인덱스를 통해 키 추출
        try:
            col_idx = int(column.replace('#', '')) - 1
            if col_idx < 0 or col_idx >= len(col_keys): return
            key = col_keys[col_idx]
        except: return
        
        # 수정 차단 컬럼 (V, 상태 아이콘 등)
        if key in ["selected", "_status", "V", "ST"]:
            if key in ["selected", "V"]: self.toggle_item_selection(item_id, mode)
            return

        # 현재 값 가져오기
        x, y, w, h = tree.bbox(item_id, column)
        view_idx = tree.index(item_id)
        if not (0 <= view_idx < len(idx_map)): return
        actual_idx = idx_map[view_idx]
        
        curr_val = data_list[actual_idx].get(key, "")
        numeric_keys = ["Th'k(mm)", "Start", "End", "Length(mm)", "Upper", "Lower", "Height(mm)", "TCG MAX(%)", "a/l", "a/t"]
        if (key in numeric_keys or (key not in ["No", "Joint No.", "Line No.", "Joint", "Loc", "Grade", "Date", "Dwg", "ISO", "ISO/DWG", "Type of Flaw", "Evaluation", "Remarks", "Result", "Location", "Nature", "selected", "_status", "ST", "V"])) and self.to_float(curr_val) > 0:
            curr_val = f"{self.to_float(curr_val):.2f}"
            
        entry = ttk.Entry(tree, font=("Malgun Gothic", 10), exportselection=True)
        entry.insert(0, str(curr_val))
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        
        def finish_edit(e=None):
            if not entry.winfo_exists(): return
            new_val = entry.get().strip()
            
            # [PMI 특화] 함량 소수점 처리 및 판정 재계산
            if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                f_val = self.to_float(new_val)
                data_list[actual_idx][key] = f"{f_val:.2f}" if f_val > 0 else ""
                # 함량 변경 시 등급 판정 재수행
                new_grade = self.check_material_grade(data_list[actual_idx])
                if new_grade: data_list[actual_idx]['Grade'] = new_grade
            elif mode == "PAUT" and key in ["Th'k(mm)", "Height(mm)", "Length(mm)", "Type of Flaw"]:
                data_list[actual_idx][key] = new_val
                # [PAUT] Recalculate Evaluation and Ratios
                t = self.to_float(data_list[actual_idx].get('Th\'k(mm)', 0))
                h = self.to_float(data_list[actual_idx].get('Height(mm)', 0))
                l = self.to_float(data_list[actual_idx].get('Length(mm)', 0))
                d = self.to_float(data_list[actual_idx].get('Upper', 0))
                nat = data_list[actual_idx].get('Type of Flaw', 'Slag')
                eval_mode = self.paut_eval_mode.get()
                
                res, _ = self.evaluate_paut_flaw(t, h, l, d, nat, mode=eval_mode)
                data_list[actual_idx]['Evaluation'] = res
                if l > 0 and h > 0: data_list[actual_idx]['a/l'] = f"{h/l:.3f}"
                else: data_list[actual_idx]['a/l'] = ""
                if t > 0 and h > 0: data_list[actual_idx]['a/t'] = f"{h/t:.3f}"
                else: data_list[actual_idx]['a/t'] = ""
                
                # [NEW] 합격 시 결함 정보 지우기 (단, 수동으로 결함 종류를 입력하는 중이면 유지)
                if "Accept" in res and key != "Type of Flaw":
                    data_list[actual_idx]['Type of Flaw'] = ""
                    data_list[actual_idx]['a/l'] = ""
                    data_list[actual_idx]['a/t'] = ""
            else:
                data_list[actual_idx][key] = new_val

            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"📝 {mode} {key} 수정: {new_val}")
            entry.destroy()

        entry.bind("<Return>", finish_edit)
        entry.bind("<FocusOut>", lambda e: finish_edit())
        entry.bind("<Escape>", lambda e: entry.destroy())

    def on_delete_key(self, event, mode="PMI"):
        """Delete 키 선택 시 현재 선택된 셀(들)의 내용을 지움 (엑셀 방식)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, col_keys = mode_info
        
        selected = tree.selection()
        if not selected: return
        
        # 드래그된 컬럼 범위 계산
        s_col = getattr(self, 'start_col', None)
        e_col = getattr(self, 'end_col', None)
        l_col = getattr(self, 'last_clicked_col', None)
        
        if not s_col or not e_col or s_col == e_col:
            target_col_ids = [l_col or s_col or "#4"]
        else:
            try:
                s_idx = int(s_col.replace('#', '')) - 1
                e_idx = int(e_col.replace('#', '')) - 1
                col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
            except:
                target_col_ids = [l_col or s_col or "#4"]
        
        key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
        count = 0
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                for c_id in target_col_ids:
                    key = key_map.get(c_id)
                    if key and key not in ["selected", "_status"]:
                        data_list[actual_idx][key] = ""
                        count += 1
        
        if count > 0:
            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"🧹 {mode} 데이터 {count}건 내용 삭제 완료")

    def show_context_menu(self, event, mode="PMI"):
        """우클릭 컨텍스트 메뉴 표시 및 현재 모드에 따라 명령 구성"""
        tree = self._get_mode_info(mode)[0]
        col = tree.identify_column(event.x)
        self.last_clicked_col = col 
        item_id = tree.identify_row(event.y)
        
        # 선택 상태 보정 (우클릭한 항목으로 선택 변경)
        if item_id:
            if item_id not in tree.selection():
                tree.selection_set(item_id)
            tree.focus(item_id)

        # 메뉴 명령 동적 업데이트
        self.ctx_menu.entryconfigure("선택 항목 체크/해제 토글 (Toggle Check)", command=lambda: self.toggle_selected_items(mode))
        self.ctx_menu.entryconfigure("현재 셀 내용만 복사 (Copy Cell)", command=lambda: self.copy_cell(mode, target="cell"))
        self.ctx_menu.entryconfigure("선택 행 전체 복사 (Copy Row)", command=lambda: self.copy_cell(mode, target="row"))
        self.ctx_menu.entryconfigure("이 컬럼 전체 복사 (Copy Column)", command=lambda: self.copy_cell(mode, target="column"))
        self.ctx_menu.entryconfigure("셀 내용 붙여넣기 (Paste)", command=lambda: self.paste_cell(mode))
        
        self.ctx_menu.entryconfigure("위로 이동 (Move Up)", command=lambda: self.move_item(-1, mode))
        self.ctx_menu.entryconfigure("아래로 이동 (Move Down)", command=lambda: self.move_item(1, mode))
        
        self.ctx_menu.entryconfigure("행 추가 (Add Row)", command=lambda: self.add_item(mode))
        self.ctx_menu.entryconfigure("선택 삭제 (Delete Selected)", command=lambda: self.delete_item(mode))
        
        # PMI 전용/특화 기능 제어
        is_pmi = (mode == "PMI")
        self.ctx_menu.entryconfigure("선택 항목 ISO 병합 (Merge ISO)", command=lambda: self.merge_selected_iso(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 Joint 병합 (Merge Joint)", command=lambda: self.merge_selected_joint(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 ISO 병합 해제 (Ungroup ISO)", command=lambda: self.ungroup_selected_iso(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 Joint 병합 해제 (Ungroup Joint)", command=lambda: self.ungroup_selected_joint(mode), state="normal" if is_pmi else "disabled")
        self.ctx_menu.entryconfigure("선택 항목 일괄 변경 (Bulk Update)", command=lambda: self.show_bulk_update_dialog(mode))

        # 컬럼 설정 기능 (모든 모드 지원)
        self.ctx_menu.entryconfigure("컬럼 설정 (Column Manager)", command=lambda: self.manage_columns(mode))
        self.ctx_menu.entryconfigure("컬럼 설정 (Column Manager)", state="normal")

        self.ctx_menu.post(event.x_root, event.y_root)

    def update_date_listbox(self, mode="PMI"):
        if mode == "RT":
            listbox = self.rt_date_listbox
            data = self.rt_extracted_data
        elif mode == "KOGAS":
            listbox = self.kogas_date_listbox
            data = self.kogas_extracted_data
        elif mode == "PT":
            listbox = self.pt_date_listbox
            data = self.pt_extracted_data
        elif mode == "MT":
            listbox = self.mt_date_listbox
            data = self.mt_extracted_data
        elif mode == "PAUT":
            listbox = self.paut_date_listbox
            data = self.paut_extracted_data
        else:
            listbox = self.date_listbox
            data = self.extracted_data
        
        if listbox is None or not listbox.winfo_exists(): return
        
        listbox.delete(0, tk.END)
        # [NEW] 데이터 기반으로 날짜별 선택 상태(date_filtered) 수집
        date_status = {}
        for item in data:
            dt = str(item.get('Date', '') or '').strip()
            if not dt:
                continue
            if dt not in date_status:
                date_status[dt] = item.get('date_filtered', True)
        
        dates = sorted(list(date_status.keys()))
        for d in dates:
            is_v = date_status.get(d, True)
            prefix = "[v]" if is_v else "[ ]"
            listbox.insert(tk.END, f"{prefix} {d}")
            if is_v: listbox.select_set(tk.END)
        self.log(
            f"📅 {mode} 날짜 필터 갱신: {listbox.size()}개 "
            f"({', '.join(dates) if dates else '날짜 없음'})"
        )
        if mode == "PMI":
            self.update_pmi_loc_listbox()
        elif mode == "PAUT":
            self.update_paut_filter_listboxes()
        elif mode == 'MT':
            self._refresh_mt_filter_headers()
        elif mode in ('RT', 'KOGAS'):
            self._refresh_rt_filter_headers(mode)

    def _position_paut_filter_dropdown(self, popup, column, width, height, tree=None):
        """Place an in-app filter panel directly below its column header."""
        previous = getattr(self, '_paut_filter_dropdown', None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
        self._paut_filter_dropdown = popup
        tree = tree if tree is not None else self.paut_preview_tree
        columns = list(tree['displaycolumns'])
        if columns == ['#all']:
            columns = list(tree['columns'])
        widths = [tree.column(key, 'width') for key in columns]
        offset = sum(widths[:columns.index(column)]) if column in columns else 0
        offset -= int(tree.xview()[0] * sum(widths))
        app_width, app_height = self.root.winfo_width(), self.root.winfo_height()
        x = max(0, min(tree.winfo_rootx() - self.root.winfo_rootx() + offset, app_width - width))
        header_height = 32
        found_header = False
        for row_y in range(min(100, tree.winfo_height())):
            region = tree.identify_region(min(20, max(1, tree.winfo_width() - 2)), row_y)
            if region == 'heading':
                found_header = True
            elif found_header:
                header_height = row_y
                break
        y = max(0, tree.winfo_rooty() - self.root.winfo_rooty() + header_height)
        width, height = min(width, app_width), min(height, max(1, app_height - y))
        def outside_click(event):
            if not (popup.winfo_rootx() <= event.x_root < popup.winfo_rootx() + popup.winfo_width()
                    and popup.winfo_rooty() <= event.y_root < popup.winfo_rooty() + popup.winfo_height()):
                popup.destroy()
                return 'break'
        popup.bind('<ButtonPress-1>', outside_click)
        popup.bind('<Escape>', lambda _event: popup.destroy())
        def show_dropdown():
            if not popup.winfo_exists():
                return
            popup.place(x=x, y=y, width=width, height=height)
            popup.update_idletasks()
            popup.lift()
            popup.focus_force()
            popup.grab_set()
        # Build the contents before displaying the panel within the main window.
        popup.after_idle(show_dropdown)

    def _refresh_mt_filter_headers(self):
        tree = self.mt_preview_tree
        columns = list(tree['columns'])
        for key, flag, candidates in [('Date', 'date_filtered', ('Date',)),
                                       ('Dwg', 'drawing_filtered', ('Dwg', 'ISO/Dwg', 'ISO/DWG'))]:
            column = next((value for value in candidates if value in columns), None)
            if column is None:
                continue
            values = {str(item.get(key, '') or '').strip() for item in self.mt_extracted_data}
            active = {str(item.get(key, '') or '').strip() for item in self.mt_extracted_data if item.get(flag, True)}
            suffix = f' ({len(active)}개)' if values != active else ''
            tree.heading(column, text=key + ' ▾' + suffix,
                         command=lambda k=key, f=flag, c=column: self._choose_mt_filter(k, f, c))
        if hasattr(self, 'mt_check_menu') and 'V' in columns:
            tree.heading('V', text='V ▾', command=self._show_mt_check_menu)

    def _choose_mt_filter(self, key, flag, column):
        states = {}
        for item in self.mt_extracted_data:
            states.setdefault(str(item.get(key, '') or '').strip(), bool(item.get(flag, True)))
        states = dict(sorted(states.items()))
        panel = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(panel, column, 560 if key == 'Dwg' else 420, 420,
                                            tree=self.mt_preview_tree)
        ttk.Label(panel, text=f'{key} — 여러 항목 체크 후 적용').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        ttk.Entry(panel, textvariable=search).pack(fill='x', padx=10)
        frame = ttk.Frame(panel)
        frame.pack(fill='both', expand=True, padx=10, pady=6)
        choices = tk.Listbox(frame, font=('Malgun Gothic', 10), exportselection=False)
        choices.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(frame, command=choices.yview)
        scrollbar.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scrollbar.set)
        visible = []
        def refresh(*_args):
            visible[:] = [value for value in states if search.get().strip().casefold() in value.casefold()]
            choices.delete(0, tk.END)
            for value in visible:
                choices.insert(tk.END, ('[v] ' if states[value] else '[ ] ') + (value or '(비어 있음)'))
        def toggle(event):
            index = choices.nearest(event.y) if event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                value = visible[index]
                states[value] = not states[value]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for value in states:
                states[value] = state
            refresh()
        def apply():
            for item in self.mt_extracted_data:
                item[flag] = states.get(str(item.get(key, '') or '').strip(), False)
            self.update_date_listbox(mode='MT')
            self.populate_preview(self.mt_extracted_data, switch_tab=False, mode='MT')
            panel.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', padx=8, pady=6)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', panel.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(panel, text='정렬 ↑ / ↓', command=lambda: self.sort_by_column(key, mode='MT')).pack(fill='x', padx=10, pady=4)

    def _choose_paut_dates(self):
        """Choose multiple dates without changing independent line/size filters."""
        states = {}
        for item in self.paut_extracted_data:
            date = str(item.get('Date', '') or '').strip()
            states.setdefault(date, bool(item.get('date_filtered', True)))
        states = dict(sorted(states.items()))
        tree = self.paut_preview_tree
        dialog = tk.Frame(self.root, background='white', borderwidth=1, relief='solid')
        self._position_paut_filter_dropdown(dialog, 'Date', 420, 440)
        ttk.Label(dialog, text='날짜를 여러 개 체크한 후 적용하세요.').pack(anchor='w', padx=10, pady=8)
        search = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=search)
        entry.pack(fill='x', padx=10)
        frame = ttk.Frame(dialog)
        frame.pack(fill='both', expand=True, padx=10, pady=8)
        choices = tk.Listbox(frame, exportselection=False, font=('Malgun Gothic', 10))
        choices.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(frame, command=choices.yview)
        scroll.pack(side='right', fill='y')
        choices.configure(yscrollcommand=scroll.set)
        visible = []
        def refresh(*_args):
            visible[:] = [date for date in states if search.get().strip() in date]
            choices.delete(0, tk.END)
            for date in visible:
                choices.insert(tk.END, ('[v] ' if states[date] else '[ ] ') + (date or '(날짜 없음)'))
        def toggle(event=None):
            index = choices.nearest(event.y) if event and event.keysym != 'space' else (
                choices.curselection()[0] if choices.curselection() else -1)
            if 0 <= index < len(visible):
                date = visible[index]
                states[date] = not states[date]
                refresh()
                choices.selection_set(index)
            return 'break'
        def set_all(state):
            for date in states:
                states[date] = state
            refresh()
        def apply():
            for item in self.paut_extracted_data:
                item['date_filtered'] = states.get(str(item.get('Date', '') or '').strip(), False)
            self.update_date_listbox(mode='PAUT')
            self.populate_preview(self.paut_extracted_data, switch_tab=False, mode='PAUT')
            self._sync_paut_photo_page_numbers()
            dialog.destroy()
        choices.bind('<ButtonRelease-1>', toggle)
        choices.bind('<space>', toggle)
        search.trace_add('write', refresh)
        refresh()
        buttons = ttk.Frame(dialog)
        buttons.pack(fill='x', padx=10, pady=4)
        for label, command in [('전체 선택', lambda: set_all(True)), ('전체 해제', lambda: set_all(False)),
                               ('적용', apply), ('취소', dialog.destroy)]:
            ttk.Button(buttons, text=label, command=command).pack(side='left', padx=2)
        ttk.Button(dialog, text='날짜 정렬 ↑ / ↓',
                   command=lambda: self.sort_by_column('Date', mode='PAUT')).pack(fill='x', padx=10, pady=(0, 8))
        entry.focus_set()
        dialog.bind('<Escape>', lambda _event: dialog.destroy())

    def update_paut_filter_listboxes(self):
        """Refresh PAUT Line No./Size choices while preserving check states."""
        line_listbox = getattr(self, 'paut_line_listbox', None)
        size_listbox = getattr(self, 'paut_size_listbox', None)
        if not line_listbox or not size_listbox:
            return
        if not line_listbox.winfo_exists() or not size_listbox.winfo_exists():
            return

        def refresh(listbox, key, flag):
            previous = {}
            for idx in range(listbox.size()):
                text = listbox.get(idx)
                value = text.replace("[v] ", "").replace("[ ] ", "")
                previous[value] = text.startswith("[v] ")

            values = sorted({
                str(item.get(key, '') or '').strip()
                for item in self.paut_extracted_data
                if str(item.get(key, '') or '').strip()
            })
            listbox.delete(0, tk.END)
            active_values = set()
            for value in values:
                checked = previous.get(value, True)
                listbox.insert(tk.END, ("[v] " if checked else "[ ] ") + value)
                if checked:
                    active_values.add(value)
            for item in self.paut_extracted_data:
                item_value = str(item.get(key, '') or '').strip()
                item[flag] = item_value in active_values

        refresh(line_listbox, 'Line No.', 'line_filtered')
        refresh(size_listbox, 'Size', 'size_filtered')
        dates = {str(item.get('Date', '') or '').strip() for item in self.paut_extracted_data}
        active_dates = {str(item.get('Date', '') or '').strip() for item in self.paut_extracted_data
                        if item.get('date_filtered', True)}
        suffix = f' ({len(active_dates)}개)' if active_dates != dates else ''
        self.paut_preview_tree.heading('Date', text='Date ▾' + suffix, command=self._choose_paut_dates)
        if hasattr(self, 'paut_check_menu'):
            self.paut_preview_tree.heading('V', text='V ▾', command=self._show_paut_check_menu)
        if hasattr(self, '_choose_paut_line'):
            checked = sum(line_listbox.get(i).startswith('[v] ') for i in range(line_listbox.size()))
            label = self.config.get('PAUT_NAME_LINE', 'Line No.')
            suffix = f' ({checked}개)' if checked != line_listbox.size() else ''
            self.paut_preview_tree.heading('Line No.', text=label + ' ▾' + suffix,
                                           command=self._choose_paut_line)

    def _get_mode_info(self, mode):
        """Helper to get core UI/Data objects for a specific module mode."""
        if mode == "MT":
            return (self.mt_preview_tree, self.mt_item_idx_map, self.mt_extracted_data, self.mt_column_keys)
        if mode == "PMI":
            return (self.preview_tree, self.item_idx_map, self.extracted_data, self.column_keys)
        elif mode == "RT":
            return (self.rt_preview_tree, self.rt_item_idx_map, self.rt_extracted_data, self.rt_column_keys)
        elif mode == "KOGAS":
            return (self.kogas_preview_tree, self.kogas_item_idx_map, self.kogas_extracted_data, self.kogas_column_keys)
        elif mode == "PT":
            return (self.pt_preview_tree, self.pt_item_idx_map, self.pt_extracted_data, self.pt_column_keys)
        elif mode == "PAUT":
            return (self.paut_preview_tree, self.paut_item_idx_map, self.paut_extracted_data, self.paut_column_keys)
        return None

    def copy_cell(self, mode="PMI", target="smart"):
        """선택된 영역의 내용을 엑셀 형식(\t)으로 클립보드에 복사"""
        try:
            mode_info = self._get_mode_info(mode)
            if not mode_info: return
            tree, idx_map, data_list, col_keys = mode_info

            selected = list(tree.selection())
            if not selected: return
            selected.sort(key=lambda x: tree.index(x))

            # [REFINED] 드래그 영역 계산
            s_col = getattr(self, 'start_col', None)
            e_col = getattr(self, 'end_col', None)
            l_col = getattr(self, 'last_clicked_col', None)

            if target == "cell":
                target_col_ids = [l_col or s_col or "#4"]
            elif not s_col or not e_col or s_col == e_col:
                target_col_ids = [l_col or s_col or "#4"]
            else:
                try:
                    s_idx = int(s_col.replace('#', '')) - 1
                    e_idx = int(e_col.replace('#', '')) - 1
                    col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                    target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
                except:
                    target_col_ids = [l_col or s_col or "#4"]

            key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
            output = []

            for item_id in selected:
                view_idx = tree.index(item_id)
                if not (0 <= view_idx < len(idx_map)): continue
                actual_idx = idx_map[view_idx]
                item = data_list[actual_idx]
                
                row_vals = []
                for c_id in target_col_ids:
                    key = key_map.get(c_id)
                    if key == "selected":
                        row_vals.append("●" if item.get('selected', True) else "○")
                    elif key == "_status":
                        is_def = False
                        if mode == "PMI" and hasattr(self, 'element_filters'):
                            for f in self.element_filters:
                                fk = f['key'].get().strip()
                                if not fk: continue
                                v = self.to_float(self._get_val_ci(item, fk))
                                fm, fx = self.to_float(f['min'].get()), self.to_float(f['max'].get())
                                if (fm > 0 and v < fm) or (fx > 0 and v > fx):
                                    is_def = True; break
                        row_vals.append("⚠️" if is_def else "✅")
                    else:
                        val = item.get(key, "")
                        if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                            f_val = self.to_float(val)
                            val = f"{f_val:.2f}" if f_val > 0 else ""
                        row_vals.append(str(val))
                
                output.append("\t".join(row_vals))

            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(output))
            self.log(f"📋 {mode} 복사 완료: {len(selected)}행 x {len(target_col_ids)}열")
        except Exception as e:
            self.log(f"❌ 복사 오류: {e}")

    def paste_cell(self, mode="PMI"):
        """클립보드 내용을 선택된 영역에 스마트하게(Replicate 지원) 붙여넣기"""
        try:
            mode_info = self._get_mode_info(mode)
            if not mode_info: return
            tree, idx_map, data_list, col_keys = mode_info
            
            try: 
                clipboard_val = self.root.clipboard_get()
                if not clipboard_val: return
                # 엑셀 형식(\t, \n) 파싱
                paste_rows = [line.split("\t") for line in clipboard_val.strip().split("\n")]
            except Exception as e:
                self.log(f"⚠️ 클립보드 데이터를 가져올 수 없습니다: {e}")
                return
            
            selected = list(tree.selection())
            if not selected:
                messagebox.showwarning("붙여넣기 오류", "붙여넣을 셀을 선택해주세요.")
                return
            selected.sort(key=lambda x: tree.index(x))
            
            # [SMART AREA] 드래그된 컬럼 범위를 계산
            s_col = getattr(self, 'start_col', None)
            e_col = getattr(self, 'end_col', None)
            l_col = getattr(self, 'last_clicked_col', None)
            
            if not s_col or not e_col or s_col == e_col:
                target_col = l_col or s_col or "#4"
                try:
                    s_idx = int(target_col.replace('#', '')) - 1
                    col_start = s_idx
                    # 단일 셀 선택 시 클립보드 너비만큼 자동 확장
                    max_cols_in_paste = len(paste_rows[0]) if paste_rows else 1
                    col_end = min(len(col_keys) - 1, col_start + max_cols_in_paste - 1)
                except:
                    col_start, col_end = 0, len(col_keys) - 1
            else:
                try:
                    s_idx = int(s_col.replace('#', '')) - 1
                    e_idx = int(e_col.replace('#', '')) - 1
                    col_start, col_end = min(s_idx, e_idx), max(s_idx, e_idx)
                except:
                    col_start, col_end = 0, len(col_keys) - 1
            
            target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
            key_map = {f"#{i+1}": k for i, k in enumerate(col_keys)}
            
            # [EXCEL-LIKE REPLICATION]
            # 한 행만 선택했는데 클립보드가 여러 행이면 -> 자동 드래그 (Fill Down)
            all_children = tree.get_children('')
            if len(selected) == 1 and len(paste_rows) > 1:
                curr_idx = all_children.index(selected[0])
                for i in range(1, len(paste_rows)):
                    if curr_idx + i < len(all_children):
                        selected.append(all_children[curr_idx + i])
            
            for r_idx, item_id in enumerate(selected):
                view_idx = tree.index(item_id)
                if not (0 <= view_idx < len(idx_map)): continue
                actual_idx = idx_map[view_idx]
                
                # 행 복제(Replicate): 클립보드 행보다 선택된 행이 많으면 순환하며 붙여넣음
                source_row = paste_rows[r_idx % len(paste_rows)]
                
                for c_idx, c_id in enumerate(target_col_ids):
                    # 열 복제(Replicate): 클립보드 열보다 선택된 열이 많으면 순환하며 붙여넣음
                    cell_val = source_row[c_idx % len(source_row)].strip()
                    key = key_map.get(c_id)
                    
                    if key and key not in ["selected", "_status"]:
                        if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                            f_val = self.to_float(cell_val)
                            data_list[actual_idx][key] = f"{f_val:.2f}" if f_val > 0 else ""
                        else:
                            data_list[actual_idx][key] = cell_val
            
            self.populate_preview(data_list, switch_tab=False, mode=mode)
            self.log(f"📋 {mode} 붙여넣기 완료: {len(selected)}행 x {len(target_col_ids)}열")
        except Exception as e:
            self.log(f"❌ 붙여넣기 오류: {e}")
            messagebox.showerror("붙여넣기 실패", f"데이터를 붙여넣는 중 오류가 발생했습니다:\n{e}")

    def merge_selected_iso(self, mode="PMI"):
        """선택된 항목들의 ISO 번호를 첫 번째 항목의 것으로 통일 (병합 효과)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, _ = mode_info
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        k_iso = "ISO" if mode == "PAUT" else "Dwg"
        if actual_idx_0 is not None:
            first_iso = data_list[actual_idx_0].get(k_iso, '')
        else:
            first_iso = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data_list[actual_idx][k_iso] = first_iso
                if i > 0: data_list[actual_idx]['is_merged_iso'] = True
                else: data_list[actual_idx].pop('is_merged_iso', None)
        
        self.populate_preview(data_list, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 ISO 병합 완료: {first_iso}")

    def merge_selected_joint(self, mode="PMI"):
        """선택된 항목들의 Joint No를 첫 번째 항목의 것으로 통일 (데이터 완전 변경)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data_list, _ = mode_info
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            first_joint = data_list[actual_idx_0].get('Joint', '')
        else:
            first_joint = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data_list[actual_idx]['Joint'] = first_joint
                if i > 0: data_list[actual_idx]['is_merged_joint'] = True
                else: data_list[actual_idx].pop('is_merged_joint', None)
        
        self.populate_preview(data_list, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 Joint 병합 완료: {first_joint}")

    def group_selected_joint(self, mode="PMI"):
        """선택된 항목들을 시각적으로만 같은 Joint로 취급하여 묶음 (데이터 유지)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "그룹화할 항목을 2개 이상 선택해주세요.")
            return
            
        selected = sorted(list(selected), key=lambda x: tree.index(x))
            
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            group_val = data[actual_idx_0].get('Joint', '')
        else:
            group_val = ""
            
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['visual_group_joint'] = group_val
                if i > 0: data[actual_idx]['is_merged_joint'] = True
                else: data[actual_idx].pop('is_merged_joint', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 Joint 시각적 그룹화 완료")

    def ungroup_selected_iso(self, mode="PMI"):
        """선택된 항목들의 ISO 시각적 그룹화 연결 끊기"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_iso', None)
                data[actual_idx].pop('is_merged_iso', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 ISO 병합 해제 완료")

    def ungroup_selected_iso(self, mode="PMI"):
        """선택된 항목들의 시각적 ISO 병합 해제"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_iso', None)
                data[actual_idx].pop('is_merged_iso', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 ISO 병합 해제 완료")

    def ungroup_selected_joint(self, mode="PMI"):
        """선택된 항목들의 Joint 시각적 그룹화 연결 끊기"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
            
        for item_id in selected:
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx].pop('visual_group_joint', None)
                data[actual_idx].pop('is_merged_joint', None)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"✂️ {mode} {len(selected)}개 항목 Joint 병합 해제 완료")

    def select_all(self, mode="PMI"):
        """모든 항목 체크"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        for item in data:
            if (item.get('date_filtered', True) and (mode not in ('MT', 'RT', 'KOGAS') or item.get('drawing_filtered', True))
                    and (mode != "PAUT" or (
                        item.get('line_filtered', True)
                        and item.get('size_filtered', True)))):
                item['selected'] = True
        self.populate_preview(data, switch_tab=False, mode=mode)



    def deselect_all(self, mode="PMI"):
        """모든 항목 체크 해제"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        for item in data:
            if (item.get('date_filtered', True) and (mode not in ('MT', 'RT', 'KOGAS') or item.get('drawing_filtered', True))
                    and (mode != "PAUT" or (
                        item.get('line_filtered', True)
                        and item.get('size_filtered', True)))):
                item['selected'] = False
        self.populate_preview(data, switch_tab=False, mode=mode)



    def select_deficient_items(self):
        """PMI: 함량 미달인 항목만 자동으로 선택(체크)"""
        if not self.extracted_data or not hasattr(self, 'element_filters'): return
        
        count = 0
        for item in self.extracted_data:
            is_deficient = False
            for f_item in self.element_filters:
                f_key = f_item['key'].get().strip()
                if not f_key: continue
                f_min = round(self.to_float(f_item['min'].get()), 2)
                f_max = round(self.to_float(f_item['max'].get()), 2)
                
                val = round(self.to_float(self._get_val_ci(item, f_key)), 2)
                if (f_min > 0 and val < f_min) or (f_max > 0 and val > f_max):
                    is_deficient = True
                    break
            
            item['selected'] = is_deficient
            if is_deficient: count += 1
            
        self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        self.log(f"✅ 미달 항목 {count}건을 선택했습니다.")

    def reset_pmi_filters(self, refresh=True, clear_search=True):
        """PMI 검색 및 함량 필터 초기화"""
        if clear_search:
            self.pmi_search_loc.set("")
            # 사이드바 지점 필터도 초기화
            if hasattr(self, 'pmi_loc_listbox'):
                for i in range(self.pmi_loc_listbox.size()):
                    val = self.pmi_loc_listbox.get(i)
                    self.pmi_loc_listbox.delete(i)
                    self.pmi_loc_listbox.insert(i, val.replace("[v] ", "[ ] "))
        
        self.pmi_show_deficiency_only.set(False)
        if hasattr(self, 'element_filters'):
            for f in self.element_filters:
                f['min'].set("")
                f['max'].set("")
        if refresh:
            self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        self.log("🧹 PMI 필터가 초기화되었습니다.")

    def update_pmi_loc_listbox(self):
        """현재 데이터의 Test Location 목록을 추출하여 드롭다운 갱신 (날짜 필터 적용된 항목만)"""
        if not hasattr(self, 'pmi_loc_listbox') or not self.extracted_data: return
        
        # 기존 선택 세트 백업
        selected_locs = set()
        for i in range(self.pmi_loc_listbox.size()):
            val = self.pmi_loc_listbox.get(i)
            if val.startswith("[v] "):
                selected_locs.add(val.replace("[v] ", ""))
        
        # 새 목록 추출
        current_locs = sorted(list(set(str(item.get('Loc', '')).strip() for item in self.extracted_data if item.get('Loc') and item.get('date_filtered', True))))
        
        self.pmi_loc_listbox.delete(0, tk.END)
        for loc in current_locs:
            prefix = "[v] " if loc in selected_locs else "[ ] "
            self.pmi_loc_listbox.insert(tk.END, prefix + loc)

    def add_item(self, mode="PMI"):
        """새 데이터 행 추가"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        
        new_item = {'selected': True, 'date_filtered': True, 'order_index': len(data)}
        # 기본 키 초기화
        keys = self.kogas_column_keys if mode == "KOGAS" else self.mt_column_keys if mode == "MT" else self.rt_column_keys if mode == "RT" else (self.pt_column_keys if mode == "PT" else (self.paut_column_keys if mode == "PAUT" else self.column_keys))
        for k in keys:
            if k not in new_item: new_item[k] = ""
            
        data.append(new_item)
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"➕ {mode} 새 데이터 항목 추가 완료")



    def move_item(self, direction, mode="PMI"):
        """선택된 아이템의 순서를 위/아래로 이동 (필터 상태 대응)"""
        # direction이 문자열인 경우 처리 ("up" -> -1, "down" -> 1)
        if isinstance(direction, str):
            direction = -1 if direction.lower() == "up" else 1
        
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, data, _ = mode_info
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        view_indices = sorted([tree.index(sid) for sid in selected_ids])
        
        if direction == -1: # 위로
            if view_indices[0] == 0: return
        else: # 아래로
            if view_indices[-1] == len(tree.get_children()) - 1: return
            
        # 데이터 리스트(data)에서 위치 교환
        # Treeview 인덱스를 item_idx_map을 통해 실제 데이터 인덱스로 변환
        actual_indices = [idx_map[vi] for vi in view_indices]
        
        # 이동 방향에 따라 처리 순서 결정
        if direction == -1: # 위로
            # 가장 위 아이템의 위쪽 아이템 찾기 (Treeview상에서)
            target_view_idx = view_indices[0] - 1
            target_actual_idx = idx_map[target_view_idx]
            
            # 선택된 아이템들을 한 칸씩 위로 (target_actual_idx 앞으로)
            items_to_move = [data.pop(ai) for ai in reversed(actual_indices)]
            # target_actual_idx가 팝업으로 인해 밀렸을 수 있으므로 다시 계산
            # 하지만 팝업 시 뒤쪽부터 뺐으므로 target_actual_idx 위치는 유지됨
            for item in items_to_move:
                data.insert(target_actual_idx, item)
        else: # 아래로
            # 가장 아래 아이템의 아래쪽 아이템 찾기
            target_view_idx = view_indices[-1] + 1
            target_actual_idx = idx_map[target_view_idx]
            
            # 선택된 아이템들을 한 칸씩 아래로 (target_actual_idx 뒤로)
            items_to_move = [data.pop(ai) for ai in actual_indices]
            # 팝업으로 인해 인덱스가 당겨졌으므로 target_actual_idx에서 len(items_to_move) 만큼 보정하거나 단순히 insert
            insert_pos = target_actual_idx - len(items_to_move) + 1
            for item in reversed(items_to_move):
                data.insert(insert_pos, item)
                
        self.populate_preview(data, switch_tab=False, mode=mode)
        
        # 선택 상태 복원 (Treeview 아이템이 다시 생성되므로 인덱스로 재선택)
        new_children = tree.get_children()
        for vi in view_indices:
            new_pos = max(0, min(len(new_children)-1, vi + direction))
            tree.selection_add(new_children[new_pos])

    def delete_item(self, mode="PMI"):
        """선택된 아이템 삭제"""
        if mode == "RT":
            tree, idx_map, data = self.rt_preview_tree, self.rt_item_idx_map, self.rt_extracted_data
        elif mode == "KOGAS":
            tree, idx_map, data = self.kogas_preview_tree, self.kogas_item_idx_map, self.kogas_extracted_data
        elif mode == "PT":
            tree, idx_map, data = self.pt_preview_tree, self.pt_item_idx_map, self.pt_extracted_data
        elif mode == "MT":
            tree, idx_map, data = self.mt_preview_tree, self.mt_item_idx_map, self.mt_extracted_data
        elif mode == "PAUT":
            tree, idx_map, data = self.paut_preview_tree, self.paut_item_idx_map, self.paut_extracted_data
        else:
            tree, idx_map, data = self.preview_tree, self.item_idx_map, self.extracted_data
        
        selected_ids = tree.selection()
        if not selected_ids: return
        
        if not messagebox.askyesno("삭제 확인", f"선택한 {len(selected_ids)}개 항목을 정말 삭제하시겠습니까?"):
            return
            
        view_indices = sorted([tree.index(sid) for sid in selected_ids], reverse=True)
        for vi in view_indices:
            actual_idx = idx_map[vi]
            data.pop(actual_idx)
            
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.update_date_listbox(mode)
        self.log(f"🗑️ {mode} {len(selected_ids)}개 항목 삭제 완료")


    def clear_all(self, mode="PMI"):
        """모든 데이터 초기화"""
        if messagebox.askyesno("확인", f"모든 {mode} 데이터를 초기화하시겠습니까?"):
            if mode == "RT": self.rt_extracted_data = []
            elif mode == "KOGAS": self.kogas_extracted_data = []
            elif mode == "PT": self.pt_extracted_data = []
            elif mode == "MT": self.mt_extracted_data = []
            elif mode == "PAUT": self.paut_extracted_data = []
            else: self.extracted_data = []
            self.update_date_listbox(mode)
            self.populate_preview([], mode=mode)
            self.log(f"🧹 모든 {mode} 데이터 초기화 완료")


    def export_to_excel(self, mode="PMI"):
        """현재 미리보기 목록(필터링/선택 반영)을 엑셀 파일로 내보냄 (서식 및 병합 시인성 유지)"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        _, _, data, col_keys = mode_info
        
        if not data:
            messagebox.showwarning("알림", f"{mode} 내보낼 데이터가 없습니다.")
            return

        # 1. 엑셀에 저장할 대상 데이터 필터링
        filter_enabled = self.show_selected_only.get()
        final_list = [
            d for d in data
            if d.get('date_filtered', True)
            and (mode not in ('MT', 'RT', 'KOGAS') or d.get('drawing_filtered', True))
            and (mode != "PAUT" or (
                d.get('line_filtered', True) and d.get('size_filtered', True)))
            and (not filter_enabled or d.get('selected', True))
        ]
        
        if not final_list:
            messagebox.showinfo("알림", f"{mode} 현재 조건에 맞는 데이터가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"{mode}_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            title=f"{mode} 서식 있는 엑셀 파일로 내보내기"
        )
        if not file_path: return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
            import openpyxl.utils

            wb = Workbook()
            ws = wb.active
            ws.title = f"{mode}_Preview"

            # [REFINED] Export with consistent headers from mode_info
            if mode == "PMI":
                # PMI는 특수하게 ST, V 헤더 추가
                headers = ["ST", "V", "No", "Date", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"]
            else:
                # RT/PT/PAUT는 Treeview 컬럼 헤더 그대로 사용
                tree = self._get_mode_info(mode)[0]
                headers = [tree.heading(c)['text'] for c in tree['columns']]
            
            export_rows = []
            for item in final_list:
                row = {}
                if mode == "PMI":
                    # Status Icon
                    is_def = False
                    if hasattr(self, 'element_filters'):
                        for f in self.element_filters:
                            fk = f['key'].get().strip()
                            if not fk: continue
                            v = self.to_float(self._get_val_ci(item, fk))
                            fm, fx = self.to_float(f['min'].get()), self.to_float(f['max'].get())
                            if (fm > 0 and v < fm) or (fx > 0 and v > fx):
                                is_def = True; break
                    
                    row["ST"] = "⚠️" if is_def else "✅"
                    row["V"] = "●" if item.get('selected', True) else "○"
                    row["No"] = item.get('No', '')
                    row["Date"] = item.get('Date', '')
                    row["ISO/DWG"] = item.get('Dwg', '')
                    row["Joint No"] = item.get('Joint', '')
                    row["Test Location"] = item.get('Loc', '')
                    row["Ni"] = round(self.to_float(item.get('Ni', 0)), 2)
                    row["Cr"] = round(self.to_float(item.get('Cr', 0)), 2)
                    row["Mo"] = round(self.to_float(item.get('Mo', 0)), 2)
                    row["Grade"] = item.get('Grade', '')
                else:
                    # Generic mapping for other modes
                    # col_keys: ("_status", "selected", "No", ...)
                    for i, h in enumerate(headers):
                        k = col_keys[i] if i < len(col_keys) else None
                        if k == "_status":
                            row[h] = "✅" 
                        elif k == "selected":
                            row[h] = "●" if item.get('selected', True) else "○"
                        elif k:
                            row[h] = item.get(k, "")
                        else:
                            row[h] = ""
                export_rows.append(row)

            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')

            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row_idx, row_data in enumerate(export_rows, 2):
                for col_idx, key in enumerate(headers, 1):
                    val = row_data.get(key, "")
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border

            if mode == "PMI":
                column_widths = {1: 5, 2: 5, 3: 8, 4: 12, 5: 25, 6: 15, 7: 15, 8: 8, 9: 8, 10: 8, 11: 15}
            else:
                # RT Column Widths
                column_widths = {1: 5, 2: 8, 3: 12, 4: 25, 5: 15, 6: 10}
                for i in range(7, 22): column_widths[i] = 4 # Defects
                column_widths[22] = 8 # Result
                column_widths[23] = 12 # Welder
                column_widths[24] = 20 # Remarks
                
            for col_idx, width in column_widths.items():
                if col_idx <= len(headers):
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = width

            wb.save(file_path)
            self.log(f"📊 {mode} 서식 포함 엑셀 내보내기 완료: {os.path.basename(file_path)}")
            messagebox.showinfo("완료", f"화면 형식과 동일하게 엑셀 파일로 저장했습니다.\n(총 {len(export_rows)} 건)")
        except Exception as e:
            self.log(f"❌ {mode} 엑셀 내보내기 오류: {e}")
            messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {e}")
        finally:
            if 'wb' in locals():
                try: wb.close()
                except: pass

    def _browse_dir(self, var):
        current = var.get()
        init_dir = current if current and os.path.exists(current) else RESOURCE_DIR
        path = filedialog.askdirectory(initialdir=init_dir)
        if path: var.set(path)

    def _update_file_info(self, mode, path):
        """Extracts and formats file metadata for the UI header."""
        if not path or not os.path.exists(path):
            self.file_info_vars[mode].set("📄 파일을 선택해주세요.")
            return

        try:
            fname = os.path.basename(path)
            fsize = os.path.getsize(path) / 1024 # KB
            mtime = os.path.getmtime(path)
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            
            size_str = f"{fsize:.1f}KB" if fsize < 1024 else f"{fsize/1024:.2f}MB"
            self.file_info_vars[mode].set(f"📄 파일: {fname}  |  📏 크기: {size_str}  |  📅 수정일: {dt_str}")
        except Exception as e:
            self.file_info_vars[mode].set(f"⚠️ 파일 정보 읽기 오류: {str(e)}")

    def _sync_all_file_infos(self):
        """Updates all mode headers from current config paths."""
        modes = [
            ('PMI', self.target_file_path.get()),
            ('RT', self.rt_target_file_path.get()),
            ('KOGAS', self.kogas_target_file_path.get()),
            ('PT', self.pt_target_file_path.get()),
            ('MT', self.mt_target_file_path.get()),
            ('PAUT', self.paut_target_file_path.get())
        ]
        for mode, path in modes:
            self._update_file_info(mode, path)

    def _browse_file(self, var, types):
        current = var.get()
        first_current = current.split('|')[0] if current else ""
        init_dir = os.path.dirname(first_current) if first_current and os.path.exists(first_current) else BASE_DIR
        paths = filedialog.askopenfilenames(initialdir=init_dir, filetypes=types)
        if paths: 
            var.set("|".join(paths))
            # Find which mode this belongs to and update info
            for mode, m_var in [('PMI', self.target_file_path), ('RT', self.rt_target_file_path), 
                               ('KOGAS', self.kogas_target_file_path), ('PT', self.pt_target_file_path), ('MT', self.mt_target_file_path), 
                               ('PAUT', self.paut_target_file_path)]:
                if var == m_var:
                    try: first_path = paths[0]
                    except: first_path = paths
                    self._update_file_info(mode, first_path)
                    break 

    def to_float(self, val):
        try:
            s = str(val).replace('%', '').strip()
            if not s or s.lower() == 'nan': return 0.0
            return float(s)
        except: return 0.0

    def _get_val_ci(self, item, key):
        """대소문자 구분 없이 딕셔너리에서 값 찾기 (Cr vs CR 등)"""
        if not key or not isinstance(item, dict): return None
        if key in item: return item[key]
        k_lower = key.lower()
        for k in item.keys():
            if str(k).lower() == k_lower:
                return item[k]
        return None

    # --- PMI 재질별 표준 화학성분 스펙 및 판정 DB ---
    PMI_MATERIAL_SPECS = {
        "A312-TP304": {
            "name": "A312-TP304",
            "title": "ASTM A312 TP304",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A312-TP304: Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%",
                "ER308L: Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304: Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (8.0, 11.0), "Mo": (0.0, 0.5)}
        },
        "A312-TP304L": {
            "name": "A312-TP304L",
            "title": "ASTM A312 TP304L",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A312-TP304L: Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "ER308L: Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304L: Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (9.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "A312-TP316": {
            "name": "A312-TP316",
            "title": "ASTM A312 TP316",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A312-TP316: Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316: Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A182-F316: Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A312-TP316L": {
            "name": "A312-TP316L",
            "title": "ASTM A312 TP316L",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A312- Ni:10.0 ~ 14.0%, Cr:16.0 ~ 18.0%, Mo:2.0 ~ 3.0%",
                "ER316L- Ni:11.0 ~ 14.0%, Cr:18.0 ~ 20.0%, Mo:2.0 ~ 3.0%",
                "A182-Ni:10.0 ~ 15.0%, Cr:16.0 ~ 18.0%, Mo:2.0 ~ 3.0%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A312-TP321": {
            "name": "A312-TP321",
            "title": "ASTM A312 TP321",
            "elements": ["Ni", "Cr", "Ti", "-", "-"],
            "specs": [
                "A312-TP321: Ni:9.00 ~ 12.00%, Cr:17.00 ~ 19.00%, Ti:0.35 ~ 0.70%",
                "ER321: Ni:9.00 ~ 11.00%, Cr:18.50 ~ 20.50%, Ti:0.30 ~ 0.75%",
                "A182-F321: Ni:9.00 ~ 12.00%, Cr:17.00 ~ 19.00%, Ti:0.35 ~ 0.70%"
            ],
            "limits": {"Ni": (9.0, 12.0), "Cr": (17.0, 19.0), "Ti": (0.35, 0.70)}
        },
        "A312-TP347": {
            "name": "A312-TP347",
            "title": "ASTM A312 TP347",
            "elements": ["Ni", "Cr", "Nb", "-", "-"],
            "specs": [
                "A312-TP347: Ni:9.00 ~ 13.00%, Cr:17.00 ~ 19.00%, Nb:0.80 ~ 1.00%",
                "ER347: Ni:9.00 ~ 11.00%, Cr:19.00 ~ 21.50%, Nb:0.75 ~ 1.00%",
                "A182-F347: Ni:9.00 ~ 13.00%, Cr:17.00 ~ 19.00%, Nb:0.80 ~ 1.00%"
            ],
            "limits": {"Ni": (9.0, 13.0), "Cr": (17.0, 19.0), "Nb": (0.80, 1.00)}
        },
        "DUPLEX (S31803)": {
            "name": "DUPLEX (S31803)",
            "title": "ASTM A790 S31803",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "S31803: Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%",
                "ER2209: Cr:21.50 ~ 23.50%, Ni:7.50 ~ 9.50%, Mo:2.50 ~ 3.50%",
                "A182-F51: Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%"
            ],
            "limits": {"Cr": (21.0, 23.0), "Ni": (4.5, 6.5), "Mo": (2.5, 3.5)}
        },
        "INCONEL 625": {
            "name": "INCONEL 625",
            "title": "ASTM B444 N06625",
            "elements": ["Ni", "Cr", "Mo", "Nb", "-"],
            "specs": [
                "Inconel 625: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%",
                "ERNiCrMo-3: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%",
                "B564-N06625: Ni:min 58.0%, Cr:20.0 ~ 23.0%, Mo:8.0 ~ 10.0%, Nb:3.15 ~ 4.15%"
            ],
            "limits": {"Ni": (58.0, 100.0), "Cr": (20.0, 23.0), "Mo": (8.0, 10.0), "Nb": (3.15, 4.15)}
        },
        # --- 구조용 강재 (COLUMN, GIRDER, BEAM / H-BEAM) ---
        "ASTM A992": {
            "name": "ASTM A992",
            "title": "ASTM A992 (BEAM / COLUMN)",
            "elements": ["Mn", "Si", "V", "Cu", "Cr"],
            "specs": [
                "ASTM A992: Mn:0.50 ~ 1.60%, Si:0.15 ~ 0.40%, V:≤0.15%",
                "WELD (E7018): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "A992 (Residual): Cu:≤0.60%, Ni:≤0.45%, Cr:≤0.35%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.15, 0.40), "V": (0.00, 0.15)}
        },
        "ASTM A572 Gr.50": {
            "name": "ASTM A572 Gr.50",
            "title": "ASTM A572 Gr.50 (GIRDER / PLATE)",
            "elements": ["Mn", "Si", "V", "Nb", "Cr"],
            "specs": [
                "A572 Gr.50: Mn:≤1.35%, Si:0.15 ~ 0.40%, V:0.01 ~ 0.15%",
                "WELD (E7018): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "A572 Gr.50: Nb(Cb):0.005 ~ 0.05%, Cu:≤0.50%"
            ],
            "limits": {"Mn": (0.50, 1.35), "Si": (0.15, 0.40), "V": (0.01, 0.15)}
        },
        "SM355 (SM490)": {
            "name": "SM355 (SM490)",
            "title": "KS D 3515 SM355 (용접구조용)",
            "elements": ["Mn", "Si", "P", "S", "Cu"],
            "specs": [
                "SM355: Mn:≤1.60%, Si:≤0.55%, P:≤0.030%",
                "WELD (E7018/YGW11): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "SM355: S:≤0.030%, Ceq:≤0.44%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.10, 0.55)}
        },
        "SN355": {
            "name": "SN355",
            "title": "KS D 3866 SN355 (건축구조용 H형강)",
            "elements": ["Mn", "Si", "P", "S", "Cr"],
            "specs": [
                "SN355: Mn:≤1.60%, Si:≤0.55%, P:≤0.030%",
                "WELD (E7018/YGW11): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "SN355: S:≤0.015%, Ceq:≤0.45%"
            ],
            "limits": {"Mn": (0.50, 1.60), "Si": (0.10, 0.55)}
        },
        "ASTM A36": {
            "name": "ASTM A36",
            "title": "ASTM A36 (탄소 구조강)",
            "elements": ["Mn", "Si", "Cu", "P", "S"],
            "specs": [
                "ASTM A36: Mn:0.60 ~ 1.20%, Si:≤0.40%",
                "WELD (E6013/E7018): Mn:0.40 ~ 1.25%, Si:≤0.60%",
                "ASTM A36: P:≤0.030%, S:≤0.030%"
            ],
            "limits": {"Mn": (0.40, 1.20), "Si": (0.10, 0.40)}
        },
        "SS275 (SS400)": {
            "name": "SS275 (SS400)",
            "title": "KS D 3503 SS275 (일반구조용)",
            "elements": ["Mn", "Si", "P", "S", "-"],
            "specs": [
                "SS275: Mn:≤1.40%, Si:≤0.45%",
                "WELD (E4313/E7016): Mn:0.40 ~ 1.20%, Si:≤0.50%",
                "SS275: P:≤0.040%, S:≤0.040%"
            ],
            "limits": {"Mn": (0.30, 1.40), "Si": (0.05, 0.45)}
        },
        # --- Project PROVIDENCE 현장 공식 성분표 규격 (SA-351 / SA-403 / SFA-5.9) ---
        "A351-CF8M": {
            "name": "A351-CF8M",
            "title": "ASTM A351 CF8M (316 VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF8M (Valve): Cr:18.00 ~ 21.00%, Ni:9.00 ~ 12.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (16.0, 21.0), "Ni": (9.0, 15.0), "Mo": (2.0, 3.0)}
        },
        "A351-CF3M": {
            "name": "A351-CF3M",
            "title": "ASTM A351 CF3M (316L VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF3M (Valve): Cr:17.00 ~ 21.00%, Ni:9.00 ~ 13.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (16.0, 21.0), "Ni": (9.0, 15.0), "Mo": (2.0, 3.0)}
        },
        "A351-CF8": {
            "name": "A351-CF8",
            "title": "ASTM A351 CF8 (304 VALVE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A351-CF8 (Valve): Cr:18.00 ~ 21.00%, Ni:8.00 ~ 11.00%, Mo:≤0.50%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A182-F304 (Flange): Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%"
            ],
            "limits": {"Cr": (18.0, 22.0), "Ni": (8.0, 11.0), "Mo": (0.0, 0.75)}
        },
        "A403-WP316L": {
            "name": "A403-WP316L",
            "title": "ASTM A403 WP316L (FITTINGS)",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A403-WP316L (Fitting): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 14.0), "Cr": (16.0, 20.0), "Mo": (2.0, 3.0)}
        },
        "A403-WP304L": {
            "name": "A403-WP304L",
            "title": "ASTM A403 WP304L (FITTINGS)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A403-WP304L (Fitting): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (9.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "A182-F316L": {
            "name": "A182-F316L",
            "title": "ASTM A182 F316L (FLANGE)",
            "elements": ["Ni", "Cr", "Mo", "-", "-"],
            "specs": [
                "A182-F316L (Flange): Ni:10.00 ~ 15.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%",
                "ER316L (Weld): Ni:11.00 ~ 14.00%, Cr:18.00 ~ 20.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Ni:10.00 ~ 14.00%, Cr:16.00 ~ 18.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Ni": (10.0, 15.0), "Cr": (16.0, 18.0), "Mo": (2.0, 3.0)}
        },
        "A182-F304L": {
            "name": "A182-F304L",
            "title": "ASTM A182 F304L (FLANGE)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "A182-F304L (Flange): Cr:18.00 ~ 20.00%, Ni:8.00 ~ 13.00%",
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (8.0, 13.0), "Mo": (0.0, 0.5)}
        },
        "ER309L (이재질)": {
            "name": "ER309L (이재질)",
            "title": "AWS A5.9 ER309L (DISSIMILAR WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER309L (Weld): Cr:23.00 ~ 25.00%, Ni:12.00 ~ 14.00%, Mo:≤0.75%",
                "A312-TP316L / TP304L: Cr:16.00 ~ 20.00%, Ni:10.00 ~ 14.00%",
                "Carbon Steel (Base): Fe:Balance, Mn:≤1.60%"
            ],
            "limits": {"Cr": (16.0, 25.0), "Ni": (8.0, 14.0)}
        },
        # --- 용접봉 및 용접재 단독 규격 (AWS A5.9 / A5.18 / ASME SFA-5.9) ---
        "ER316L": {
            "name": "ER316L",
            "title": "AWS A5.9 ER316L (WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER316L (Weld): Cr:18.00 ~ 20.00%, Ni:11.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A312-TP316L (Pipe): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 14.00%, Mo:2.00 ~ 3.00%",
                "A182-F316L (Flange): Cr:16.00 ~ 18.00%, Ni:10.00 ~ 15.00%, Mo:2.00 ~ 3.00%"
            ],
            "limits": {"Cr": (18.0, 20.0), "Ni": (11.0, 14.0), "Mo": (2.0, 3.0)}
        },
        "ER308L": {
            "name": "ER308L",
            "title": "AWS A5.9 ER308L (WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER308L (Weld): Cr:19.50 ~ 22.00%, Ni:9.00 ~ 11.00%, Mo:≤0.75%",
                "A312-TP304L (Pipe): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%",
                "A182-F304L (Flange): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%"
            ],
            "limits": {"Cr": (19.5, 22.0), "Ni": (9.0, 11.0), "Mo": (0.0, 0.75)}
        },
        "ER2209": {
            "name": "ER2209",
            "title": "AWS A5.9 ER2209 (DUPLEX WELD)",
            "elements": ["Cr", "Ni", "Mo", "-", "-"],
            "specs": [
                "ER2209 (Duplex Weld): Cr:21.50 ~ 23.50%, Ni:7.50 ~ 9.50%, Mo:2.50 ~ 3.50%",
                "S31803 / S32205 (Pipe): Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%",
                "A182-F51 (Flange): Cr:21.00 ~ 23.00%, Ni:4.50 ~ 6.50%, Mo:2.50 ~ 3.50%"
            ],
            "limits": {"Cr": (21.5, 23.5), "Ni": (7.5, 9.5), "Mo": (2.5, 3.5)}
        },
        "ER70S-6": {
            "name": "ER70S-6",
            "title": "AWS A5.18 ER70S-6 / E7018 (WELD)",
            "elements": ["Mn", "Si", "P", "S", "Cu"],
            "specs": [
                "ER70S-6 (Weld Wire): Mn:1.40 ~ 1.85%, Si:0.80 ~ 1.15%",
                "E7018 (SMAW): Mn:0.80 ~ 1.60%, Si:0.20 ~ 0.75%",
                "ASTM A992 / SM355: Mn:0.50 ~ 1.60%, Si:0.15 ~ 0.55%"
            ],
            "limits": {"Mn": (0.80, 1.85), "Si": (0.20, 1.15)}
        }
    }

    PMI_MATERIAL_ALIASES = {
        "TP304": "A312-TP304", "SUS304": "A312-TP304", "SS304": "A312-TP304", "304": "A312-TP304",
        "TP304L": "A312-TP304L", "SUS304L": "A312-TP304L", "SS304L": "A312-TP304L", "304L": "A312-TP304L",
        "TP316": "A312-TP316", "SUS316": "A312-TP316", "SS316": "A312-TP316", "316": "A312-TP316",
        "TP316L": "A312-TP316L", "SUS316L": "A312-TP316L", "SS316L": "A312-TP316L", "316L": "A312-TP316L",
        "TP321": "A312-TP321", "SUS321": "A312-TP321", "SS321": "A312-TP321", "321": "A312-TP321",
        "TP347": "A312-TP347", "SUS347": "A312-TP347", "SS347": "A312-TP347", "347": "A312-TP347",
        "DUPLEX": "DUPLEX (S31803)", "S31803": "DUPLEX (S31803)", "S32205": "DUPLEX (S31803)",
        "INCONEL": "INCONEL 625", "INCONEL625": "INCONEL 625",
        "A992": "ASTM A992", "WBEAM": "ASTM A992", "HBEAM": "ASTM A992", "BEAM": "ASTM A992", "COLUMN": "ASTM A992", "GIRDER": "ASTM A992",
        "A572": "ASTM A572 Gr.50", "A572GR50": "ASTM A572 Gr.50", "GR50": "ASTM A572 Gr.50",
        "SM355": "SM355 (SM490)", "SM490": "SM355 (SM490)", "SM355A": "SM355 (SM490)", "SM355B": "SM355 (SM490)",
        "SN355": "SN355", "SN490": "SN355", "SHN355": "SN355",
        "A36": "ASTM A36",
        "SS275": "SS275 (SS400)", "SS400": "SS275 (SS400)",
        "CF8M": "A351-CF8M", "CF3M": "A351-CF3M", "CF8": "A351-CF8", "CF3": "A351-CF8", "VALVE": "A351-CF8M", "351": "A351-CF8M",
        "WP316L": "A403-WP316L", "WP316": "A403-WP316L", "WP304L": "A403-WP304L", "WP304": "A403-WP304L", "FITTING": "A403-WP316L",
        "ELBOW": "A403-WP316L", "TEE": "A403-WP316L", "REDUCER": "A403-WP316L",
        "엘보": "A403-WP316L", "티": "A403-WP316L", "레듀샤": "A403-WP316L",
        "ER309L": "ER309L (이재질)", "309L": "ER309L (이재질)", "309": "ER309L (이재질)", "이재질": "ER309L (이재질)",
        "ER316L": "ER316L", "E316L": "ER316L", "316LWELD": "ER316L", "ER316": "ER316L",
        "ER308L": "ER308L", "E308L": "ER308L", "308L": "ER308L", "ER308": "ER308L",
        "ER2209": "ER2209", "2209": "ER2209", "E2209": "ER2209",
        "ER70S6": "ER70S-6", "ER70S-6": "ER70S-6", "YGW11": "ER70S-6", "YGW12": "ER70S-6", "E7018": "ER70S-6", "WELD": "ER316L",
        "F316L": "A182-F316L", "F316": "A182-F316L", "F304L": "A182-F304L", "F304": "A182-F304L", "FLANGE": "A182-F316L", "A182": "A182-F316L"
    }

    # 기준 1 / 2 / 3 전용 빠른 선택 목록 (선택한 자재에 맞춰 모든 기준에서 자유롭게 선택 가능)
    PMI_COMMON_SPEC_OPTIONS = [
        "(공란 / 비움)",
        # 배관 (PIPE)
        "A312-TP316L", "A312-TP304L", "A312-TP316", "A312-TP304",
        "A312-TP321", "A312-TP347", "DUPLEX (S31803)", "INCONEL 625",
        # 피팅 (ELBOW / TEE / REDUCER)
        "A403-WP316L (피팅/엘보)", "A403-WP304L (피팅/엘보)",
        # 플랜지 (FLANGE)
        "A182-F316L (플랜지)", "A182-F304L (플랜지)",
        # 밸브 (VALVE)
        "A351-CF8M (주조 밸브)", "A351-CF3M (밸브)", "A351-CF8 (304 밸브)",
        # 용접봉 (WELD)
        "ER316L (용접봉)", "ER308L (용접봉)", "ER309L (이재질)", "ER2209", "ER70S-6", "ERNiCrMo-3",
        # 구조용강 (BEAM)
        "ASTM A992", "ASTM A572 Gr.50", "SM355 (SM490)", "SN355", "ASTM A36", "SS275 (SS400)"
    ]
    PMI_SPEC1_OPTIONS = PMI_COMMON_SPEC_OPTIONS
    PMI_SPEC2_OPTIONS = PMI_COMMON_SPEC_OPTIONS
    PMI_SPEC3_OPTIONS = PMI_COMMON_SPEC_OPTIONS

    # --- 품명(형태)별 추천 재질 및 기본 프리셋 매핑 ---
    PMI_ITEM_PRESETS = {
        "PIPE": {
            "default_material": "A312-TP316L",
            "materials": ["A312-TP316L", "A312-TP304L", "A312-TP316", "A312-TP304", "A312-TP321", "A312-TP347", "DUPLEX (S31803)", "INCONEL 625"]
        },
        "FLANGE": {
            "default_material": "A182-F316L",
            "materials": ["A182-F316L", "A182-F304L"]
        },
        "VALVE": {
            "default_material": "A351-CF8M",
            "materials": ["A351-CF8M", "A351-CF3M", "A351-CF8"]
        },
        "ELBOW": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "TEE": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "REDUCER": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "FITTING": {
            "default_material": "A403-WP316L",
            "materials": ["A403-WP316L", "A403-WP304L"]
        },
        "WELD": {
            "default_material": "ER316L",
            "materials": ["ER316L", "ER308L", "ER309L (이재질)", "ER2209", "ER70S-6"]
        },
        "BEAM/COLUMN": {
            "default_material": "ASTM A992",
            "materials": ["ASTM A992", "ASTM A572 Gr.50", "SM355 (SM490)", "SN355", "ASTM A36", "SS275 (SS400)"]
        }
    }

    # --- Excel Helper Logic ---

if __name__ == "__main__":
    root = tk.Tk()
    app = PMIReportApp(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        # Ctrl+C from a terminal is a normal shutdown request, not an app error.
        try:
            app.on_closing()
        except tk.TclError:
            # The window may already have been destroyed while handling shutdown.
            pass
