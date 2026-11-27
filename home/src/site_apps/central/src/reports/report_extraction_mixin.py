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



class ReportExtractionMixin:
    def _lookup_pmi_material_spec(self, mat_str):
        if not mat_str: return None
        raw = str(mat_str).strip().upper().replace(" ", "").replace("-", "").replace("_", "")
        # 1. Exact match with primary spec keys
        for k, v in self.PMI_MATERIAL_SPECS.items():
            k_clean = k.upper().replace(" ", "").replace("-", "").replace("_", "")
            if raw == k_clean:
                return v
        # 2. Exact match with alias keys
        for alias, target in self.PMI_MATERIAL_ALIASES.items():
            a_clean = alias.upper().replace(" ", "").replace("-", "").replace("_", "")
            if raw == a_clean:
                return self.PMI_MATERIAL_SPECS.get(target)
        # 3. Partial match (longest alias first to prevent short substring false-positive like '316' inside 'WP316L')
        sorted_aliases = sorted(self.PMI_MATERIAL_ALIASES.items(), key=lambda x: len(x[0]), reverse=True)
        for alias, target in sorted_aliases:
            a_clean = alias.upper().replace(" ", "").replace("-", "").replace("_", "")
            if len(a_clean) >= 3 and a_clean in raw:
                return self.PMI_MATERIAL_SPECS.get(target)
        # 4. Partial match with primary spec keys
        sorted_specs = sorted(self.PMI_MATERIAL_SPECS.items(), key=lambda x: len(x[0]), reverse=True)
        for k, v in sorted_specs:
            k_clean = k.upper().replace(" ", "").replace("-", "").replace("_", "")
            if len(k_clean) >= 4 and (k_clean in raw or raw in k_clean):
                return v
        return None

    def _determine_pmi_target_slot(self, spec):
        """새로 선택된 재질/스펙이 들어갈 최적의 기준 슬롯(1, 2, 3)을 순차적으로 결정"""
        m_name = spec["name"]
        is_weld = any(k in m_name for k in ["ER316", "ER308", "ER309", "ER2209", "ER70S", "ERNICRMO", "E7018", "WELD", "용접"])
        
        # 현재 슬롯별 상태 확인 (공란 여부)
        s1 = self.pmi_chem_specs[0].get().strip() if len(self.pmi_chem_specs) > 0 else ""
        s2 = self.pmi_chem_specs[1].get().strip() if len(self.pmi_chem_specs) > 1 else ""
        s3 = self.pmi_chem_specs[2].get().strip() if len(self.pmi_chem_specs) > 2 else ""

        is_s1_empty = (not s1 or s1 == "-")
        is_s2_empty = (not s2 or s2 == "-")
        is_s3_empty = (not s3 or s3 == "-")

        # 1. 용접봉/용접재인 경우: 기준 2(용접부) 우선 배정
        if is_weld:
            if is_s2_empty or (m_name in s2):
                return 2
            elif is_s3_empty:
                return 3
            elif is_s1_empty:
                return 1
            return 2

        # 2. 일반 자재(모재 / 피팅 / 플랜지 등):
        # 2-1. 기준 1이 비어있거나, 기존 기준 1과 동일한 재질 계열이면 기준 1 갱신
        if is_s1_empty or (m_name in s1):
            return 1

        # 2-2. 기준 1에 이미 다른 자재가 들어가 있는 경우:
        # 차례대로 기준 2가 비어있으면 기준 2로 자동 입력
        if is_s2_empty or (m_name in s2):
            return 2

        # 2-3. 기준 2도 차있으면 기준 3으로 자동 입력
        if is_s3_empty or (m_name in s3):
            return 3

        # 모두 차있으면 기존 기준 보호를 위해 덮어쓰지 않음
        return None

    def apply_pmi_material_preset(self, mat_name, fill_all_specs=False):
        spec = self._lookup_pmi_material_spec(mat_name)
        if not spec: return False
        
        # 1. Chemical elements (Ni, Cr, Mo...)
        for idx, elem in enumerate(spec["elements"]):
            if idx < len(self.pmi_chem_elements):
                self.pmi_chem_elements[idx].set(elem)
                
        # 2. Chemical specs (기준 1, 2, 3 차례대로 스마트 순차 자동 입력)
        m_name = spec["name"]
        target_slot = self._determine_pmi_target_slot(spec)
        if target_slot is not None:
            target_idx = target_slot - 1

            if target_idx < len(self.pmi_chem_specs):
                self.pmi_chem_specs[target_idx].set(spec["specs"][0])

            mat_var = getattr(self, f'pmi_spec{target_slot}_mat', None)
            if mat_var:
                matched = next((opt for opt in self.PMI_COMMON_SPEC_OPTIONS if spec["name"] in opt), spec["name"])
                mat_var.set(matched)

        # fill_all_specs=True(명시적 플래그)일 때만 기준 2, 3을 자동 채움
        if fill_all_specs:
            if len(spec.get("specs", [])) > 1 and len(self.pmi_chem_specs) > 1:
                self.pmi_chem_specs[1].set(spec["specs"][1])
                if hasattr(self, 'pmi_spec2_mat'):
                    raw_s2 = spec["specs"][1].split(":")[0].strip()
                    matched_s2 = next((opt for opt in self.PMI_SPEC2_OPTIONS if raw_s2 in opt or opt in raw_s2), raw_s2)
                    self.pmi_spec2_mat.set(matched_s2)
            if len(spec.get("specs", [])) > 2 and len(self.pmi_chem_specs) > 2:
                self.pmi_chem_specs[2].set(spec["specs"][2])
                if hasattr(self, 'pmi_spec3_mat'):
                    raw_s3 = spec["specs"][2].split(":")[0].strip()
                    matched_s3 = next((opt for opt in self.PMI_SPEC3_OPTIONS if raw_s3 in opt or opt in raw_s3), raw_s3)
                    self.pmi_spec3_mat.set(matched_s3)
                
        # 3. element_filters (Min ~ Max)
        if hasattr(self, 'element_filters'):
            self.element_filters.clear()
            for elem_name, (f_min, f_max) in spec["limits"].items():
                min_str = f"{f_min:.2f}" if f_min > 0 else ""
                max_str = f"{f_max:.2f}" if f_max > 0 else ""
                self.element_filters.append({
                    'key': tk.StringVar(value=elem_name),
                    'min': tk.StringVar(value=min_str),
                    'max': tk.StringVar(value=max_str)
                })
            self._update_pmi_filter_ui()
            
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"📋 PMI 재질 기준 적용: {spec['name']}")
        return True
        
        # 4. If extracted data exists, re-evaluate automatically
        if getattr(self, 'extracted_data', None):
            self.evaluate_pmi_data_compliance(show_msg=False)
            
        return True

    def _on_pmi_item_changed(self, event=None):
        """품명 변경 시 해당 품명 형태(파이프, 엘보, 티, 레듀샤, 단조품, 주물, 피팅, 용접봉 등)에 맞는 재질 드롭다운 재정렬 및 기본 재질 자동 연동"""
        if not hasattr(self, 'gapji_item') or not hasattr(self, 'gapji_material'):
            return
        raw_item = self.gapji_item.get().strip().upper()
        if not raw_item: return

        target_cat = None
        if any(k in raw_item for k in ["PIPE", "파이프", "배관", "TUBE"]):
            target_cat = "PIPE"
        elif any(k in raw_item for k in ["ELBOW", "엘보"]):
            target_cat = "ELBOW"
        elif any(k in raw_item for k in ["TEE", "티"]):
            target_cat = "TEE"
        elif any(k in raw_item for k in ["REDUCER", "레듀샤", "레듀서"]):
            target_cat = "REDUCER"
        elif any(k in raw_item for k in ["FLANGE", "플랜지", "단조", "FORG"]):
            target_cat = "FLANGE"
        elif any(k in raw_item for k in ["VALVE", "밸브", "주물", "주조", "CAST"]):
            target_cat = "VALVE"
        elif any(k in raw_item for k in ["FITTING", "피팅", "이음쇠"]):
            target_cat = "FITTING"
        elif any(k in raw_item for k in ["WELD", "용접", "용접봉", "와이어", "WIRE", "ELECTRODE"]):
            target_cat = "WELD"
        elif any(k in raw_item for k in ["BEAM", "COLUMN", "GIRDER", "보", "기둥", "거더", "형강"]):
            target_cat = "BEAM/COLUMN"

        if target_cat and target_cat in self.PMI_ITEM_PRESETS:
            preset = self.PMI_ITEM_PRESETS[target_cat]
            rec_materials = list(preset["materials"])
            if hasattr(self, 'pmi_material_combobox') and self.pmi_material_combobox:
                self.pmi_material_combobox['values'] = rec_materials

            current_mat = self.gapji_material.get().strip()
            # 품명 변경 시: 해당 품명의 전용 재질 목록만 드롭다운에 표시하고 기본 재질 힌트 세팅.
            # 하단 기준 1, 2, 3은 사용자가 [입력] 또는 [➕ 다음 기준에 추가] 버튼을 누를 때 반영되도록 보존!
            if not current_mat or current_mat not in rec_materials:
                new_mat = preset["default_material"]
                self.gapji_material.set(new_mat)
            self._update_pmi_spec_labels()
        else:
            # 기타/미지정 품목인 경우 전체 재질 목록 제공
            if hasattr(self, 'pmi_material_combobox') and self.pmi_material_combobox:
                self.pmi_material_combobox['values'] = list(self.PMI_MATERIAL_SPECS.keys())

    def _on_pmi_material_changed(self, event=None):
        """재질 변경 시 품명 상호 연동 및 원소명 동기화 (하단 기준 슬롯은 [입력] 버튼으로 적용)"""
        mat = self.gapji_material.get().strip()
        if not mat: return

        spec = self._lookup_pmi_material_spec(mat)
        if spec:
            # 원소명(Ni, Cr, Mo...) 동기화
            for idx, elem in enumerate(spec["elements"]):
                if idx < len(self.pmi_chem_elements):
                    self.pmi_chem_elements[idx].set(elem)

            # 품명 역연동
            if hasattr(self, 'gapji_item'):
                cur_item = self.gapji_item.get().strip().upper()
                m_name = spec['name']
                new_item = None
                if "A182" in m_name:
                    new_item = "FLANGE"
                elif "A351" in m_name:
                    new_item = "VALVE"
                elif "A403" in m_name:
                    if cur_item in ("ELBOW", "TEE", "REDUCER"):
                        new_item = cur_item
                    else:
                        new_item = "ELBOW"
                elif any(k in m_name for k in ["ER316L", "ER308L", "ER309L", "ER2209", "ER70S"]):
                    new_item = "WELD"
                elif any(k in m_name for k in ["A992", "A572", "SM355", "SN355", "A36", "SS275"]):
                    new_item = "H-BEAM"
                elif "A312" in m_name:
                    new_item = "PIPE"

                if new_item and (not cur_item or cur_item in ("PIPE", "FLANGE", "VALVE", "FITTING", "ELBOW", "TEE", "REDUCER", "WELD", "H-BEAM", "BEAM", "HOT TAPPING")):
                    self.gapji_item.set(new_item)
                    self.save_settings()

    def _get_pmi_spec_item_title(self, slot_idx, mat_txt="", spec_txt=""):
        """기준 슬롯(1, 2, 3)의 재질 및 규격에 대응하는 품명(Item Name)을 도출"""
        mat_clean = mat_txt.strip()
        spec_clean = spec_txt.strip()
        if not mat_clean and not spec_clean:
            return ""
        if any(k in mat_clean for k in ["공란", "비움", "해당없음"]) and (not spec_clean or spec_clean == "-"):
            return ""

        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        combined = f"{mat_clean} {spec_clean}".upper()

        # 1. 용접봉/용접부
        if any(k in combined for k in ["ER316", "ER308", "ER309", "ER2209", "ER70S", "ERNICRMO", "E7018", "WELD", "용접"]):
            return "용접봉"

        # 2. 플랜지
        if any(k in combined for k in ["A182", "F316", "F304", "F51", "B564", "FLANGE", "플랜지"]):
            return "FLANGE"

        # 3. 밸브
        if any(k in combined for k in ["A351", "CF8M", "CF3M", "CF8", "VALVE", "밸브"]):
            return "VALVE"

        # 4. 피팅 (엘보/티/레듀샤)
        if any(k in combined for k in ["A403", "WP316", "WP304", "ELBOW", "TEE", "REDUCER", "엘보", "티", "레듀샤", "피팅"]):
            if "TEE" in combined or "티" in combined:
                return "TEE"
            if "REDUCER" in combined or "레듀샤" in combined:
                return "REDUCER"
            return "ELBOW"

        # 5. 형강/구조재
        if any(k in combined for k in ["A992", "A572", "SM355", "SN355", "A36", "SS275", "BEAM", "COLUMN", "GIRDER", "형강"]):
            return "H-BEAM"

        # 6. 배관 (PIPE)
        if any(k in combined for k in ["A312", "TP316", "TP304", "TP321", "TP347", "S31803", "PIPE", "배관"]):
            return "PIPE"

        # 7. 괄호 안에 품명 힌트가 있는 경우 (예: "피팅/엘보", "플랜지", "용접봉")
        if "(" in mat_clean and ")" in mat_clean:
            hint = mat_clean.split("(")[1].split(")")[0].strip()
            if hint: return hint

        # 8. 그 외 기본 재질명에서 품명 유추 불가 시
        if mat_clean and not any(k in mat_clean for k in ["공란", "비움"]):
            return mat_clean.split("(")[0].strip()

        return ""

    def _get_combined_pmi_item_names(self):
        """보고서 발행 및 갑지 표기 시 기준(1~6)에 언급된 모든 품명을 중복 없이 순서대로 수집하여 콤마로 연결"""
        collected = []
        # 1. 활성화된 기준 그룹에서 품명 수집
        grouped = self._collect_pmi_active_criteria() if hasattr(self, '_collect_pmi_active_criteria') else []
        for grp in grouped:
            item_txt = grp.get('item', '').strip()
            if item_txt:
                for part in item_txt.split(','):
                    p = part.strip()
                    if p and p not in collected and not any(k in p for k in ("공란", "비움", "-")):
                        collected.append(p)

        # 2. 기준 슬롯들 직접 검사 (grouped에 누락된 품명이 있는 경우 보완)
        specs_list = getattr(self, 'pmi_chem_specs', [])
        for slot_idx, spec_var in enumerate(specs_list, start=1):
            s_val = spec_var.get().strip()
            if not s_val or s_val == "-": continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            m_val = mat_var.get().strip() if mat_var else ""
            item_t = self._get_pmi_spec_item_title(slot_idx, m_val, s_val) if hasattr(self, '_get_pmi_spec_item_title') else ""
            if item_t:
                for part in item_t.split(','):
                    p = part.strip()
                    if p and p not in collected and not any(k in p for k in ("공란", "비움", "-")):
                        collected.append(p)

        # 3. 만약 수집된 것이 없다면 상단 기본 정보 폴백
        if not collected:
            base_item = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or self.config.get('GAPJI_ITEM', '')
            if base_item:
                return base_item
            return "PIPE"

        return ", ".join(collected)

    def _create_pmi_spec_slot_ui(self, slot_idx, container):
        """지정된 번호(1~6)의 기준 슬롯 UI 위젯 생성"""
        if not hasattr(self, 'pmi_spec_slots_ui'):
            self.pmi_spec_slots_ui = {}
        if slot_idx in self.pmi_spec_slots_ui:
            return

        idx_zero = slot_idx - 1
        while len(self.pmi_chem_specs) <= idx_zero:
            new_v = tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{len(self.pmi_chem_specs)+1}', ""))
            self.pmi_chem_specs.append(new_v)
            self.setting_vars[f'PMI_CHEM_SPEC_{len(self.pmi_chem_specs)}'] = new_v

        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if not mat_var:
            mat_var = tk.StringVar(value=self.config.get(f'PMI_SPEC{slot_idx}_MAT', ''))
            setattr(self, f'pmi_spec{slot_idx}_mat', mat_var)
            self.setting_vars[f'PMI_SPEC{slot_idx}_MAT'] = mat_var

        s_hdr = tk.Frame(container, background="#ffffff")
        s_hdr.pack(fill='x', pady=(2, 1))
        
        lbl = tk.Label(s_hdr, text=f"기준 {slot_idx}:", background="#ffffff", font=("Malgun Gothic", 8, "bold"))
        lbl.pack(side='left')
        setattr(self, f'pmi_spec{slot_idx}_lbl', lbl)

        ttk.Button(s_hdr, text="✕", width=2, command=lambda i=slot_idx: self.clear_pmi_spec(i)).pack(side='right', padx=(2, 0))
        ttk.Button(s_hdr, text="▼입력", width=4, command=lambda i=slot_idx: self.apply_current_to_spec(i)).pack(side='right', padx=(2, 0))

        options = getattr(self, f'PMI_SPEC{slot_idx}_OPTIONS', self.PMI_COMMON_SPEC_OPTIONS)
        combo = ttk.Combobox(
            s_hdr, textvariable=mat_var,
            values=options, width=20, state='readonly'
        )
        combo.pack(side='right')
        combo.bind('<<ComboboxSelected>>', lambda e, i=slot_idx: self._on_pmi_spec_selected(i))
        setattr(self, f'pmi_spec{slot_idx}_combo', combo)

        entry = ttk.Entry(container, textvariable=self.pmi_chem_specs[idx_zero])
        entry.pack(fill='x', padx=2, pady=(0, 3))

        self.pmi_spec_slots_ui[slot_idx] = {
            'header': s_hdr,
            'lbl': lbl,
            'combo': combo,
            'entry': entry,
            'mat_var': mat_var,
            'spec_var': self.pmi_chem_specs[idx_zero]
        }

    def add_pmi_spec_slot(self, show_msg=True):
        """새로운 기준 슬롯(4, 5, 6)을 동적으로 추가 (최대 6개)"""
        cur_count = len(self.pmi_chem_specs)
        if cur_count >= 6:
            if show_msg:
                messagebox.showinfo("알림", "갑지 성적서 1페이지 인쇄 규격상 기준 슬롯은 최대 6개까지 지원됩니다.")
            return None
        
        new_idx = cur_count + 1
        new_spec_var = tk.StringVar(value=self.config.get(f'PMI_CHEM_SPEC_{new_idx}', ""))
        new_mat_var = tk.StringVar(value=self.config.get(f'PMI_SPEC{new_idx}_MAT', ''))
        self.pmi_chem_specs.append(new_spec_var)
        setattr(self, f'pmi_spec{new_idx}_mat', new_mat_var)
        self.setting_vars[f'PMI_CHEM_SPEC_{new_idx}'] = new_spec_var
        self.setting_vars[f'PMI_SPEC{new_idx}_MAT'] = new_mat_var
        
        if hasattr(self, 'pmi_specs_container') and self.pmi_specs_container.winfo_exists():
            self._create_pmi_spec_slot_ui(new_idx, self.pmi_specs_container)
            self._update_pmi_spec_labels()
            if hasattr(self, 'gapji_item') and self.gapji_item.get():
                self._on_pmi_item_changed()
            try: self.root.update_idletasks()
            except: pass
            self.log(f"➕ 기준 슬롯 {new_idx} 추가 완료 (현재 총 {new_idx}개)")
        return new_idx

    def remove_pmi_spec_slot(self, idx):
        """동적으로 추가된 4~6번 슬롯 제거 (중간 슬롯 삭제 시 뒤쪽 슬롯 값을 앞으로 이동)"""
        cur_total = len(self.pmi_chem_specs)
        if idx <= 3 or idx > cur_total:
            return
        
        # 중간 슬롯 삭제 시 뒤쪽 슬롯 값들을 앞으로 한 칸씩 당김
        for s in range(idx, cur_total):
            next_s = s + 1
            cur_spec = self.pmi_chem_specs[next_s - 1].get()
            cur_mat = getattr(self, f'pmi_spec{next_s}_mat', None)
            cur_mat_val = cur_mat.get() if cur_mat else ""
            
            self.pmi_chem_specs[s - 1].set(cur_spec)
            target_mat = getattr(self, f'pmi_spec{s}_mat', None)
            if target_mat: target_mat.set(cur_mat_val)

        # 맨 마지막 슬롯의 UI 및 변수 파괴
        last_idx = cur_total
        slot_data = getattr(self, 'pmi_spec_slots_ui', {}).get(last_idx)
        if slot_data:
            if slot_data.get('header') and slot_data['header'].winfo_exists():
                slot_data['header'].destroy()
            if slot_data.get('entry') and slot_data['entry'].winfo_exists():
                slot_data['entry'].destroy()
            del self.pmi_spec_slots_ui[last_idx]
        
        self.pmi_chem_specs.pop()
        self.setting_vars.pop(f'PMI_CHEM_SPEC_{last_idx}', None)
        self.setting_vars.pop(f'PMI_SPEC{last_idx}_MAT', None)
        self.config.pop(f'PMI_CHEM_SPEC_{last_idx}', None)
        self.config.pop(f'PMI_SPEC{last_idx}_MAT', None)
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"➖ 기준 슬롯 제거 완료 (현재 총 {len(self.pmi_chem_specs)}개)")

    def reset_pmi_spec_slots_to_default(self, show_msg=False):
        """4번 이상 슬롯을 모두 닫아 기본 3개 상태로 원위치"""
        while len(self.pmi_chem_specs) > 3:
            last_idx = len(self.pmi_chem_specs)
            self.remove_pmi_spec_slot(last_idx)
        for i in (4, 5, 6):
            self.config.pop(f'PMI_CHEM_SPEC_{i}', None)
            self.config.pop(f'PMI_SPEC{i}_MAT', None)
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log("↩️ 기준 슬롯을 기본 3개로 원위치했습니다.")
        if show_msg:
            messagebox.showinfo("알림", "기준 슬롯을 기본 3개로 원위치했습니다.")

    def _update_pmi_spec_labels(self):
        """기준 라벨들을 재질명이 아닌 품명(Item Name)으로 동적 갱신"""
        num_slots = len(getattr(self, 'pmi_chem_specs', []))
        for slot_idx in range(1, num_slots + 1):
            lbl = getattr(self, f'pmi_spec{slot_idx}_lbl', None)
            if not lbl: continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            spec_var = self.pmi_chem_specs[slot_idx - 1] if slot_idx <= len(self.pmi_chem_specs) else None
            mat_txt = mat_var.get().strip() if mat_var else ""
            spec_txt = spec_var.get().strip() if spec_var else ""
            item_title = self._get_pmi_spec_item_title(slot_idx, mat_txt, spec_txt)
            if item_title:
                lbl.config(text=f"기준 {slot_idx} ({item_title}):")
            else:
                lbl.config(text=f"기준 {slot_idx}:")

    def add_current_to_next_spec(self, show_msg=True):
        """상단에서 선택된 품명 및 재질을 비어있는 다음 기준 슬롯(1 -> N)에 순차 입력 (부족 시 자동 슬롯 확장)"""
        cur_mat = self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else ""
        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        if not cur_mat:
            if cur_item:
                cur_mat = self._lookup_pmi_material_spec(cur_item)
                if cur_mat: cur_mat = cur_mat["name"]
        if not cur_mat:
            if show_msg:
                messagebox.showinfo("알림", "상단에서 먼저 재질을 선택해주세요.")
            return

        spec = self._lookup_pmi_material_spec(cur_mat)
        if not spec:
            if show_msg:
                messagebox.showinfo("알림", f"선택한 재질({cur_mat})에 대한 표준 규격을 찾을 수 없습니다.")
            return

        # 빈 슬롯(1 -> N) 찾기
        target_slot = None
        for slot in range(1, len(self.pmi_chem_specs) + 1):
            idx = slot - 1
            val = self.pmi_chem_specs[idx].get().strip()
            if not val or val == "-":
                target_slot = slot
                break

        # 기존 슬롯이 모두 채워진 경우 6개 미만이면 자동으로 슬롯 추가
        if target_slot is None:
            if len(self.pmi_chem_specs) < 6:
                target_slot = self.add_pmi_spec_slot(show_msg=show_msg)
            else:
                if show_msg:
                    messagebox.showinfo(
                        "알림",
                        "기준 슬롯 1~6이 이미 모두 채워져 있습니다.\n"
                        "수정을 원하시는 기준 옆의 [▼입력] 버튼을 누르시면 해당 기준이 즉시 변경됩니다."
                    )
                return

        if target_slot:
            self.apply_current_to_spec(target_slot)
            self.log(f"📋 기준 {target_slot}에 입력 완료: {cur_item} {spec['name']}")

    def apply_current_to_spec(self, slot_idx):
        """상단에서 선택된 품명 및 재질을 지정된 기준 슬롯에 즉시 입력"""
        cur_mat = self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else ""
        cur_item = self.gapji_item.get().strip().upper() if hasattr(self, 'gapji_item') else ""
        if not cur_mat:
            if cur_item:
                cur_mat = self._lookup_pmi_material_spec(cur_item)
                if cur_mat: cur_mat = cur_mat["name"]
        if not cur_mat:
            messagebox.showinfo("알림", "상단에서 먼저 재질을 선택해주세요.")
            return

        spec = self._lookup_pmi_material_spec(cur_mat)
        if not spec: return

        idx = slot_idx - 1
        if idx < len(self.pmi_chem_specs):
            spec_text = spec["specs"][0]
            if cur_item and cur_item not in spec_text.upper():
                spec_text = f"{spec['name']} ({cur_item}): " + spec_text.split(":", 1)[-1].strip()
            self.pmi_chem_specs[idx].set(spec_text)

        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if mat_var:
            matched = next((opt for opt in self.PMI_COMMON_SPEC_OPTIONS if spec["name"] in opt), spec["name"])
            mat_var.set(matched)

        if slot_idx == 1 or not getattr(self, 'element_filters', None):
            for e_idx, elem in enumerate(spec["elements"]):
                if e_idx < len(self.pmi_chem_elements):
                    self.pmi_chem_elements[e_idx].set(elem)
            if hasattr(self, 'element_filters'):
                self.element_filters.clear()
                for elem_name, (f_min, f_max) in spec["limits"].items():
                    self.element_filters.append({
                        'key': tk.StringVar(value=elem_name),
                        'min': tk.StringVar(value=f"{f_min:.2f}" if f_min > 0 else ""),
                        'max': tk.StringVar(value=f"{f_max:.2f}" if f_max > 0 else "")
                    })
                self._update_pmi_filter_ui()

        self._update_pmi_spec_labels()
        self.save_settings()
        self.log(f"📋 기준 {slot_idx}에 선택한 자재({spec['name']})를 입력했습니다.")

    def clear_pmi_spec(self, idx):
        """지정된 기준 슬롯을 공란으로 초기화하거나 4번 이상 슬롯은 즉시 닫기(제거)"""
        slot_idx = idx - 1
        if 0 <= slot_idx < len(self.pmi_chem_specs):
            # 4번 이상 확장 슬롯에서 ✕를 누르면 슬롯 자체를 즉시 닫아 원위치
            if idx > 3:
                # idx번부터 마지막까지 모두 비어있거나 현재 슬롯을 지우면 닫힘
                self.remove_pmi_spec_slot(idx)
                return
            self.pmi_chem_specs[slot_idx].set("")
            mat_var = getattr(self, f'pmi_spec{idx}_mat', None)
            if mat_var: mat_var.set("(공란 / 비움)")
            self.log(f"📋 기준 {idx}을 공란으로 비웠습니다.")
        self._update_pmi_spec_labels()
        self.save_settings()

    def clear_all_pmi_specs(self):
        """모든 기준 슬롯을 비우고 4~6번 슬롯은 기본 3개로 자동 원위치"""
        # 4번 이상 슬롯이 열려있다면 전부 닫아서 3개로 축소
        if hasattr(self, 'reset_pmi_spec_slots_to_default'):
            self.reset_pmi_spec_slots_to_default()
        for i in range(1, 4):
            slot_idx = i - 1
            if slot_idx < len(self.pmi_chem_specs):
                self.pmi_chem_specs[slot_idx].set("")
            mat_var = getattr(self, f'pmi_spec{i}_mat', None)
            if mat_var: mat_var.set("(공란 / 비움)")
        self._update_pmi_spec_labels()
        self.save_settings()
        self.log("📋 모든 기준 슬롯을 비우고 기본 3개로 원위치했습니다.")

    def _on_pmi_spec_selected(self, slot_idx, event=None):
        """기준 슬롯 콤보박스 선택 공통 핸들러"""
        mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
        if not mat_var: return
        val = mat_var.get().strip()
        idx = slot_idx - 1
        if idx >= len(self.pmi_chem_specs): return
        if not val or "비움" in val or val == "-":
            self.pmi_chem_specs[idx].set("")
            self.log(f"📋 기준 {slot_idx}: 공란 설정")
            self._update_pmi_spec_labels()
            self.save_settings()
            return
        spec = self._lookup_pmi_material_spec(val)
        if spec:
            self.pmi_chem_specs[idx].set(spec["specs"][0])
            if slot_idx == 1:
                self.gapji_material.set(spec["name"])
                self.apply_pmi_material_preset(spec["name"])
            self.log(f"📋 기준 {slot_idx} 설정: {spec['name']}")
            self.save_settings()
        self._update_pmi_spec_labels()

    def _on_pmi_spec1_selected(self, event=None):
        self._on_pmi_spec_selected(1, event)

    def _on_pmi_spec2_selected(self, event=None):
        self._on_pmi_spec_selected(2, event)

    def _on_pmi_spec3_selected(self, event=None):
        self._on_pmi_spec_selected(3, event)

    def _collect_pmi_active_criteria(self):
        """활성화된 기준 슬롯들로부터 화학성분 규격을 추출하고, 동일한 함량 기준끼리 품목명을 묶음(그룹화)"""
        import re
        raw_list = []
        specs_list = getattr(self, 'pmi_chem_specs', [])
        for slot_idx, spec_var in enumerate(specs_list, start=1):
            spec_txt = spec_var.get().strip()
            if not spec_txt or spec_txt == "-":
                continue
            mat_var = getattr(self, f'pmi_spec{slot_idx}_mat', None)
            mat_txt = mat_var.get().strip() if mat_var else ""
            item_title = self._get_pmi_spec_item_title(slot_idx, mat_txt, spec_txt)
            if not item_title:
                item_title = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or "PIPE"

            spec_obj = self._lookup_pmi_material_spec(mat_txt) or self._lookup_pmi_material_spec(spec_txt)
            spec_name = spec_obj.get('title', spec_obj['name']) if spec_obj else (mat_txt or spec_txt.split(':')[0].strip())

            limits = {}
            if spec_obj and spec_obj.get("limits"):
                limits = dict(spec_obj["limits"])
            else:
                matches = re.findall(r'([A-Za-z]{1,2})\s*:\s*([\d\.]+)\s*(?:~|-)\s*([\d\.]+)', spec_txt)
                for el, mi, ma in matches:
                    limits[el] = (float(mi), float(ma))

            raw_list.append({
                'slot': slot_idx,
                'name': spec_name,
                'item': item_title,
                'limits': limits,
                'spec_text': spec_txt
            })

        # 활성 슬롯에 내용이 없으면 상단 기본 정보 폴백
        if not raw_list:
            mat_input = (self.gapji_material.get().strip() if hasattr(self, 'gapji_material') else '') or self.config.get('GAPJI_MATERIAL', '')
            spec_obj = self._lookup_pmi_material_spec(mat_input) if hasattr(self, '_lookup_pmi_material_spec') else None
            item_name = (self.gapji_item.get().strip() if hasattr(self, 'gapji_item') else '') or self.config.get('GAPJI_ITEM', 'PIPE')
            if spec_obj:
                raw_list.append({
                    'slot': 1,
                    'name': spec_obj.get('title', spec_obj['name']),
                    'item': item_name,
                    'limits': dict(spec_obj.get("limits", {})),
                    'spec_text': spec_obj.get("specs", [""])[0]
                })

        # [원칙 1] 동일 함량 자동 묶기 (그룹화)
        grouped = []
        for item_data in raw_list:
            matched_grp = None
            for grp in grouped:
                if grp['limits'] == item_data['limits'] and grp['limits']:
                    matched_grp = grp
                    break
            if matched_grp:
                cur_items = [x.strip() for x in matched_grp['item'].split(',') if x.strip()]
                if item_data['item'] not in cur_items:
                    cur_items.append(item_data['item'])
                    matched_grp['item'] = ", ".join(cur_items)
                cur_names = [x.strip() for x in matched_grp['name'].split(',') if x.strip()]
                if item_data['name'] not in cur_names:
                    cur_names.append(item_data['name'])
                    matched_grp['name'] = ", ".join(cur_names)
            else:
                grouped.append({
                    'name': item_data['name'],
                    'item': item_data['item'],
                    'limits': item_data['limits'],
                    'spec_text': item_data['spec_text']
                })
        return grouped

    def _apply_pmi_acceptance_criteria(self, ws):
        """갑지 1. Acceptance Criteria (%) 표(24~39행, A~T열)를 [방안 B] 가로 비교 매트릭스 형태로 작성"""
        if ws.cell(row=24, column=1).value is None:
            return

        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils.cell import range_boundaries

        def _style_range(ws_obj, cell_range, font=None, alignment=None, border=None):
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
            for row in ws_obj.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    if font: cell.font = font
                    if alignment: cell.alignment = alignment
                    if border: cell.border = border

        grouped = self._collect_pmi_active_criteria()
        if not grouped:
            return

        # 1. 기존 24~39행, A~T열의 병합 셀 해제 및 내용 초기화
        to_unmerge = [rng for rng in ws.merged_cells.ranges if 24 <= rng.min_row <= 39 and (rng.min_col <= 20 or rng.max_col <= 20)]
        for rng in to_unmerge:
            try: ws.unmerge_cells(str(rng))
            except: pass
        for r in range(24, 40):
            for c in range(1, 21):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()

        num_groups = len(grouped)

        # 스타일 정의
        thin_side = Side(style='thin')
        b_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        f_title = Font(name='Times New Roman', size=11, bold=True)
        # 기준 개수(열 너비)에 따라 헤더 폰트 크기 세분화 (글자 가려짐 완전 방지):
        # - 1개 기준: 10pt
        # - 2개 기준 (7열 너비): 8.5pt
        # - 3개 이상 기준 (5열 너비): 8pt (긴 규격명과 품목명이 여유 있게 한눈에 표시됨)
        if num_groups >= 3:
            hdr_size = 8
        elif num_groups == 2:
            hdr_size = 8.5
        else:
            hdr_size = 10
        f_hdr = Font(name='Times New Roman', size=hdr_size, bold=True)
        f_data = Font(name='Times New Roman', size=10, bold=False)
        f_note = Font(name='Times New Roman', size=9, bold=False)
        align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)

        elem_full_names = {
            'Cr': 'Chromium', 'Ni': 'Nickel', 'Mo': 'Molybdenum',
            'Mn': 'Manganese', 'Si': 'Silicon', 'V': 'Vanadium',
            'Nb': 'Niobium', 'Ti': 'Titanium', 'Cu': 'Copper',
            'P': 'Phosphorus', 'S': 'Sulfur'
        }

        # 대상 원소 선정 (Cr, Ni, Mo 우선, 최대 3개)
        all_elems = []
        for g in grouped:
            for k in g['limits'].keys():
                if k not in all_elems and k != "-":
                    all_elems.append(k)
        target_elems = []
        for p in ["Cr", "Ni", "Mo"]:
            if p in all_elems:
                target_elems.append(p)
        for e in all_elems:
            if e not in target_elems:
                target_elems.append(e)
        if not target_elems:
            target_elems = ["Cr", "Ni", "Mo"]
        target_elems = target_elems[:3]

        def _format_limit(limits_dict, el_key):
            if el_key in limits_dict:
                l_min, l_max = limits_dict[el_key]
                if l_min > 0 and l_max < 900.0:
                    return f"{l_min:.2f}~{l_max:.2f}"
                elif l_min > 0:
                    return f"Min {l_min:.2f}"
                elif l_max < 900.0:
                    return f"Max {l_max:.2f}"
            return "-"

        def _get_hdr_label(grp):
            name = grp.get('name', '')
            item = grp.get('item', '')
            if item:
                return f"{name}\n({item})"
            return name

        # [케이스 1] 고유 기준 1개: 기존 10열(A~J) 단일 표
        if num_groups == 1:
            grp = grouped[0]
            ws.merge_cells('A24:J26')
            ws['A24'] = f"1. Acceptance Criteria (%)\n{grp['name']} — {grp['item']}"
            _style_range(ws, 'A24:J26', font=f_title, alignment=align_c, border=b_thin)

            ws.row_dimensions[27].height = 20.0
            ws.merge_cells('A27:F27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:F27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('G27:J27')
            ws['G27'] = 'Composition, %'
            _style_range(ws, 'G27:J27', font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                val_txt = _format_limit(grp['limits'], el_k)
                ws.merge_cells(f'A{idx}:F{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:F{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'G{idx}:J{idx}')
                ws[f'G{idx}'] = val_txt
                _style_range(ws, f'G{idx}:J{idx}', font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:J31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:J31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 2] 고유 기준 2개: [방안 B] 가로 비교 2열 매트릭스 (A~F: 원소, G~M: 기준1, N~T: 기준2)
        elif num_groups == 2:
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 27행 헤더 2줄 텍스트가 잘리지 않도록 높이 34pt 지정, 빈 32행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[27].height = 34.0
            ws.row_dimensions[32].height = 1.77

            ws.merge_cells('A27:F27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:F27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('G27:M27')
            ws['G27'] = _get_hdr_label(grouped[0])
            _style_range(ws, 'G27:M27', font=f_hdr, alignment=align_c, border=b_thin)

            ws.merge_cells('N27:T27')
            ws['N27'] = _get_hdr_label(grouped[1])
            _style_range(ws, 'N27:T27', font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                v1 = _format_limit(grouped[0]['limits'], el_k)
                v2 = _format_limit(grouped[1]['limits'], el_k)

                ws.merge_cells(f'A{idx}:F{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:F{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'G{idx}:M{idx}')
                ws[f'G{idx}'] = v1
                _style_range(ws, f'G{idx}:M{idx}', font=f_data, alignment=align_c, border=b_thin)

                ws.merge_cells(f'N{idx}:T{idx}')
                ws[f'N{idx}'] = v2
                _style_range(ws, f'N{idx}:T{idx}', font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 3] 고유 기준 3개: [방안 B] 가로 비교 3열 매트릭스 (A~E, F~J, K~O, P~T)
        elif num_groups == 3:
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 27행 헤더 2줄 텍스트가 잘리지 않도록 높이 34pt 지정, 빈 32행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[27].height = 34.0
            ws.row_dimensions[32].height = 1.77

            ws.merge_cells('A27:E27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:E27', font=f_hdr, alignment=align_c, border=b_thin)

            ranges_col = [('F', 'J'), ('K', 'O'), ('P', 'T')]
            for g_i, (c_start, c_end) in enumerate(ranges_col):
                rng_txt = f"{c_start}27:{c_end}27"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}27"] = _get_hdr_label(grouped[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)

                for g_i, (c_start, c_end) in enumerate(ranges_col):
                    v = _format_limit(grouped[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

        # [케이스 4] 고유 기준 4~6개: 상단(24~31행) 1차 + 하단(32~39행) 2차 분할 2단 매트릭스
        else:
            tier1 = grouped[:3]
            tier2 = grouped[3:6]

            # 1단 (24~31행)
            ws.merge_cells('A24:T26')
            ws['A24'] = '1. Acceptance Criteria (%)'
            _style_range(ws, 'A24:T26', font=f_title, alignment=align_c, border=b_thin)

            # 헤더 높이 32pt씩 부여하고, 25/26행에서 상쇄하여 총 세로높이 100% 보존
            ws.row_dimensions[25].height = 10.885
            ws.row_dimensions[26].height = 10.885
            ws.row_dimensions[27].height = 32.0

            ws.merge_cells('A27:E27')
            ws['A27'] = 'Element'
            _style_range(ws, 'A27:E27', font=f_hdr, alignment=align_c, border=b_thin)

            ranges_col = [('F', 'J'), ('K', 'O'), ('P', 'T')]
            for g_i in range(len(tier1)):
                c_start, c_end = ranges_col[g_i]
                rng_txt = f"{c_start}27:{c_end}27"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}27"] = _get_hdr_label(tier1[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=28):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)
                for g_i in range(len(tier1)):
                    c_start, c_end = ranges_col[g_i]
                    v = _format_limit(tier1[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A31:T31')
            ws['A31'] = '* Tolerance ±10%'
            _style_range(ws, 'A31:T31', font=f_note, alignment=align_c, border=b_thin)

            # 2단 (32~39행)
            ws.merge_cells('A32:T33')
            ws['A32'] = '1. Acceptance Criteria (%) - Continued'
            _style_range(ws, 'A32:T33', font=f_title, alignment=align_c, border=b_thin)

            ws.row_dimensions[34].height = 32.0
            ws.merge_cells('A34:E34')
            ws['A34'] = 'Element'
            _style_range(ws, 'A34:E34', font=f_hdr, alignment=align_c, border=b_thin)

            for g_i in range(len(tier2)):
                c_start, c_end = ranges_col[g_i]
                rng_txt = f"{c_start}34:{c_end}34"
                ws.merge_cells(rng_txt)
                ws[f"{c_start}34"] = _get_hdr_label(tier2[g_i])
                _style_range(ws, rng_txt, font=f_hdr, alignment=align_c, border=b_thin)

            for idx, el_k in enumerate(target_elems, start=35):
                full_n = elem_full_names.get(el_k, el_k)
                ws.merge_cells(f'A{idx}:E{idx}')
                ws[f'A{idx}'] = full_n
                _style_range(ws, f'A{idx}:E{idx}', font=f_data, alignment=align_c, border=b_thin)
                for g_i in range(len(tier2)):
                    c_start, c_end = ranges_col[g_i]
                    v = _format_limit(tier2[g_i]['limits'], el_k)
                    rng_txt = f"{c_start}{idx}:{c_end}{idx}"
                    ws.merge_cells(rng_txt)
                    ws[f"{c_start}{idx}"] = v
                    _style_range(ws, rng_txt, font=f_data, alignment=align_c, border=b_thin)

            ws.merge_cells('A38:T38')
            ws['A38'] = '* Tolerance ±10%'
            _style_range(ws, 'A38:T38', font=f_note, alignment=align_c, border=b_thin)

    def evaluate_pmi_data_compliance(self, show_msg=True):
        """추출된 PMI 측정 데이터를 현재 재질 기준(또는 element_filters)과 비교하여 합/부(ALL MATCH / MISMATCH) 자동 판정"""
        if not getattr(self, 'extracted_data', None):
            if show_msg:
                messagebox.showinfo("알림", "판정할 PMI 추출 데이터가 없습니다.")
            return

        limits = {}
        if hasattr(self, 'element_filters') and self.element_filters:
            for f in self.element_filters:
                k = f['key'].get().strip()
                if not k: continue
                min_v = self.to_float(f['min'].get())
                max_v = self.to_float(f['max'].get())
                if min_v > 0 or max_v > 0:
                    limits[k] = (min_v, max_v if max_v > 0 else 999.9)

        if not limits:
            spec = self._lookup_pmi_material_spec(self.gapji_material.get().strip())
            if spec:
                limits = spec["limits"]

        if not limits:
            if show_msg:
                messagebox.showwarning("경고", "판정할 원소 기준(Min ~ Max)이 설정되지 않았습니다.\n재질을 선택하거나 성분 필터를 입력해주세요.")
            return

        pass_count = 0
        fail_count = 0
        margin = 0.10 # 10% tolerance (갑지 Acceptance Criteria 규격)

        for item in self.extracted_data:
            mismatches = []
            for elem_k, (f_min, f_max) in limits.items():
                val = self.to_float(self._get_val_ci(item, elem_k))
                tol_min = round(f_min * (1.0 - margin), 2) if f_min > 0 else 0.0
                tol_max = round(f_max * (1.0 + margin), 2) if f_max > 0 else 999.9

                if f_min > 0 and val < tol_min:
                    mismatches.append(f"{elem_k}미달({val:.2f}%)")
                elif f_max > 0 and val > tol_max:
                    mismatches.append(f"{elem_k}초과({val:.2f}%)")

            if not mismatches:
                item['Grade'] = "ALL MATCH"
                item['Result'] = "ALL MATCH"
                item['is_deficient'] = False
                pass_count += 1
            else:
                reason = ", ".join(mismatches)
                item['Grade'] = f"MISMATCH ({reason})"
                item['Result'] = "MISMATCH"
                item['is_deficient'] = True
                fail_count += 1

        self.populate_preview(self.extracted_data, switch_tab=False, mode="PMI")
        msg = f"⚖️ 규격 판정 완료: 전체 {len(self.extracted_data)}건 중 합격(ALL MATCH) {pass_count}건, 부적합(MISMATCH) {fail_count}건"
        self.log(msg)
        if show_msg:
            messagebox.showinfo("판정 결과", msg)

    def show_pmi_element_selector_dialog(self):
        """추출된 데이터 또는 전체 화학 원소 목록에서 리포트에 표기할 원소(최대 5개)를 선택하는 다이얼로그"""
        detected_elements = set()
        if getattr(self, 'extracted_data', None):
            ignore_keys = {"_status", "selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Grade", "Result", "Remarks", "Welder", "order_index"}
            for item in self.extracted_data[:30]:
                for k, v in item.items():
                    if k not in ignore_keys and isinstance(v, (int, float)) and v > 0:
                        detected_elements.add(k)
        
        common_elements = ["Ni", "Cr", "Mo", "Ti", "Nb", "Mn", "Si", "Cu", "V", "W", "Co", "Fe", "Al", "C", "P", "S"]
        ordered_candidates = [e for e in common_elements if e in detected_elements] + [e for e in detected_elements if e not in common_elements] + [e for e in common_elements if e not in detected_elements]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("PMI 출력 원소 선택")
        dialog.geometry("450x340")
        dialog.minsize(420, 300)
        dialog.configure(background="#f9fafb")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_f = ttk.Frame(dialog, padding=15)
        main_f.pack(fill='both', expand=True)
        
        # 1. 상단 안내
        top_f = ttk.Frame(main_f)
        top_f.pack(fill='x', pady=(0, 10))
        ttk.Label(top_f, text="🔬 성적서 출력 원소 선택 (최대 5개)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 3))
        ttk.Label(top_f, text="출력할 화학 원소를 체크하세요. (선택된 순서대로 성적서에 반영됩니다)", font=("Malgun Gothic", 8), foreground="#6b7280").pack(anchor='w')
        
        # 2. 하단 버튼 영역 (창 크기와 무관하게 항상 보이도록 side='bottom'으로 먼저 패킹)
        btn_f = ttk.Frame(main_f)
        btn_f.pack(side='bottom', fill='x', pady=(10, 0))
        
        def apply_selected():
            chosen = [e for e, v in check_vars.items() if v.get()]
            if not chosen:
                messagebox.showwarning("경고", "최소 1개 이상의 원소를 선택해주세요.")
                return
            if len(chosen) > 5:
                messagebox.showwarning("경고", f"성적서 양식상 최대 5개 원소까지만 선택 가능합니다.\n(현재 {len(chosen)}개 선택됨)")
                return
                
            for idx in range(5):
                val = chosen[idx] if idx < len(chosen) else "-"
                self.pmi_chem_elements[idx].set(val)
                
            self.save_settings()
            self.log(f"✅ PMI 출력 원소 설정 완료: {', '.join(chosen)}")
            dialog.destroy()
            
        ttk.Button(btn_f, text=" 적용 (Apply) ", command=apply_selected, width=14).pack(side='right', padx=5)
        ttk.Button(btn_f, text=" 취소 (Cancel) ", command=dialog.destroy, width=14).pack(side='right')

        # 3. 중앙 체크박스 영역 (4열 배치로 콤팩트하게 정렬)
        content_box = ttk.LabelFrame(main_f, text=" 화학 원소 목록 (선택) ", padding=10)
        content_box.pack(fill='both', expand=True, pady=5)
        for col_i in range(4):
            content_box.columnconfigure(col_i, weight=1)

        current_selected = [var.get().strip() for var in self.pmi_chem_elements if var.get().strip() and var.get().strip() != '-']
        check_vars = {}
        for i, elem in enumerate(ordered_candidates[:16]):
            var = tk.BooleanVar(value=(elem in current_selected))
            check_vars[elem] = var
            chk = ttk.Checkbutton(content_box, text=elem, variable=var)
            chk.grid(row=i//4, column=i%4, sticky='w', padx=12, pady=6)

        return dialog

    def check_material_grade(self, row_data):
        """jjchRFIPMI.py에서 이식된 10% 여유치 판정 로직"""
        cr = self.to_float(self._get_val_ci(row_data, 'Cr'))
        ni = self.to_float(self._get_val_ci(row_data, 'Ni'))
        mo = self.to_float(self._get_val_ci(row_data, 'Mo'))
        mn = self.to_float(self._get_val_ci(row_data, 'Mn'))
        
        margin = 0.1 # 10% 여유
        
        # 1. SUS 316 (Cr:16~18 / Ni:10~14 / Mo:2~3)
        if (16.0*(1-margin) <= cr <= 18.0*(1+margin)) and (10.0*(1-margin) <= ni <= 14.0*(1+margin)) and (2.0*(1-margin) <= mo <= 3.0*(1+margin)):
            return "SS316"
        
        # 2. Duplex (Cr:22~23 / Mo:3~3.5 / Ni:4.5~6.5 / Mn:2.0이하)
        if (22.0*(1-margin) <= cr <= 23.0*(1+margin)) and (4.5*(1-margin) <= ni <= 6.5*(1+margin)) and (3.0*(1-margin) <= mo <= 3.5*(1+margin)) and (mn <= 2.2):
            return "DUPLEX"

        # 3. SUS 310 (Cr:24~26 / Ni:19~22)
        if (24.0*(1-margin) <= cr <= 26.0*(1+margin)) and (19.0*(1-margin) <= ni <= 22.0*(1+margin)):
            return "SS310"

        # 4. SUS 304 (Cr:18↑ / Ni:8↑ / Mo:0.5↓ / Mn:2.0↓)
        if (cr >= 16.2) and (ni >= 7.2) and (mo <= 0.55) and (mn <= 2.2):
            return "SS304"

        return None # 판정 불가 시 원본 사용

    def normalize_iso(self, val):
        """정렬 및 그룹화 시와 동일하게 정규화된 값으로 비교하여 도면번호 일관성 유지"""
        s_val = str(val).strip()
        try:
            if re.match(r'^\d+(\.\d+)?$', s_val):
                f_val = float(s_val)
                if f_val == int(f_val): return str(int(f_val))
                return str(f_val)
        except: pass
        return s_val.lower()

    def find_image_smart(self, keyword, exclude_keyword=None):
        def _search_in_folder(folder_path):
            if not folder_path or not os.path.exists(folder_path): return None
            candidates = glob.glob(os.path.join(folder_path, "*.*"))
            valid_extensions = ['.PNG', '.JPG', '.JPEG', '.BMP', '.GIF']
            # Prefer an exact file-stem match before falling back to substring
            # matching (PMI must not resolve to PMI-1 or PMI갑).
            for path in candidates:
                fname = os.path.basename(path); ext = os.path.splitext(path)[1].upper()
                if ext not in valid_extensions: continue
                if os.path.splitext(fname)[0].upper() == keyword.upper():
                    if exclude_keyword and exclude_keyword.upper() in fname.upper(): continue
                    return path
            for path in candidates:
                fname = os.path.basename(path).upper(); ext = os.path.splitext(path)[1].upper()
                if ext not in valid_extensions: continue 
                if keyword.upper() in fname:
                    if exclude_keyword and exclude_keyword.upper() in fname: continue
                    return path 
            return None

        # 1. UI에서 설정한 폴더에서 먼저 검색
        found = _search_in_folder(self.logo_folder_path.get())
        if found: return found

        # Use the bundled project resources when no logo folder/path was selected.
        # This is required for PMI generation because the data-sheet images are
        # cleared before the configured logos are added again.
        app_home = os.path.abspath(
            os.path.join(os.path.dirname(__file__), '..', '..', '..', '..')
        )
        for resource_dir in (
            os.path.join(app_home, 'resources'),
            os.path.join(os.path.dirname(app_home), 'Assets'),
        ):
            found = _search_in_folder(resource_dir)
            if found:
                return found
        
        # 2. PyInstaller 묶음(실행파일 내부 임시 폴더)에서 검색 (Standalone 지원)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            found = _search_in_folder(sys._MEIPASS)
            if found: return found
            
        return None

    def place_image_freely(self, ws, img_path, anchor_cell_str, w, h, x_offset, y_offset):
        try:
            if not img_path or not os.path.exists(img_path): return
            
            # [ENHANCED] Robust image processing with terminal & UI logging
            original = PILImage.open(img_path).convert("RGBA")
            w_px = max(5, int(float(w))); h_px = max(5, int(float(h)))
            
            temp_name = f"temp_{int(time.time())}_{os.path.basename(img_path)}.png"
            temp_full_path = os.path.join(tempfile.gettempdir(), temp_name)
            
            resized = original.resize((w_px, h_px), PILImage.Resampling.LANCZOS)
            resized.save(temp_full_path, "PNG")
            
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(temp_full_path)
            img.width = w_px; img.height = h_px
            from openpyxl.utils import column_index_from_string
            from openpyxl.utils.cell import coordinate_from_string
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
            
            # [FIX] Handle 'None' or empty anchor string
            if not anchor_cell_str or str(anchor_cell_str).strip() == "None" or len(str(anchor_cell_str).strip()) < 2:
                anchor_cell_str = "A1"

            try:
                col_str, row_num = coordinate_from_string(anchor_cell_str)
                col_idx = max(0, column_index_from_string(col_str) - 1)
                row_idx = max(0, row_num - 1) 
            except Exception as e:
                col_idx, row_idx = 0, 0
                anchor_cell_str = "A1"
            
            # [STABLE] Use OneCellAnchor instead of AbsoluteAnchor for better Excel compatibility (prevents corruption)
            # 1 Pixel = 9525 EMU (Standard 96 DPI)
            final_emu_x = max(0, int(float(x_offset) * 9525))
            final_emu_y = max(0, int(float(y_offset) * 9525))
            emu_w = max(9525, int(float(w_px) * 9525))
            emu_h = max(9525, int(float(h_px) * 9525))
            
            # [FIX] Use AnchorMarker (cell + offset) for stable positioning
            marker = AnchorMarker(col=col_idx, colOff=final_emu_x, row=row_idx, rowOff=final_emu_y)
            size = XDRPositiveSize2D(cx=emu_w, cy=emu_h)
            
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(img)
            
            msg = f"[SUCCESS] 로고 배치 완료: {os.path.basename(img_path)}"
            self.log(f"   {msg}")
        except Exception as e:
            msg = f"[ERROR] 로고 배치 실패: {e}"
            self.log(f"   {msg}")

    @staticmethod
    def _white_background_to_transparent_png(source):
        """이미지의 흰 배경만 제거하고 원본 색상/획은 그대로 유지한다."""
        if isinstance(source, (bytes, bytearray)):
            source = io.BytesIO(source)

        with PILImage.open(source) as original:
            image = original.convert("RGBA")
            converted = []
            for red, green, blue, alpha in image.getdata():
                # 완전한 흰색은 투명하게, 가장자리의 밝은 픽셀은 부드럽게
                # 반투명 처리하여 흰 테두리가 남지 않도록 한다.
                whiteness = min(red, green, blue)
                if whiteness >= 248:
                    new_alpha = 0
                elif whiteness >= 210:
                    new_alpha = int(alpha * (248 - whiteness) / 38)
                else:
                    new_alpha = alpha
                converted.append((red, green, blue, new_alpha))
            image.putdata(converted)

            output = io.BytesIO()
            image.save(output, format="PNG")
            output.seek(0)
            return output

    def _prepare_paut_cover_identity(self, ws):
        """PAUT 갑지 직인/검사자 이름/서명을 투명 이미지로 정리한다."""
        # 원본 양식의 이름 셀 서식을 유지하면서 검사자만 교체한다.
        self.safe_set_value(ws, "D45", "박 광 복")
        self.safe_set_value(ws, "AA44", "한국지역난방공사 중앙지사")

        signature_path = os.path.join(SRC_ROOT or "", "signs", "박광복.png")
        if not os.path.exists(signature_path):
            self.log(f"⚠️ 박광복 서명 파일을 찾을 수 없습니다: {signature_path}")
            return

        reviewer_signature_path = os.path.join(
            SRC_ROOT or "", "signs", "주진철.png"
        )
        if not os.path.exists(reviewer_signature_path):
            self.log(
                f"⚠️ 주진철 서명 파일을 찾을 수 없습니다: "
                f"{reviewer_signature_path}"
            )
            return

        # 원본 양식 기준 그림 위치(0부터 시작):
        # 직인 Q3, 검사자 서명 I44, 판독자 서명 S44.
        replacement_specs = {
            (16, 2): None,
            (8, 43): signature_path,
            (18, 43): reviewer_signature_path,
        }
        replaced_images = []

        for original_image in list(getattr(ws, "_images", [])):
            marker = getattr(getattr(original_image, "anchor", None), "_from", None)
            position = (
                getattr(marker, "col", None),
                getattr(marker, "row", None),
            )
            if position not in replacement_specs:
                replaced_images.append(original_image)
                continue

            try:
                replacement_source = replacement_specs[position]
                if replacement_source:
                    transparent_stream = self._white_background_to_transparent_png(
                        replacement_source
                    )
                else:
                    transparent_stream = self._white_background_to_transparent_png(
                        original_image._data()
                    )

                replacement = XLImage(transparent_stream)
                replacement.anchor = copy.deepcopy(original_image.anchor)
                # workbook 저장 시까지 BytesIO가 닫히지 않게 참조를 유지한다.
                replacement._pmi_source_stream = transparent_stream
                replaced_images.append(replacement)
            except Exception as exc:
                replaced_images.append(original_image)
                self.log(f"⚠️ PAUT 직인/서명 투명화 실패({position}): {exc}")

        ws._images = replaced_images
        self.log(
            "✅ PAUT 갑지 직인 투명화 및 박광복/주진철 투명 서명 교체 완료"
        )

    def _prepare_paut_data_stamp(self, ws):
        """PAUT 두 번째 시트 헤더 직인의 흰 배경만 투명하게 만든다."""
        # 원본 양식 drawing2 기준 직인 위치(0부터 시작): O2.
        stamp_position = (14, 1)
        replaced_images = []
        stamp_replaced = False

        for original_image in list(getattr(ws, "_images", [])):
            marker = getattr(getattr(original_image, "anchor", None), "_from", None)
            position = (
                getattr(marker, "col", None),
                getattr(marker, "row", None),
            )
            if position != stamp_position:
                replaced_images.append(original_image)
                continue

            try:
                transparent_stream = self._white_background_to_transparent_png(
                    original_image._data()
                )
                replacement = XLImage(transparent_stream)
                replacement.anchor = copy.deepcopy(original_image.anchor)
                replacement._pmi_source_stream = transparent_stream
                replaced_images.append(replacement)
                stamp_replaced = True
            except Exception as exc:
                replaced_images.append(original_image)
                self.log(f"⚠️ PAUT 두 번째 시트 직인 투명화 실패: {exc}")

        ws._images = replaced_images
        if stamp_replaced:
            self.log("✅ PAUT 두 번째 시트 헤더 직인 투명화 완료")
        else:
            self.log("⚠️ PAUT 두 번째 시트 헤더 직인을 찾지 못했습니다.")

    def load_daily_work_history(self):
        """사업장별 작업 감독일보에서 확인한 NDT 항목만 읽어 가져옵니다."""
        try:
            try:
                main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
                if "RT" in main_tab:
                    sub_tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                    mode = "KOGAS" if "가스공사" in sub_tab_text else "RT"
                elif "PT" in main_tab: mode = "PT"
                elif "MT" in main_tab: mode = "MT"
                elif "PAUT" in main_tab: mode = "PAUT"
                else: mode = "PMI"
            except: mode = "PMI"
            
            from site_apps.central.src.reports.ndt_history_import import select_history_rows
            ndt_list = select_history_rows(self.root, mode)
            if not ndt_list:
                return
            tree, idx_map, data, keys_attr = self._get_mode_info(mode)
            
            # Map NDT data to the tree columns
            source_site = ndt_list[0]['_source_site']
            if data and any(row.get('_source_site') != source_site for row in data):
                if not messagebox.askyesno(
                    '사업장 확인',
                    f'현재 보고서에 다른 사업장 또는 출처 미확인 자료가 있습니다.\n{source_site} 자료를 함께 추가할까요?',
                    parent=self.root,
                ):
                    return
            existing_ids = {row.get('_source_id') for row in data if row.get('_source_id')}
            added_count = 0
            for r in ndt_list:
                if r.get('_source_id') and r['_source_id'] in existing_ids:
                    continue
                new_row = {k: "" for k in keys_attr}
                new_row['_source_site'] = r['_source_site']
                new_row['_source_path'] = r['_source_path']
                new_row['_source_id'] = r.get('_source_id')
                new_row['_source_contractor'] = r.get('업체', '')
                new_row['_source_tested_length'] = r.get('검사길이', '') or r.get(mode, '')
                new_row['_source_work_type'] = r.get('근무구분', '')
                for quantity_key in ('RT_OR', 'RT_RE'):
                    if quantity_key in r:
                        new_row[quantity_key] = r[quantity_key]
                new_row["Date"] = r.get("Date", "")

                # 중앙지사 작업/감독일보의 '구간'은 RT의 SEC 값이다.
                if mode in ("RT", "KOGAS"):
                    new_row["Sec"] = r.get("구간", "")
                 
                # 라인번호 / ISO / Dwg
                dwg = r.get("라인번호", "")
                if "ISO" in keys_attr: new_row["ISO"] = dwg
                elif "Line No." in keys_attr: new_row["Line No."] = dwg
                elif "Dwg" in keys_attr: new_row["Dwg"] = dwg
                elif "Drawing No." in keys_attr: new_row["Drawing No."] = dwg
                
                # Joint No.
                joint = r.get("Joint No.", "")
                if "Joint" in keys_attr: new_row["Joint"] = joint
                elif "Joint No." in keys_attr: new_row["Joint No."] = joint
                
                # 관경 / Size / T / Th'k
                size = r.get("관경", "")
                thickness = r.get("두께", "")
                if "Size" in keys_attr: new_row["Size"] = size
                if "T" in keys_attr: new_row["T"] = thickness
                if "Th'k(mm)" in keys_attr: new_row["Th'k(mm)"] = thickness
                if not thickness and any(k in keys_attr for k in ('T', "Th'k(mm)")):
                    new_row['Remarks'] = '두께 확인 필요'
                if any(k in keys_attr for k in ('Material', 'Mat')):
                    from site_apps.central.src.reports.pipe_material_spec import infer_row_pipe_material
                    inferred_mat = infer_row_pipe_material(
                        {'Size': size, 'Line No.': dwg},
                        item_name=self._mode_item_name(mode) if hasattr(self, '_mode_item_name') else '',
                    )
                    if inferred_mat:
                        if 'Material' in keys_attr and not new_row.get('Material'):
                            new_row['Material'] = inferred_mat
                        if 'Mat' in keys_attr and not new_row.get('Mat'):
                            new_row['Mat'] = inferred_mat
                
                # 용접사 / Welder
                welder = self._normalize_daily_photo_welder(
                    r.get("용접사", "")
                )
                if "Welder" in keys_attr: new_row["Welder"] = welder
                
                # 검사길이 / Tested Length
                test_len = r.get("검사길이", "")
                if not test_len and mode in ["PAUT", "MT", "PT"]:
                    test_len = r.get(mode, "")
                if "Tested Length" in keys_attr: new_row["Tested Length"] = test_len
                
                # 구간 / Loc
                # RT의 Film Location은 중앙지사 작업/감독일보
                # '2. 비파괴검사결과서'의 '구간정보' 값을 사용한다.
                if mode in ("RT", "KOGAS"):
                    loc = r.get("구간정보", "") or r.get("구간", "")
                else:
                    loc = r.get("구간", "")
                if "Loc" in keys_attr: new_row["Loc"] = loc
                elif "Location" in keys_attr: new_row["Location"] = loc
                if mode in ("RT", "KOGAS"):
                    from site_apps.central.src.reports.ndt_history_import import rt_film_findings
                    new_row.update(rt_film_findings(r.get("구간정보", "")))
                
                # 결과 / Result / Acc / Rej
                if mode == "PAUT":
                    from site_apps.central.src.reports.ndt_history_import import paut_flaw_findings
                    findings = paut_flaw_findings(r.get("구간정보", ""), thickness)
                    if findings.get('Remarks') and new_row.get('Remarks'):
                        findings['Remarks'] = new_row['Remarks'] + '; ' + findings['Remarks']
                    new_row.update(findings)
                res = r.get("결과", "")
                if "Result" in keys_attr:
                    new_row["Result"] = res
                if res == "합격":
                    if "Acc" in keys_attr: new_row["Acc"] = "√"
                    if "Evaluation" in keys_attr: new_row["Evaluation"] = "Acc"
                elif res == "불합격":
                    if "Rej" in keys_attr: new_row["Rej"] = "X"
                    if "Evaluation" in keys_attr: new_row["Evaluation"] = "Rej"
                
                new_row["selected"] = True
                new_row["_status"] = "new"
                if "No" in keys_attr: new_row["No"] = str(len(data) + 1)
                
                data.append(new_row)
                if r.get('_source_id'):
                    existing_ids.add(r['_source_id'])
                
                row_vals = [new_row.get(k, "") for k in keys_attr]
                row_tags = []
                if new_row.get("Result") == "불합격" or new_row.get("Evaluation") == "Rej":
                    row_tags.append("fail")
                elif new_row.get("Result") == "합격" or new_row.get("Evaluation") == "Acc":
                    row_tags.append("pass")
                
                tree.insert("", "end", values=tuple(row_vals), tags=tuple(row_tags))
                idx_map.append(len(data) - 1)
                added_count += 1
                
            self.update_date_listbox(mode=mode)
            getattr(self, '_last_auto_material_sig', {}).pop(mode, None)
            if mode == 'PAUT':
                companies = {row.get('업체', '').strip() for row in ndt_list if row.get('업체', '').strip()}
                if len(companies) == 1 and hasattr(self, 'paut_contractor'):
                    self.paut_contractor.set(next(iter(companies)))
                    self.config['PAUT_CONTRACTOR'] = self.paut_contractor.get()
                    self.save_settings()
                self.populate_preview(data, switch_tab=False, mode=mode)
            elif mode in ('PT', 'MT', 'RT', 'KOGAS'):
                self.populate_preview(data, switch_tab=False, mode=mode)
            messagebox.showinfo(
                "완료",
                f"{ndt_list[0]['_source_site']} 작업 감독일보에서 {added_count}개의 {mode} NDT 결과를 성공적으로 불러왔습니다.",
                parent=self.root,
            )
            self.root.after_idle(self.root.focus_set)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "오류", f"일보 불러오기 중 오류가 발생했습니다:\n{e}",
                parent=self.root,
            )

    def extract_only(self, show_msg=True):
        """데이터만 추출하여 리스트와 미리보기에 반영 (PMI/RT 대응)"""
        self.save_settings() # [FIX] UI 설정(컬럼명 등)을 즉시 반영하여 리셋 방지
        # 현재 활성 탭에 따른 모드 결정
        try:
            main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in main_tab:
                # [FIX] 하부 탭 이름을 직접 확인하여 모드 결정
                sub_tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab_text else "RT"
            elif "PT" in main_tab: mode = "PT"
            elif "MT" in main_tab: mode = "MT"
            elif "PAUT" in main_tab: mode = "PAUT"
            else: mode = "PMI"
        except Exception as e:
            self.log(f"⚠️ 모드 판별 오류: {e}")
            mode = "PMI"

        if mode == "PT":
            target_file_str = self.pt_target_file_path.get()
        elif mode == "MT":
            target_file_str = self.mt_target_file_path.get()
        elif mode == "PAUT":
            target_file_str = self.paut_target_file_path.get()
        elif mode == "RT":
            target_file_str = self.rt_target_file_path.get()
        elif mode == "KOGAS":
            target_file_str = self.kogas_target_file_path.get()
        else:
            target_file_str = self.target_file_path.get()
            
        if not target_file_str:
            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일을 선택해주세요.")
            return False
            
        target_files = target_file_str.split('|') if '|' in target_file_str else [target_file_str]
        
        if not target_files: return False
        
        success_count = 0
        for i, tf in enumerate(target_files):
            if not tf.strip(): continue
            is_last = (i == len(target_files) - 1)
            if self._extract_single_file(mode, tf.strip(), show_msg=(show_msg and is_last)):
                success_count += 1
                
        return success_count > 0

    def _extract_single_file(self, mode, target_file, show_msg=True):
        self.log(f"🔍 {mode} 데이터 추출 시작: {os.path.basename(target_file)}")
        
        # [NEW] Extract date from filename
        fname = os.path.basename(target_file)
        date_match = re.search(r'(\d{4}[-._]\d{2}[-._]\d{2}|\d{2}[-._]\d{2}[-._]\d{2}|\d{8}|\d{6})', fname)
        extracted_date = date_match.group(0) if date_match else ""
        
        # Standardize formatting
        if extracted_date:
            clean_date = re.sub(r'[-._]', '', extracted_date)
            if len(clean_date) == 8:
                extracted_date = f"{clean_date[:4]}-{clean_date[4:6]}-{clean_date[6:]}"
            elif len(clean_date) == 6:
                extracted_date = f"20{clean_date[:2]}-{clean_date[2:4]}-{clean_date[4:]}"
        
        if not extracted_date:
            extracted_date = datetime.datetime.now().strftime("%Y-%m-%d")
            self.log(f"⚠️ 파일명에서 날짜를 찾지 못해 오늘 날짜({extracted_date})로 설정합니다.")
        else:
            self.log(f"📅 파일 날짜 인식 성공: {extracted_date}")

        def _is_date_like_dwg(value):
            """Reject timestamps that can be mistaken for a drawing number."""
            if value is None or pd.isna(value):
                return False
            if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
                return True
            text = str(value).strip()
            return bool(re.fullmatch(
                r'\d{4}[-./]\d{1,2}[-./]\d{1,2}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?',
                text,
            ))

        self.progress['value'] = 0
        all_extracted_data = []

        # [AUTO COLUMN ADDITION] 
        if mode == "PMI" and hasattr(self, 'element_filters') and self.element_filters:
            new_keys = []
            for f_item in self.element_filters:
                f_k = f_item['key'].get().strip()
                if f_k and f_k.upper() not in ["CR", "NI", "MO", "MN"] and f_k.capitalize() not in self.column_keys and f_k.upper() not in self.column_keys and f_k not in self.column_keys:
                    new_keys.append(f_k)
            
            if new_keys:
                self.column_keys.extend(new_keys)
                current_cols = list(self.preview_tree["columns"])
                current_display = list(self.preview_tree.cget("displaycolumns"))
                if current_display[0] == "#all": current_display = list(current_cols)

                # Save existing configurations to survive the tree["columns"] reset
                configs = {}
                for c in current_cols:
                    configs[c] = {
                        "text": self.preview_tree.heading(c).get("text", c),
                        "width": self.preview_tree.column(c).get("width", 80)
                    }

                for k in new_keys:
                    if k not in current_cols: current_cols.append(k)
                    if k not in current_display: current_display.append(k)
                
                self.preview_tree["columns"] = tuple(current_cols)
                self.preview_tree["displaycolumns"] = tuple(current_display)
                
                # Re-apply all configs + new keys setup
                for c in current_cols:
                    header_txt = configs.get(c, {}).get("text", c)
                    w = configs.get(c, {}).get("width", 80)
                    self.preview_tree.heading(c, text=header_txt, anchor='center', command=lambda _c=c: self.sort_by_column(_c, mode="PMI"))
                    self.preview_tree.column(c, width=w, anchor='center', stretch=False)
        
        def _find_col(df, keywords, exclude=None):
            cleaned_keywords = [re.sub(r'\s+', '', k).upper() for k in keywords]
            # 1. Exact match by keyword priority
            for k in cleaned_keywords:
                for col in df.columns:
                    c_clean = re.sub(r'\s+', '', str(col)).upper()
                    if exclude and any(re.sub(r'\s+', '', ex).upper() in c_clean for ex in exclude): continue
                    if k == c_clean: return col
            # 2. Substring match by keyword priority
            for k in cleaned_keywords:
                for col in df.columns:
                    c_clean = re.sub(r'\s+', '', str(col)).upper()
                    if exclude and any(re.sub(r'\s+', '', ex).upper() in c_clean for ex in exclude): continue
                    if k in c_clean:
                        if "NI" in keywords and ("UNIT" in c_clean or "LINE" in c_clean): continue
                        return col
            return None

        def _get_kw(k): 
            if hasattr(self, 'rt_extract_mappings') and k in self.rt_extract_mappings:
                return [x.strip().upper() for x in self.rt_extract_mappings[k].get().split(',')]
            return []

        try:
            target_input = self.sequence_filter.get().strip()
            target_no_list = [x.strip() for x in target_input.replace(',', ' ').split() if target_input and x.strip()] if target_input else []
            
            with pd.ExcelFile(target_file) as xls:
                for s_idx, sheet_name in enumerate(xls.sheet_names):
                    self.log(f"📄 시트 스캔: {sheet_name}")
                    try: temp_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=50)
                    except: continue
                    
                    # [지능형 시트 내 날짜 추출]
                    # 의뢰서의 L5, L6 셀(머지된 영역)을 직접 파싱하여 
                    # 파일명 날짜보다 우선하여 해당 시트의 기본 날짜(sheet_level_date)로 활용합니다.
                    sheet_level_date = extracted_date
                    try:
                        for r_pos in [4, 5]: # 0-indexed 4=5행, 5=6행
                            for c_pos in [11, 10, 12]: # 11=L열, 10=K열, 12=M열
                                if r_pos < len(temp_df) and c_pos < len(temp_df.columns):
                                    val = temp_df.iloc[r_pos, c_pos]
                                    if pd.notna(val):
                                        if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)) or hasattr(val, 'strftime'):
                                            sheet_level_date = val.strftime('%Y-%m-%d')
                                            break
                                        cleaned = str(val).strip()
                                        # YYYY-MM-DD 혹은 YYYY.MM.DD 등의 형태에서 날짜만 클리닝
                                        cleaned_num = re.sub(r'[^0-9]', '', cleaned)
                                        if len(cleaned_num) == 8:
                                            sheet_level_date = f"{cleaned_num[:4]}-{cleaned_num[4:6]}-{cleaned_num[6:]}"
                                            break
                                        elif len(cleaned_num) == 6:
                                            prefix = "20" if int(cleaned_num[:2]) < 50 else "19"
                                            sheet_level_date = f"{prefix}{cleaned_num[:2]}-{cleaned_num[2:4]}-{cleaned_num[4:]}"
                                            break
                            if sheet_level_date != extracted_date: break
                    except: pass
                    
                    if sheet_level_date != extracted_date:
                        self.log(f"   📅 의뢰서 헤더(L5/L6 셀)에서 날짜 파싱 성공: {sheet_level_date}")
                        
                    # [지능형 시트 내 도면번호 추출]
                    # 의뢰서의 C5, C6 셀(머지된 영역)을 직접 파싱하여
                    # 해당 시트의 기본 도면번호(sheet_level_dwg)로 활용합니다.
                    # PT/MT 모드는 의뢰서 고정 헤더 구조가 아니므로 건너뜁니다.
                    sheet_level_dwg = ""
                    if mode not in ["PT", "MT"]:
                        try:
                            # C5/C6 is a valid fixed drawing-number area only
                            # when a nearby cell explicitly labels it as such.
                            nearby_text = " ".join(
                                str(temp_df.iloc[r_pos, c_pos]).strip().upper()
                                for r_pos in range(3, min(7, len(temp_df)))
                                for c_pos in range(0, min(5, len(temp_df.columns)))
                                if pd.notna(temp_df.iloc[r_pos, c_pos])
                            )
                            has_dwg_label = any(
                                label in nearby_text
                                for label in ["DWG", "DRAWING", "도면번호", "도면 NO", "도면NO"]
                            )
                            if has_dwg_label:
                                for r_pos in [4, 5]: # 0-indexed 4=5행, 5=6행
                                    for c_pos in [2, 1, 3]: # 2=C열, 1=B열, 3=D열
                                        if r_pos < len(temp_df) and c_pos < len(temp_df.columns):
                                            val = temp_df.iloc[r_pos, c_pos]
                                            if pd.notna(val):
                                                if _is_date_like_dwg(val):
                                                    continue
                                                val_str = str(val).strip()
                                                if val_str and val_str.lower() != 'nan' and len(val_str) > 2:
                                                    if val_str.upper() in ["도면번호", "도면", "DWG", "DWG NO", "DWG.NO", "도면 NO", "도면 NO."]:
                                                        continue
                                                    val_upper = val_str.upper()
                                                    if any(k in val_upper for k in ["CO.", "LTD", "INSPECTION", "TESTING", "CORP", "INC", "SEOUL", "주식회사"]):
                                                        continue
                                                    if val_str.count(" ") > 2:
                                                        continue
                                                    sheet_level_dwg = val_str
                                                    break
                                    if sheet_level_dwg: break
                        except: pass

                    if sheet_level_dwg:
                        self.log(f"   📐 의뢰서 헤더(C5/C6 셀)에서 도면번호 파싱 성공: {sheet_level_dwg}")
                    
                    header_idx = None
                    is_fixed_range = False
                    # [NEW] KOGAS 고정 영역 처리 (A9:AB13 헤더, A14:AB25 데이터)
                    if mode == "KOGAS":
                        header_idx = 8 # Excel Row 9
                        is_fixed_range = True
                    else:
                        for i, row in temp_df.iterrows():
                            row_str = str(row.values).upper()
                            if mode == "RT":
                                if any(k in row_str for k in ["FILM", "DEFECT", "IQI", "순번", "용접부"]):
                                    header_idx = i; break
                            elif mode == "PT":
                                if (("LINE" in row_str or "ISO" in row_str or "DWG" in row_str) and ("JOINT" in row_str or "WELD" in row_str)):
                                    header_idx = i; break
                            else:
                                if ("CR" in row_str and "NI" in row_str) or ("CHROMIUM" in row_str):
                                    header_idx = i; break
                    
                    if header_idx is None: continue
                    
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_idx)
                    if df.empty: continue

                    if is_fixed_range:
                        # [KOGAS] 9~13행 헤더 병합 (df.columns + df.iloc[0:3])
                        new_cols = []
                        for col_idx in range(len(df.columns)):
                            parts = [str(df.columns[col_idx])]
                            for r_idx in range(0, 4): # 10, 11, 12, 13행
                                val = df.iloc[r_idx, col_idx]
                                if pd.notna(val) and "UNNAMED" not in str(val).upper():
                                    parts.append(str(val))
                            new_cols.append(" ".join(parts).upper())
                        df.columns = new_cols
                        # [KOGAS] 14~25행 데이터 추출 (인덱스 4~15)
                        df = df.iloc[4:16].reset_index(drop=True)
                    else:
                        # [STANDARD] 다중 헤더 병합 (최대 2행까지 지원)
                        row0_str = " ".join([str(x).upper() for x in df.iloc[0] if pd.notna(x)]) if len(df) > 0 else ""
                        row1_str = " ".join([str(x).upper() for x in df.iloc[1] if pd.notna(x)]) if len(df) > 1 else ""
                        
                        merge_count = 0
                        if any(k in row1_str for k in ["ORIGIN", "FACTOR", "WELDER1"]):
                            merge_count = 2
                        elif any(k in row0_str for k in ["ORIGIN", "FACTOR", "WELDER1"]):
                            merge_count = 1
                            
                        if merge_count > 0:
                            self.log(f"   ℹ️ [자동] 다중 헤더 병합 ({merge_count}행) (시트: {sheet_name})")
                            new_cols = []
                            for c_idx, col in enumerate(df.columns):
                                c_txt = str(col) if pd.notna(col) and "Unnamed" not in str(col) else ""
                                v_parts = []
                                for r in range(merge_count):
                                    v = df.iloc[r, c_idx]
                                    if pd.notna(v) and "Unnamed" not in str(v): v_parts.append(str(v).strip())
                                v_txt = " ".join(v_parts)
                                new_cols.append(f"{c_txt} {v_txt}".strip())
                            df.columns = new_cols
                            df = df.iloc[merge_count:].reset_index(drop=True)

                    # --- Column Identification for THIS sheet ---
                    col_date = None
                    if mode in ["RT", "KOGAS"]:
                        def _get_col(config_key, search_keywords):
                            c_idx_str = str(self.config.get(config_key, "")).strip()
                            if c_idx_str:
                                try:
                                    idx = self.col_to_num(c_idx_str)
                                    if 1 <= idx <= len(df.columns):
                                        candidate = df.columns[idx - 1]
                                        
                                        # [지능형 자동 매핑 우선 정책] 
                                        # 오직 사용자가 설정을 변경하지 않고 '기본값' 상태일 때만, 헤더 매칭이 틀렸을 경우 자동 매핑을 시도합니다.
                                        # 사용자가 명시적으로 컬럼을 직접 변경한 경우는 설정을 100% 신뢰하여 그대로 사용합니다.
                                        DEFAULT_COLUMNS = {
                                            "KOGAS_R_COL_NO": "1", "KOGAS_R_COL_DATE": "2", "KOGAS_R_COL_DWG": "3",
                                            "KOGAS_R_COL_JOINT": "4", "KOGAS_R_COL_LOC": "5", "KOGAS_R_COL_THK": "6",
                                            "KOGAS_R_COL_MAT": "7", "KOGAS_R_COL_SIZE": "8", "KOGAS_R_COL_RES": "9",
                                            "KOGAS_R_COL_WELDER": "10", "KOGAS_R_COL_REM": "11", "KOGAS_R_COL_D1": "12",
                                            "RT_COL_NO": "1", "RT_COL_DATE": "2", "RT_COL_DWG": "3", "RT_COL_JOINT": "4",
                                            "RT_COL_LOC": "5", "RT_COL_THK": "6", "RT_COL_MAT": "7", "RT_COL_SIZE": "0",
                                            "RT_COL_ACC": "8", "RT_COL_REJ": "9", "RT_COL_DEG": "10", "RT_COL_RES": "0",
                                        }
                                        is_default = False
                                        default_val = DEFAULT_COLUMNS.get(config_key)
                                        if default_val is not None:
                                            try:
                                                is_default = (self.col_to_num(c_idx_str) == self.col_to_num(default_val))
                                            except:
                                                is_default = (c_idx_str == default_val)
                                                
                                        if is_default:
                                            cand_clean = re.sub(r'\s+', '', str(candidate)).upper()
                                            cleaned_keywords = [re.sub(r'\s+', '', k).upper() for k in search_keywords]
                                            if search_keywords and not any(k in cand_clean for k in cleaned_keywords):
                                                found = _find_col(df, search_keywords)
                                                if found: return found
                                                
                                        return candidate
                                except: pass
                            return _find_col(df, search_keywords)

                        col_prefix = "KOGAS_R" if mode == "KOGAS" else "RT"
                        col_no = _get_col(f"{col_prefix}_COL_NO", _get_kw('No') or ["NO.", "NO", "SEQ", "ITEM", "순번"])
                        col_dwg = _get_col(f"{col_prefix}_COL_DWG", _get_kw('Dwg') or ["ISO", "DWG", "DRAWING", "도면"])
                        col_joint = _get_col(f"{col_prefix}_COL_JOINT", _get_kw('Joint') or ["JOINT", "WELD NO", "J/N", "FILM IDENT", "용접부"])
                        col_loc = _get_col(f"{col_prefix}_COL_LOC", _get_kw('Loc') or ["LOCATION", "POSITION", "FILM LOC"])
                        col_size = _get_col(f"{col_prefix}_COL_SIZE", _get_kw('Size') or ["SIZE", "DIA", "INCH", "구경", "관경"])
                        col_welder = _get_col(f"{col_prefix}_COL_WELDER", _get_kw('Welder') or ["WELDER", "W/N"])
                        col_remarks = _get_col(f"{col_prefix}_COL_REM", _get_kw('Remarks') or ["REMARKS", "REMARK", "비고"])
                        col_date = _get_col(f"{col_prefix}_COL_DATE", _get_kw('Date') or ["DATE", "검사일"])
                        col_t = _get_col(f"{col_prefix}_COL_THK", _get_kw('T') or ["T", "THICK", "THK", "두께"])
                        col_mat = _get_col(f"{col_prefix}_COL_MAT", _get_kw('Mat') or ["MAT", "MATERIAL"])
                        col_weld = _get_col(f"{col_prefix}_COL_WELD", _get_kw('Weld') or ["WELD", "TYPE"])
                        col_iqi = _get_col(f"{col_prefix}_COL_IQI", _get_kw('IQI') or ["IQI"])
                        col_sens = _get_col(f"{col_prefix}_COL_SENS", _get_kw('Sens') or ["SENS", "SENSITIVITY"])
                        col_den = _get_col(f"{col_prefix}_COL_DEN", _get_kw('Den') or ["DEN", "DENSITY"])
                        col_acc = _get_col(f"{col_prefix}_COL_ACC", _get_kw('Acc') or ["ACC", "합격"])
                        col_rej = _get_col(f"{col_prefix}_COL_REJ", _get_kw('Rej') or ["REJ", "불합격"])
                        col_deg = _get_col(f"{col_prefix}_COL_DEG", _get_kw('Deg') or ["DEG", "물성", "수정브랜드", "GRADE"])
                        col_result = _get_col(f"{col_prefix}_COL_RES", _get_kw('Result') or ["RESULT", "판정"])
                        
                        defect_cols = {}
                        for i in range(1, 16):
                            defect_cols[f"D{i}"] = _get_col(f"{col_prefix}_COL_D{i}", [f"D{i}", f"DEFECT{i}", chr(9311 + i), f"{i}"])
                        defect_cols = {k: v for k, v in defect_cols.items() if v}

                        # [NEW] KOGAS Fallback (B column for Joint, E column for Size) ONLY if not found
                        if mode == "KOGAS":
                            if col_joint is None and len(df.columns) > 1: col_joint = df.columns[1]
                            if col_size is None and len(df.columns) > 4: col_size = df.columns[4]
                    elif mode == "PT":
                        # PT requires an exact standalone sequence header.
                        # Do not treat Spool No / Request No / Film No as No.
                        # If absent, the row loop below assigns 1, 2, 3, ... .
                        col_no = find_exact_column(df, ["NO.", "NO", "SEQ", "ITEM", "순번"])
                        # Use the request date from each PT source row instead
                        # of the file name/header-level fallback date.
                        col_date = _find_col(df, ["TEST DATE", "REQUEST DATE", "DATE", "검사일자", "일자"])
                        col_report_no = _find_col(df, ["REPORT NO", "REPORT"])
                        col_dwg = _find_col(df, ["ISO", "LINE", "DWG", "DRAWING"], exclude=["JOINT", "WELD"]) 
                        col_joint = _find_col(df, ["JOINT NO", "JOINT NUMBER"], exclude=["ISO", "LINE", "ITEM"])
                        if not col_joint:
                            col_joint = _find_col(df, ["JOINT", "WELD"], exclude=["ISO", "LINE", "ITEM", "WPS", "REPORT", "DATE", "TYPE"])
                        col_mat = _find_col(df, ["MAT", "MATERIAL", "재질", "CLASS", "M'TL"])
                        col_size = _find_col(df, ["SIZE", "NPS", "DIA", "INCH", "관경", "ORIGIN", "DI ORIGIN"])
                        col_thk = _find_col(df, ["THK", "THICK", "SCH", "두께"])
                        col_welder = _find_col(df, ["WELDER", "ID", "용접사", "WELDER1"])
                        col_wtype = _find_col(df, ["WELD TYPE", "WELDTYPE", "W.TYPE", "TYPE"], exclude=["JOINT"])
                        col_result = _find_col(df, ["RESULT", "결과", "판정"])
                    # Placeholder for Loc to maintain structure
                        col_loc = None 
                    else:
                        col_cr = _find_col(df, ["CR", "CHROMIUM"]); col_ni = _find_col(df, ["NI", "NICKEL"])
                        col_mo = _find_col(df, ["MO", "MOLYBDENUM"]); col_mn = _find_col(df, ["MN", "MANGANESE"])
                        col_no = _find_col(df, ["NO.", "NO", "SEQ", "NUM", "POS", "ITEM"], exclude=["REPORT", "DWG", "DRAWING", "LINE", "WELD", "JOINT", "ORIGIN"])
                        col_date = _find_col(
                            df,
                            ["DATE/TIME", "DATETIME", "DATE", "TIME", "검사일", "검사일자", "측정시간", "일자"],
                        )
                        col_joint = _find_col(df, ["JOINT", "J/N", "JOINT NO", "PUNCH", "WELD NO"])
                        col_loc = _find_col(df, ["LOCATION", "TEST POSITION", "POINT", "AREA", "POSITION"])
                        col_dwg = _find_col(df, ["ISO", "DWG", "DRAWING", "LINE"])
                        col_report_no = _find_col(df, ["REPORT NO", "REPORT"])
                        col_grade_orig = _find_col(df, ["GRADE", "MATERIAL", "SPEC", "TYPE"])
                        
                        # [DYNAMIC] Scan for dynamically added filter elements (e.g., C, P, S)
                        custom_cols = {}
                        if hasattr(self, 'element_filters') and self.element_filters:
                            for f_item in self.element_filters:
                                f_key = f_item['key'].get().strip()
                                if f_key and f_key.upper() not in ["CR", "NI", "MO", "MN"]:
                                    c_found = _find_col(df, [f_key.upper()])
                                    if c_found is not None:
                                        custom_cols[f_key] = c_found

                    # [NEW] Forward-fill (Fill Down) for ISO and Joint
                    last_dwg = ""
                    last_joint = ""
                    last_date = ""

                    def clean_v(v):
                        if pd.isna(v): return ""
                        # 만약 날짜/시간(Timestamp) 형태라면 YYYY-MM-DD 형식으로 클리닝
                        if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)) or hasattr(v, 'strftime'):
                            try:
                                return v.strftime('%Y-%m-%d')
                            except: pass
                        if isinstance(v, str):
                            v_strip = v.strip()
                            # [NEW] Preserve leading zeros for purely numeric strings (e.g. "075")
                            if v_strip.startswith('0') and v_strip.isdigit() and len(v_strip) > 1:
                                return v_strip
                                
                        try:
                            if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace('.','').isdigit()):
                                fv = float(v)
                                if fv.is_integer(): return str(int(fv))
                                return str(fv)
                        except: pass
                        # 만약 문자열이고 날짜 형식인 경우 시간 정보(00:00:00) 제거
                        ret_str = str(v).strip()
                        if len(ret_str) >= 10:
                            if ret_str[:4].isdigit() and ret_str[5:7].isdigit() and ret_str[8:10].isdigit():
                                if ret_str[4] == '-' and ret_str[7] == '-':
                                    ret_str = ret_str[:10]
                                elif ret_str[4] == '.' and ret_str[7] == '.':
                                    ret_str = ret_str[:10]
                                elif ret_str[4] == '/' and ret_str[7] == '/':
                                    ret_str = ret_str[:10]
                        return ret_str

                    def process_welder_string(w_str):
                        if not w_str or str(w_str).lower() == 'nan': return ''
                        import re
                        parts = re.split(r'[/,]', str(w_str))
                        processed = []
                        for p in parts:
                            p = p.strip()
                            if len(p) > 3:
                                p = p[-3:]
                            if p and p not in processed:
                                processed.append(p)
                        return "/".join(processed)

                if mode == "KOGAS":
                    # ===== 가스공사 모드: 2개 행을 1개 Joint로 페어링 =====
                    num_rows = len(df)
                    r_idx = 0
                    while r_idx < num_rows:
                        row_top = df.iloc[r_idx]
                        row_bot = df.iloc[r_idx + 1] if (r_idx + 1 < num_rows) else None

                        # 1. 순번 (No)
                        v_raw_no = clean_v(row_top[col_no]) if col_no is not None else ""
                        if not v_raw_no and row_bot is not None and col_no is not None:
                            v_raw_no = clean_v(row_bot[col_no])
                        
                        # 만약 NO 컬럼이 지정되어 있는데 값이 완전히 비어있거나 'nan'인 경우, 데이터가 없는 행이므로 건너뜁니다.
                        if col_no is not None and (not v_raw_no or v_raw_no == 'nan'):
                            r_idx += 2
                            continue
                            
                        # NO 컬럼이 지정되어 있지 않은 경우에만 인덱스 기반으로 번호를 생성합니다.
                        if not v_raw_no or v_raw_no == 'nan':
                            v_raw_no = str((r_idx // 2) + 1)

                        if target_no_list and v_raw_no not in target_no_list:
                            r_idx += 2
                            continue

                        # 키워드 필터링
                        extract_key = self.rt_extract_keyword.get().strip().lower()
                        if extract_key:
                            top_vals = [clean_v(v).lower() for v in row_top.values if pd.notna(v)]
                            bot_vals = [clean_v(v).lower() for v in row_bot.values if pd.notna(v)] if row_bot is not None else []
                            row_str = ' '.join(top_vals + bot_vals)
                            if extract_key not in row_str:
                                r_idx += 2
                                continue

                        # 2. 도면번호 (Dwg) 및 Dwg_Sub
                        curr_dwg = sheet_level_dwg if sheet_level_dwg else (clean_v(row_top[col_dwg]) if col_dwg is not None else '')
                        if (not curr_dwg or curr_dwg == 'nan'):
                            curr_dwg = last_dwg if last_dwg else ''
                        if curr_dwg and curr_dwg != 'nan': last_dwg = curr_dwg

                        curr_dwg_sub = ''
                        if sheet_level_dwg:
                            curr_dwg_sub = sheet_level_dwg
                        elif row_bot is not None and col_dwg is not None:
                            curr_dwg_sub = str(row_bot[col_dwg]).strip()
                        if not curr_dwg_sub or curr_dwg_sub == 'nan':
                            curr_dwg_sub = curr_dwg

                        # 3. 공동 (Joint)
                        curr_joint = clean_v(row_top[col_joint]) if col_joint is not None else ''
                        if (not curr_joint or curr_joint == 'nan') and last_joint: curr_joint = last_joint
                        if curr_joint and curr_joint != 'nan': last_joint = curr_joint
                        elif not curr_joint or curr_joint == 'nan': curr_joint = v_raw_no

                        # 4. 일자 (Date)
                        curr_date = clean_v(row_top[col_date]) if col_date is not None else ''
                        if (not curr_date or curr_date.lower() == 'nan') and last_date: curr_date = last_date
                        if curr_date and curr_date.lower() != 'nan': last_date = curr_date
                        elif not curr_date or curr_date.lower() == 'nan': curr_date = sheet_level_date if sheet_level_date else ''

                        v_report = clean_v(row_top[col_report_no]) if col_report_no is not None else ''
                        if (not v_report or v_report == 'nan') and getattr(self, 'last_report', ''): v_report = self.last_report
                        if v_report and v_report != 'nan': self.last_report = v_report
                        elif not v_report or v_report == 'nan': v_report = ''

                        # 5. 규격 (Size) 및 촬영 매수 (Shots)
                        raw_size = str(row_top[col_size]).strip() if col_size is not None else ''
                        s_match = re.search(r'(\d+\.?\d*)', raw_size)
                        size_val = s_match.group(1) if s_match else raw_size
                        num_shots = self.calculate_rt_shots(size_val)

                        # 6. 용접사 (Welder) 및 Welder_Sub
                        raw_welder = clean_v(row_top[col_welder]) if col_welder is not None else ''
                        v_welder = process_welder_string(raw_welder)

                        raw_welder_sub = ''
                        if row_bot is not None and col_welder is not None:
                            raw_welder_sub = clean_v(row_bot[col_welder])
                        v_welder_sub = process_welder_string(raw_welder_sub)
                        
                        # [NEW] Combine both welders into one string if they are different and not already included
                        v_welder_parts = v_welder.split('/') if v_welder else []
                        if v_welder_sub and v_welder_sub != 'nan':
                            for sub_part in v_welder_sub.split('/'):
                                if sub_part and sub_part not in v_welder_parts:
                                    v_welder_parts.append(sub_part)
                        v_welder = "/".join(v_welder_parts)
                        
                        v_welder_sub = v_welder

                        # 7. 재질 (Mat) 및 Mat_Sub
                        raw_mat = clean_v(row_top[col_mat]) if col_mat is not None else ''
                        raw_mat_sub = ''
                        if row_bot is not None and col_mat is not None:
                            raw_mat_sub = clean_v(row_bot[col_mat])
                        if not raw_mat_sub or raw_mat_sub == 'nan':
                            raw_mat_sub = raw_mat

                        # 8. 두께 (T)
                        raw_t = clean_v(row_top[col_t]) if col_t is not None else ''
                        if raw_t:
                            if '*' in raw_t or 'x' in raw_t.lower():
                                t_match = re.search(r'[*xX]\s*(\d+\.?\d*)\s*[tT]?', raw_t)
                                if t_match:
                                    raw_t = t_match.group(1)
                            else:
                                t_suffix_match = re.search(r'^(\d+\.?\d*)\s*[tT]?$', raw_t)
                                if t_suffix_match:
                                    raw_t = t_suffix_match.group(1)

                        # 9. 위치 (Loc) - 촬영 위치 추출 추가!
                        v_loc = clean_v(row_top[col_loc]) if col_loc is not None else ''
                        if not v_loc or v_loc == 'nan':
                            v_loc = '-'

                        item_data = {
                            'No': v_raw_no, 'Date': curr_date,
                            'Dwg': curr_dwg, 'Dwg_Sub': curr_dwg_sub,
                            'Joint': curr_joint, 'Loc': v_loc,
                            'Acc': clean_v(row_top[col_acc]) if col_acc is not None else '',
                            'Rej': clean_v(row_top[col_rej]) if col_rej is not None else '',
                            'Deg': clean_v(row_top[col_deg]) if col_deg is not None else '',
                            'Welder': v_welder, 'Welder_Sub': v_welder_sub,
                            'Remarks': clean_v(row_top[col_remarks]) if col_remarks is not None else '',
                            'T': raw_t,
                            'Mat': raw_mat, 'Mat_Sub': raw_mat_sub,
                            'Weld': clean_v(row_top[col_weld]) if col_weld is not None else '',
                            'IQI': clean_v(row_top[col_iqi]) if col_iqi is not None else '',
                            'Sens': clean_v(row_top[col_sens]) if col_sens is not None else '',
                            'Den': clean_v(row_top[col_den]) if col_den is not None else '',
                            'Result': clean_v(row_top[col_result]) if col_result is not None else 'ACC',
                            'Size': size_val, 'num_shots': num_shots,
                            'selected': True, 'order_index': len(self.rt_extracted_data) + len(all_extracted_data),
                            '_src': {
                                'sheet': sheet_name,
                                'row': r_idx + 4 + header_idx + 2,
                                'col_result': list(df.columns).index(col_result) + 1 if col_result is not None else None,
                                'col_remarks': list(df.columns).index(col_remarks) + 1 if col_remarks is not None else None,
                                'col_acc': list(df.columns).index(col_acc) + 1 if col_acc is not None else None,
                                'col_rej': list(df.columns).index(col_rej) + 1 if col_rej is not None else None,
                                'col_defects': {k: list(df.columns).index(v) + 1 for k, v in defect_cols.items()}
                            }
                        }
                        # Defects
                        for i in range(1, 16):
                            key = f'D{i}'; src_col = defect_cols.get(key)
                            src_val = str(row_top[src_col]).strip() if src_col is not None else ''
                            item_data[key] = src_val if (src_val and src_val.lower() != 'nan') else ''

                        if num_shots > 1:
                            for shot_idx in range(1, num_shots + 1):
                                shot_item = item_data.copy()
                                shot_item['Loc'] = f'{shot_idx}-{(shot_idx % num_shots) + 1}'
                                for i in range(1, 16):
                                    key = f'D{i}'; c = defect_cols.get(key); val = str(row_top[c]).strip() if c is not None else ''
                                    shot_item[key] = val if (val and val.lower() not in ['nan', ''] and val.lower() in ['v', 'x', 'o', '1', '?']) else ''
                                
                                # [NEW] Clear repetitive Joint number for 2nd shot and beyond
                                if shot_idx > 1:
                                    shot_item['Joint'] = ''
                                    
                                all_extracted_data.append(shot_item)
                        else:
                            all_extracted_data.append(item_data)
                        r_idx += 2

                else:
                    # ===== 표준 모드 (RT, PT, PMI) =====
                    for _, row in df.iterrows():
                        v_raw_no = clean_v(row[col_no]) if col_no is not None else ""
                        raw_dwg = clean_v(row[col_dwg]) if col_dwg is not None else ""
                        if _is_date_like_dwg(raw_dwg):
                            raw_dwg = ""
                        raw_joint = clean_v(row[col_joint]) if col_joint is not None else ""
                        
                        # 행의 주요 데이터가 모두 비어있으면 빈 행으로 간주하여 건너뜀
                        if not v_raw_no and not raw_dwg and not raw_joint:
                            continue
                            
                        # 주요 데이터는 있는데 순번만 비어있으면 자동 부여
                        if not v_raw_no or v_raw_no == "nan":
                            v_raw_no = str(len(all_extracted_data) + 1)
                            
                        if target_no_list and v_raw_no not in target_no_list: continue

                        extract_key = self.rt_extract_keyword.get().strip().lower()
                        if extract_key:
                            row_vals = [clean_v(v).lower() for v in row.values if pd.notna(v)]
                            row_str = ' '.join(row_vals)
                            if extract_key not in row_str:
                                continue

                        curr_dwg = sheet_level_dwg if sheet_level_dwg else raw_dwg
                        if _is_date_like_dwg(curr_dwg):
                            curr_dwg = ''
                        if (not curr_dwg or curr_dwg == 'nan'):
                            curr_dwg = last_dwg if last_dwg else ''
                        if curr_dwg and curr_dwg != 'nan': last_dwg = curr_dwg

                        curr_joint = clean_v(row[col_joint]) if col_joint is not None else ''
                        if (not curr_joint or curr_joint == 'nan') and last_joint: curr_joint = last_joint
                        if curr_joint and curr_joint != 'nan': last_joint = curr_joint
                        elif not curr_joint or curr_joint == 'nan': curr_joint = v_raw_no

                        curr_date = clean_v(row[col_date]) if col_date is not None else ''
                        if (not curr_date or curr_date.lower() == 'nan') and last_date: curr_date = last_date
                        if curr_date and curr_date.lower() != 'nan': last_date = curr_date
                        elif not curr_date or curr_date.lower() == 'nan': curr_date = sheet_level_date if sheet_level_date else ''

                        v_report = clean_v(row[col_report_no]) if col_report_no is not None else ''
                        if (not v_report or v_report == 'nan') and getattr(self, 'last_report', ''): v_report = self.last_report
                        if v_report and v_report != 'nan': self.last_report = v_report
                        elif not v_report or v_report == 'nan': v_report = ''

                        if mode == 'RT':
                            raw_size = str(row[col_size]).strip() if col_size is not None else ''
                            s_match = re.search(r'(\d+\.?\d*)', raw_size)
                            size_val = s_match.group(1) if s_match else raw_size

                            num_shots = self.calculate_rt_shots(size_val)

                            raw_welder = clean_v(row[col_welder]) if col_welder is not None else ''
                            v_welder = process_welder_string(raw_welder)

                            v_date = curr_date

                            item_data = {
                                'No': v_raw_no, 'Date': v_date, 'Report No': v_report,
                                'Dwg': curr_dwg, 'Joint': curr_joint, 'Loc': '-',
                                'Acc': clean_v(row[col_acc]) if col_acc is not None else '',
                                'Rej': clean_v(row[col_rej]) if col_rej is not None else '',
                                'Deg': clean_v(row[col_deg]) if col_deg is not None else '',
                                'Welder': v_welder,
                                'Remarks': clean_v(row[col_remarks]) if col_remarks is not None else '',
                                'T': clean_v(row[col_t]) if col_t is not None else '',
                                'Mat': clean_v(row[col_mat]) if col_mat is not None else '',
                                'Weld': clean_v(row[col_weld]) if col_weld is not None else '',
                                'IQI': clean_v(row[col_iqi]) if col_iqi is not None else '',
                                'Sens': clean_v(row[col_sens]) if col_sens is not None else '',
                                'Den': clean_v(row[col_den]) if col_den is not None else '',
                                'Result': clean_v(row[col_result]) if col_result is not None else 'ACC',
                                'Size': size_val, 'num_shots': num_shots,
                                'selected': True, 'order_index': len(self.rt_extracted_data) + len(all_extracted_data),
                                '_src': {
                                    'sheet': sheet_name,
                                    'row': _ + header_idx + 2,
                                    'col_result': list(df.columns).index(col_result) + 1 if col_result is not None else None,
                                    'col_remarks': list(df.columns).index(col_remarks) + 1 if col_remarks is not None else None,
                                    'col_acc': list(df.columns).index(col_acc) + 1 if col_acc is not None else None,
                                    'col_rej': list(df.columns).index(col_rej) + 1 if col_rej is not None else None,
                                    'col_defects': {k: list(df.columns).index(v) + 1 for k, v in defect_cols.items()}
                                }
                            }
                            for i in range(1, 16):
                                key = f'D{i}'; src_col = defect_cols.get(key)
                                src_val = str(row[src_col]).strip() if src_col is not None else ''
                                item_data[key] = src_val if src_val else ('1' if i <= num_shots else '')

                            if num_shots > 1:
                                for shot_idx in range(1, num_shots + 1):
                                    shot_item = item_data.copy()
                                    shot_item['Loc'] = f'{shot_idx}-{(shot_idx % num_shots) + 1}'
                                    for i in range(1, 16):
                                        key = f'D{i}'; c = defect_cols.get(key); val = str(row[c]).strip() if c is not None else ''
                                        shot_item[key] = '?' if val and val.lower() in ['v', 'x', 'o', '1', '?'] else ''
                                        
                                    # [NEW] Clear repetitive Joint number for 2nd shot and beyond
                                    if shot_idx > 1:
                                        shot_item['Joint'] = ''
                                        
                                    all_extracted_data.append(shot_item)
                            else:
                                all_extracted_data.append(item_data)
                        elif mode == "PT":
                            # 헤더 행 방어: Dwg/Joint 값이 컬럼 키워드이면 스킵
                            _dwg_up = curr_dwg.upper().strip()
                            _jnt_up = curr_joint.upper().strip()
                            _header_dwg = {"ISO", "LINE", "DWG", "DWG NO", "DWG.", "DRAWING", "ISO/DWG", "ISO NO"}
                            _header_jnt = {"JOINT", "JOINT NO", "JOINT NO.", "JOINT NUMBER", "WELD", "WELD NO"}
                            if _dwg_up in _header_dwg and _jnt_up in _header_jnt:
                                continue
                            res_str = str(row[col_result]).upper() if col_result is not None else "ACC"
                            if any(k in res_str for k in ["ACC", "OK", "ACCEPT", "합격"]):
                                item_row = {
                                    'No': v_raw_no, 'Date': curr_date, 'Report No': v_report, 'Dwg': curr_dwg, 'Joint': self.force_two_digit(curr_joint),
                                    'NPS': str(row[col_size]).strip() if col_size is not None else "", 'Thk.': convert_sch_to_thk(str(row[col_size]).strip() if col_size is not None else "", str(row[col_thk]).strip() if col_thk is not None else ""),
                                    'Material': self.fix_material_name(row[col_mat]) if col_mat is not None else "", 'Welder': str(row[col_welder]).strip() if col_welder is not None else "",
                                    'WType': str(row[col_wtype]).strip() if col_wtype is not None else "", 'Result': "Acc", 'selected': True,
                                    'order_index': len(self.pt_extracted_data) + len(all_extracted_data)
                                }
                                for k in self.pt_column_keys:
                                    if k not in item_row and k != "selected": item_row[k] = ""
                                all_extracted_data.append(item_row)
                        else:
                            # PMI-specific Row Data
                            v_cr = self.to_float(row[col_cr])
                            if v_cr > 0 or (v_raw_no != "" and v_raw_no != "nan"):
                                v_ni = self.to_float(row[col_ni])
                                v_mo = self.to_float(row[col_mo]) if col_mo is not None else 0.0
                                v_mn = self.to_float(row[col_mn]) if col_mn is not None else 0.0
                                orig_grade = str(row[col_grade_orig]).strip() if col_grade_orig is not None else ""
                                
                                final_grade = orig_grade
                                if self.auto_verify.get():
                                    detected = self.check_material_grade({'Cr': v_cr, 'Ni': v_ni, 'Mo': v_mo, 'Mn': v_mn})
                                    if detected: final_grade = detected
                                if not final_grade or final_grade == "nan":
                                    final_grade = "SS316" if v_mo >= 1.5 else "SS304"

                                pmi_item = {
                                    'No': v_raw_no, 
                                    'Joint': curr_joint,
                                    'Loc': str(row[col_loc]).strip() if col_loc is not None else "",
                                    'Cr': v_cr, 'Ni': v_ni, 'Mo': v_mo, 'Mn': v_mn,
                                    'Grade': final_grade, 'Dwg': curr_dwg,
                                    'Date': curr_date or sheet_level_date,
                                    'Report No': v_report,
                                    'selected': True,
                                    'order_index': len(self.extracted_data) + len(all_extracted_data)
                                }
                                # Extract dynamically found elements
                                for c_key, c_idx in custom_cols.items():
                                    pmi_item[c_key] = self.to_float(row[c_idx])
                                # [NEW] 기존 수동 추가 컬럼 보존
                                for k in self.column_keys:
                                    if k not in pmi_item and k != "selected":
                                        pmi_item[k] = ""
                                all_extracted_data.append(pmi_item)
                self.progress['value'] = ((s_idx + 1) / len(xls.sheet_names)) * 50

            if not all_extracted_data:
                messagebox.showerror("오류", "추출된 데이터가 없습니다.")
                return False

            # Extraction Mode Filter (PMI Only)
            if mode == "PMI":
                ext_mode = self.extraction_mode.get()
                if ext_mode != "전체":
                    original_count = len(all_extracted_data)
                    if ext_mode == "SS304 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS304"]
                    elif ext_mode == "SS316 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS316"]
                    elif ext_mode == "DUPLEX 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "DUPLEX"]
                    elif ext_mode == "SS310 만": all_extracted_data = [d for d in all_extracted_data if d['Grade'] == "SS310"]
                    elif ext_mode == "미분류(기타) 만":
                        known_grades = ["SS304", "SS316", "DUPLEX", "SS310"]
                        all_extracted_data = [d for d in all_extracted_data if d['Grade'] not in known_grades]
                    
                    self.log(f"🔍 필터링 ({ext_mode}): {original_count}개 -> {len(all_extracted_data)}개")
                    if not all_extracted_data:
                        messagebox.showinfo("알림", f"'{ext_mode}'에 해당하는 데이터가 없습니다.")
                        return False

            # [NEW] 원소 함량 필터링 적용 (PMI 전용)
            if mode == "PMI" and hasattr(self, 'element_filters') and self.element_filters:
                for f_item in self.element_filters:
                    f_key = f_item['key'].get().strip()
                    if not f_key: continue
                    f_min = self.to_float(f_item['min'].get())
                    f_max = self.to_float(f_item['max'].get())
                    
                    if f_min > 0 or f_max > 0:
                        original_count = len(all_extracted_data)
                        all_extracted_data = [d for d in all_extracted_data if (f_min <= d.get(f_key, 0.0) <= (f_max if f_max > 0 else 999.9))]
                        if len(all_extracted_data) != original_count:
                             self.log(f"🔍 원소 필터 적용 ({f_key}: {f_min}~{f_max}): {original_count} -> {len(all_extracted_data)}건 남음")

            if mode == "PT":
                # [NEW] [Refinement] Duplicate removal for PT (ISO + Joint)
                seen = set()
                unique_data = []
                for d in all_extracted_data:
                    key = (str(d['Dwg']).strip(), str(d['Joint']).strip())
                    if key not in seen:
                        seen.add(key)
                        unique_data.append(d)
                all_extracted_data = unique_data

            if target_no_list:
                all_extracted_data.sort(key=lambda x: target_no_list.index(str(x['No'])) if str(x['No']) in target_no_list else 999999)

            # [CHANGE] Overwrite -> Accumulate + Auto Sort
            getattr(self, '_last_auto_material_sig', {}).pop(mode, None)
            if mode == "RT":
                self.rt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("RT")
                self._update_rt_preview_columns(mode="RT") # [FIX] 표시 이름 보존을 위해 호출
                self.rt_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="RT") 
                total_count = len(self.rt_extracted_data)
            elif mode == "KOGAS":
                self.kogas_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("KOGAS")
                self._update_rt_preview_columns(mode="KOGAS") # [FIX] 표시 이름 보존을 위해 호출
                self.kogas_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="KOGAS")
                total_count = len(self.kogas_extracted_data)
            elif mode == "PT":
                self.pt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("PT")
                self.pt_sort_col = ""
                self.mt_sort_col = ""
                self.sort_by_column("Dwg", mode="PT") 
                total_count = len(self.pt_extracted_data)
            else:
                self.extracted_data.extend(all_extracted_data)
                self.pmi_sort_col = "" # [NEW] Force Ascending
                self.sort_by_column("Dwg", mode="PMI")
                # Refresh after sorting/populating so the visible PMI sidebar is
                # updated last and cannot be left with its initial empty state.
                self.update_date_listbox("PMI")
                # Run once more through Tk's event loop. This guarantees that the
                # visible widget is updated after all preview redraw callbacks.
                self.root.after_idle(lambda: self.update_date_listbox("PMI"))
                total_count = len(self.extracted_data)
            
            self.progress['value'] = 100
            if show_msg:
                self.log(f"✅ {mode} 데이터 누적 완료 (현재 총 {total_count} 건)")
                messagebox.showinfo("완료", f"데이터가 추가되었습니다.\n현재 목록에는 총 {total_count}건의 데이터가 있습니다.")
            return True
        except Exception as e:
            self.log(f"❌ {mode} 추출 오류: {e}")
            traceback.print_exc()
            return False

    def sync_results_to_request(self):
        """입력된 결과를 원본 의뢰서 파일로 역전송 (RT/KOGAS 대응)"""
        try:
            main_tab = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in main_tab:
                sub_tab = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab else "RT"
            else:
                messagebox.showwarning("경고", "현재 RT 모드에서만 지원됩니다.")
                return

            target_file = self.rt_target_file_path.get() if mode == "RT" else self.kogas_target_file_path.get()
            if not target_file or not os.path.exists(target_file):
                messagebox.showwarning("파일 오류", "원본 의뢰서 파일을 찾을 수 없습니다.")
                return

            data = self.rt_extracted_data if mode == "RT" else self.kogas_extracted_data
            if not data:
                messagebox.showwarning("데이터 오류", "반영할 데이터가 없습니다.")
                return

            if not messagebox.askyesno("확인", f"현재 입력된 결과를 원본 의뢰서({os.path.basename(target_file)})에 반영하시겠습니까?"):
                return

            self.log(f"🔄 의뢰서 결과 반영 시작: {os.path.basename(target_file)}")
            
            import openpyxl
            wb = openpyxl.load_workbook(target_file)
            count = 0
            
            # Group items by source row to handle multi-shot joints
            groups = {}
            for item in data:
                if not item.get('date_filtered', True): continue # [NEW] Apply Date filter
                if not item.get('_src'): continue
                src_key = (item['_src']['sheet'], item['_src']['row'])
                if src_key not in groups: groups[src_key] = []
                groups[src_key].append(item)

            for (sheet_name, row_idx), items in groups.items():
                if sheet_name not in wb.sheetnames: continue
                ws = wb[sheet_name]
                src = items[0]['_src']
                
                # Determine final result (If any item is Rej, result is Rej)
                final_res = "ACC"
                all_remarks = []
                
                if mode == "KOGAS":
                    for it in items:
                        res = str(it.get('Result', '')).upper()
                        if any(x in res for x in ["REJ", "NG", "불합격"]): final_res = "REJ"
                        rem = str(it.get('Remarks', '')).strip()
                        if rem and rem not in all_remarks: all_remarks.append(rem)
                    
                    # KOGAS 비고란은 컬럼 11 (K열)
                    ws.cell(row=row_idx, column=11, value=", ".join(all_remarks) if all_remarks else None)
                    
                    # NDT 의뢰방식 확인 (13:RT, 14:UT, 15:PT)
                    rt_val = ws.cell(row=row_idx, column=13).value
                    ut_val = ws.cell(row=row_idx, column=14).value
                    pt_val = ws.cell(row=row_idx, column=15).value
                    
                    def is_marked(v):
                        if v is None: return False
                        cleaned = str(v).strip().upper()
                        return cleaned in ["O", "V", "X", "●", "합격", "1", "TRUE"]
                    
                    rt_req = is_marked(rt_val)
                    ut_req = is_marked(ut_val)
                    pt_req = is_marked(pt_val)
                    
                    has_rt_films = False
                    for it in items:
                        loc_str = str(it.get('Loc', '')).strip()
                        if re.match(r'^\d+', loc_str):
                            has_rt_films = True
                            break
                    
                    defect_map = {
                        "D1": "C", "D2": "IP", "D3": "LF", "D4": "S", "D5": "P",
                        "D6": "UC", "D7": "RUC", "D8": "BT", "D9": "TI", "D10": "CP",
                        "D11": "RC", "D12": "Mis", "D13": "EP", "D14": "SD", "D15": "Oth"
                    }
                    
                    def get_item_val(it):
                        grade = str(it.get('Deg', '')).strip()
                        if grade.lower() == 'nan':
                            grade = ''
                        checked = []
                        for d_key, abbrev in defect_map.items():
                            val = str(it.get(d_key, '')).strip()
                            if val in ["√", "1", "v", "V", "o", "O", "●", "?"]:
                                checked.append(abbrev)
                        if checked:
                            return f"{grade if grade else '4'}/{','.join(checked)}"
                        else:
                            if any(x in str(it.get('Result', '')).upper() for x in ["ACC", "OK"]):
                                return "1"
                            return grade if grade else "4"

                    # RT (Q~W: 17~23) 반영
                    if rt_req or has_rt_films:
                        for col in range(17, 24):
                            ws.cell(row=row_idx, column=col, value=None)
                        for it in items:
                            loc_str = str(it.get('Loc', '')).strip()
                            film_num = None
                            match = re.match(r'^(\d+)', loc_str)
                            if match:
                                film_num = int(match.group(1))
                            if film_num and 1 <= film_num <= 7:
                                val = get_item_val(it)
                                ws.cell(row=row_idx, column=16 + film_num, value=val)
                                
                    # UT/UP (X: 24) 반영
                    if ut_req:
                        it = items[0] if items else {}
                        val = get_item_val(it)
                        ws.cell(row=row_idx, column=24, value=val)
                        
                    # PT (Y: 25) 반영
                    if pt_req:
                        it = items[0] if items else {}
                        val = get_item_val(it)
                        ws.cell(row=row_idx, column=25, value=val)
                else:
                    has_defects = {k: False for k in src.get('col_defects', {}).keys()}
                    
                    for it in items:
                        res = str(it.get('Result', '')).upper()
                        if any(x in res for x in ["REJ", "NG", "불합격"]): final_res = "REJ"
                        
                        rem = str(it.get('Remarks', '')).strip()
                        if rem and rem not in all_remarks: all_remarks.append(rem)
                        
                        if src.get('col_defects'):
                            for d_key, col_idx in src['col_defects'].items():
                                if it.get(d_key) in ["√", "1", "v", "V", "o", "O"]: has_defects[d_key] = True

                    # Write back to Excel
                    if src.get('col_result'):
                        ws.cell(row=row_idx, column=src['col_result'], value=final_res)
                    
                    if src.get('col_remarks') and all_remarks:
                        ws.cell(row=row_idx, column=src['col_remarks'], value=", ".join(all_remarks))
                    
                    if src.get('col_defects'):
                        for d_key, col_idx in src['col_defects'].items():
                            if has_defects[d_key]:
                                ws.cell(row=row_idx, column=col_idx, value="√")
                
                count += 1

            wb.save(target_file)
            self.log(f"✅ 총 {count}개 행의 결과를 의뢰서에 성공적으로 반영했습니다.")
            messagebox.showinfo("완료", f"{count}개 행의 결과가 의뢰서에 반영되었습니다.")

        except PermissionError:
            msg = f"파일이 Excel에서 열려 있어 저장할 수 없습니다.\n\n'{os.path.basename(target_file)}' 파일을 Excel에서 닫은 후 다시 시도해 주세요."
            self.log(f"❌ 의뢰서 결과 반영 실패: 파일이 열려 있음 (PermissionError)")
            messagebox.showerror("파일 접근 오류", msg)
        except Exception as e:
            self.log(f"❌ 의뢰서 결과 반영 실패: {e}")
            messagebox.showerror("오류", f"의뢰서 결과 반영 중 오류가 발생했습니다: {e}")

    def duplicate_selected_rows(self):
        """선택된 행들을 복제하여 데이터 리스트에 삽입"""
        selected = self.preview_tree.selection()
        if not selected: return
        
        # 시각적 순서대로 정렬 (위에서부터 복제)
        selected = sorted(list(selected), key=lambda x: self.preview_tree.index(x))
        
        new_items = []
        insert_pos = 0
        for item_id in selected:
            view_idx = self.preview_tree.index(item_id)
            actual_idx = self.item_idx_map[view_idx]
            new_item = self.extracted_data[actual_idx].copy()
            # No.는 중복 방지를 위해 비우거나 새로 부여 (여기서는 유지 후 나중에 rename 가능)
            new_items.append(new_item)
            insert_pos = actual_idx
            
        for item in reversed(new_items):
            self.extracted_data.insert(insert_pos + 1, item)
             
        self.populate_preview(self.extracted_data, switch_tab=False)
        self.log(f"👯 {len(new_items)}개 행 복제 완료")

    def save_column_widths(self):
        """현재 트리뷰의 컬럼 너비를 설정 파일에 저장"""
        widths = {}
        for col in self.preview_tree["columns"]:
            widths[col] = self.preview_tree.column(col, "width")
        
        self.config["PMI_COL_WIDTHS"] = widths
        self.save_settings()
        self.log("📏 컬럼 너비 설정 저장 완료")

    def apply_date_filter(self):
        """날짜 리스트박스 선택 상태를 데이터 모델에 반영"""
        selected_dates = []
        for i in range(self.date_listbox.size()):
            val = self.date_listbox.get(i)
            if val.startswith("[v]"):
                selected_dates.append(val.replace("[v] ", "").strip())
        
        if not selected_dates:
            messagebox.showwarning("필터 오류", "최소 하나 이상의 날짜를 선택해주세요.")
            return

        for item in self.extracted_data:
            item['date_filtered'] = (item.get('Date', '') in selected_dates)
        
        self.populate_preview(self.extracted_data, switch_tab=False)
        self.log(f"📅 날짜 필터 적용 완료: {len(selected_dates)}개 날짜 선택됨")

    def show_report_issue_history(self, mode=None):
        from site_apps.central.src.reports.report_issue_history import show_history

        def refresh_numbers():
            if not mode:
                return
            mode_info = self._get_mode_info(mode)
            if not mode_info:
                return
            data = mode_info[2]
            # Imported issue numbers are derived from active issue history.
            # Clear only source-linked rows so manually entered values survive.
            for item in data:
                if item.get('_source_id'):
                    item.pop('Report No', None)
            self.populate_preview(data, switch_tab=False, mode=mode)

        show_history(self.root, on_change=refresh_numbers)

    def _report_save_path(self, directory, extension, mode=None, choose_directory=True):
        from site_apps.central.src.reports.report_issue_history import report_output_path
        if choose_directory:
            mode = mode or getattr(self, 'current_mode', 'PMI')
            key = f'{mode}_LAST_SAVE_DIR'
            remembered = str(self.config.get(key, '')).strip()
            initial = remembered if remembered and os.path.isdir(remembered) else directory
            directory = filedialog.askdirectory(parent=self.root, title=f'{mode} 보고서 저장 폴더 선택',
                                                initialdir=initial, mustexist=True)
            if not directory:
                self.log(f'{mode} 보고서 저장이 취소되었습니다.')
                return None
            self.config[key] = directory
            self.save_settings()
        return report_output_path(directory, self.gapji_report_no.get(), extension,
                                  fallback='Report_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))

    def _record_report_issue(self, mode, workbook, save_path, items):
        from site_apps.central.src.reports.report_issue_history import record_issue
        try:
            # Read the actual saved cover value rather than a row's imported number.
            cell = {'RT': 'C5', 'KOGAS': 'C5', 'PT': 'P6', 'MT': 'S6', 'PAUT': 'AA5'}.get(mode)
            if cell is None:
                cell = self.config.get('GAPJI_REPORT_NO_CELL', 'B10')
            report_no = workbook.worksheets[0][cell].value
            if not messagebox.askyesno(
                '정식 발행 등록',
                f'생성한 {mode} 성적서를 정식 발행 이력에 등록할까요?\n\n'
                f'보고서 번호: {report_no or "(번호 미입력)"}\n'
                '아니오를 선택하면 파일만 생성되고 일보 항목은 미발행 상태로 유지됩니다.',
                parent=self.root,
            ):
                self.log(f"파일만 생성: {report_no or '(번호 미입력)'} · 발행 이력 미등록")
                return False
            record_issue(mode, report_no, save_path, items)
            if mode == 'PAUT':
                self.populate_preview(self.paut_extracted_data, switch_tab=False, mode=mode)
            elif mode in ('MT', 'RT', 'KOGAS'):
                self.populate_preview(self._get_mode_info(mode)[2], switch_tab=False, mode=mode)
            self.log(f"발행 이력 저장: {report_no or '(번호 미입력)'} · {len(items)}건")
            return True
        except Exception as exc:
            self.log(f"발행 이력 저장 실패: {exc}")
            messagebox.showwarning('발행 이력 저장 실패',
                f'보고서 파일은 생성되었으나 발행 이력을 저장하지 못했습니다.\n{save_path}\n\n{exc}')
            return False

    def _create_top_action_bar(self, parent, mode):
        """데이터 미리보기 상단에 모던 플랫 액션 툴바 생성"""
        action_bar = tk.Frame(parent, background="#ffffff", padx=8, pady=6)
        action_bar.pack(fill='x')

        # 1. 성적서 생성 / 시작 (파란색 하이라이트)
        if mode == "PAUT":
            cmd_run = self._generate_paut_report
            text_run = "✨ 성적서 생성"
        else:
            cmd_run = self.run_process
            text_run = "✨ 성적서 생성"

        btn_run = tk.Button(
            action_bar, text=text_run, command=cmd_run,
            font=("Malgun Gothic", 10, "bold"), bg="#2563eb", fg="#ffffff",
            activebackground="#1d4ed8", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=5, bd=0
        )
        btn_run.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_run.bind("<Enter>", lambda e, b=btn_run: b.configure(bg="#1d4ed8"))
        btn_run.bind("<Leave>", lambda e, b=btn_run: b.configure(bg="#2563eb"))

        # 2. 데이터 추출
        if mode == "PAUT":
            cmd_ext = self._extract_paut_data
        else:
            cmd_ext = self.extract_only

        btn_extract = tk.Button(
            action_bar, text="📝 데이터 추출", command=cmd_ext,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        btn_extract.pack(side='left', fill='x', expand=True, padx=(0, 6))
        btn_extract.bind("<Enter>", lambda e, b=btn_extract: b.configure(bg="#e2e8f0"))
        btn_extract.bind("<Leave>", lambda e, b=btn_extract: b.configure(bg="#f8fafc"))

        # 3. 일보 NDT 불러오기
        btn_daily = tk.Button(
            action_bar, text="📋 일보 NDT 불러오기", command=self.load_daily_work_history,
            font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
            activebackground="#e2e8f0", activeforeground="#1e293b",
            relief="solid", bd=1, cursor="hand2", padx=10, pady=5
        )
        has_extra = mode in ("RT", "PAUT")
        btn_daily.pack(side='left', fill='x', expand=True, padx=(0, 6) if has_extra else 0)
        btn_daily.bind("<Enter>", lambda e, b=btn_daily: b.configure(bg="#e2e8f0"))
        btn_daily.bind("<Leave>", lambda e, b=btn_daily: b.configure(bg="#f8fafc"))

        # RT 모드 추가 액션: 의뢰서 결과 반영
        if mode == "RT":
            btn_sync = tk.Button(
                action_bar, text="🔄 의뢰서에 결과 반영", command=self.sync_results_to_request,
                font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
                activebackground="#e2e8f0", activeforeground="#1e293b",
                relief="solid", bd=1, cursor="hand2", padx=8, pady=5
            )
            btn_sync.pack(side='left', fill='x', expand=True)
            btn_sync.bind("<Enter>", lambda e, b=btn_sync: b.configure(bg="#e2e8f0"))
            btn_sync.bind("<Leave>", lambda e, b=btn_sync: b.configure(bg="#f8fafc"))

        # PAUT 모드 추가 액션: 일괄 판정
        elif mode == "PAUT":
            btn_eval = tk.Button(
                action_bar, text="⚡ 일괄 판정", command=self._run_batch_paut_eval,
                font=("Malgun Gothic", 9, "bold"), bg="#f8fafc", fg="#334155",
                activebackground="#e2e8f0", activeforeground="#1e293b",
                relief="solid", bd=1, cursor="hand2", padx=8, pady=5
            )
            btn_eval.pack(side='left', fill='x', expand=True)
            btn_eval.bind("<Enter>", lambda e, b=btn_eval: b.configure(bg="#e2e8f0"))
            btn_eval.bind("<Leave>", lambda e, b=btn_eval: b.configure(bg="#f8fafc"))

        # 구분선
        tk.Frame(parent, height=1, background="#e2e8f0").pack(fill='x', pady=(0, 4))
        return action_bar

    def run_process(self):
        self.save_settings() # [FIX] Ensure UI settings are saved to config before generating
        
        # [NEW] Ensure config values are correct types for comparison (prevent str vs int errors)
        for k in list(self.config.keys()):
            if k.endswith(('_ROW', '_IDX', '_SIZE')) or any(x in k for x in ['START', 'END', 'PAGE']):
                try: self.config[k] = int(float(self.config[k]))
                except: pass
            elif any(x in k for x in ['MARGIN', 'SCALE', 'RATIO', 'POS', 'HEIGHT', 'WIDTH']):
                try: self.config[k] = float(self.config[k])
                except: pass

        # 현재 활성 탭에 따른 모드 결정
        try:
            tab_text = self.mode_notebook.tab(self.mode_notebook.select(), "text")
            if "RT" in tab_text:
                sub_tab = self.rt_preview_nb.tab(self.rt_preview_nb.select(), "text")
                mode = "KOGAS" if "가스공사" in sub_tab else "RT"
            elif "PT" in tab_text: mode = "PT"
            elif "MT" in tab_text: mode = "MT"
            elif "PAUT" in tab_text: mode = "PAUT"
            else: mode = "PMI"
        except: mode = "PMI"

        if mode == "RT":
            target_file = self.rt_target_file_path.get()
            template_path = self.rt_template_file_path.get()
            data = self.rt_extracted_data
        elif mode == "KOGAS":
            target_file = self.kogas_target_file_path.get()
            template_path = self.kogas_template_file_path.get()
            data = self.kogas_extracted_data
        elif mode == "PT":
            target_file = self.pt_target_file_path.get()
            template_path = self.pt_template_file_path.get()
            data = self.pt_extracted_data
        elif mode == "MT":
            target_file = self.mt_target_file_path.get()
            template_path = self.mt_template_file_path.get()
            data = self.mt_extracted_data
        else:
            target_file = self.target_file_path.get()
            template_path = self.template_file_path.get()
            data = self.extracted_data
        
        if not template_path:
            messagebox.showwarning("파일 미선택", f"{mode} 양식(Template) 파일을 선택해주세요.")
            return
            
        if not target_file and not data:
            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일(Excel)을 선택하거나, 저장된 데이터를 불러와주세요.")
            return

        # [NEW] 템플릿 파일 존재 여부 및 기본 구조 확인
        if not os.path.exists(template_path):
            messagebox.showerror("오류", f"템플릿 파일을 찾을 수 없습니다:\n{template_path}")
            return
        self.save_settings() # This now captures all setting_vars (Logo X/Y, Print Area, etc.)
        
        # 데이터가 비어있거나 새로 추출이 필요한 경우 수행
        if not data:
            if not self.extract_only(show_msg=False): return
            # Re-fetch data after extraction
            if mode == "PT": data = self.pt_extracted_data
            elif mode == "MT": data = self.mt_extracted_data
            elif mode == "RT": data = self.rt_extracted_data
            elif mode == "KOGAS": data = self.kogas_extracted_data
            else: data = self.extracted_data
            
        # [NEW] 체크된 항목만 필터링 (기본값은 True)
        final_list = [
            d for d in data
            if d.get('selected', True)
            and d.get('date_filtered', True)
            and (mode not in ('MT', 'RT', 'KOGAS') or d.get('drawing_filtered', True))
            and (mode != "PAUT" or (
                d.get('line_filtered', True) and d.get('size_filtered', True)))
        ]
        if not final_list:
            messagebox.showwarning("항목 미선택", f"선택된 {mode} 데이터가 없습니다. 미리보기에서 항목을 체크해주세요.")
            return

        if mode == "RT":
            self._run_rt_process(final_list, template_path, mode="RT")
        elif mode == "KOGAS":
            self._run_rt_process(final_list, template_path, mode="KOGAS")
        elif mode == "PT":
            self._run_pt_process(final_list, template_path)
        elif mode == "MT":
            self._run_mt_process(final_list, template_path)
        else:
            self._run_pmi_process(final_list, template_path)

    def _on_entry_esc(self, event):
        """Removes focus and clears selection on ESC (preserves text)."""
        try:
            widget = event.widget
            if hasattr(widget, 'selection_clear'):
                widget.selection_clear()
            self.root.focus_set()
        except: pass
