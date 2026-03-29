import sys
import os
import glob
import math
import traceback
import re
import warnings
import json
import tempfile
import datetime
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.pagebreak import Break
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, Border, Side
from PIL import Image as PILImage, ImageChops
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# Ignore library warnings
warnings.simplefilter("ignore")

# --- Constants & Paths ---
NAN_PATTERN = re.compile(r'^nan(\.0+)?$|^none$|^null$|^0\.0+|-0\.0+$', re.IGNORECASE)
DOT_ZERO_PATTERN = re.compile(r'\.0$')

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

if getattr(sys, 'frozen', False):
    SCRIPT_HOME = os.path.dirname(sys.executable)
    BASE_DIR = SCRIPT_HOME
    CONFIG_DIR = SCRIPT_HOME
    RESOURCE_DIR = SCRIPT_HOME
else:
    SCRIPT_HOME = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(SCRIPT_HOME)
    CONFIG_DIR = os.path.join(BASE_DIR, "config")
    RESOURCE_DIR = os.path.join(BASE_DIR, "resources")

SETTINGS_FILE = os.path.join(CONFIG_DIR, "logo_settings_unified.json")

class PMIReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SITCO 통합 성적서 자동 생성기 (PMI, RT, PAUT)")
        self.root.geometry("950x850") # Slightly wider/taller for multi-tab
        self.root.configure(background="#f9fafb")
        
        # 1. Initialize Configuration first (needed for state variables)
        self.config = {
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
            'RT_START_ROW': 11, 'RT_DATA_END_ROW': 45, 'RT_PRINT_END_ROW': 47
        }
        self.load_settings()

        # 2. State Variables
        self.logo_folder_path = tk.StringVar(value=RESOURCE_DIR)
        self.target_file_path = tk.StringVar()
        self.template_file_path = tk.StringVar()
        self.sequence_filter = tk.StringVar()
        self.extraction_mode = tk.StringVar(value="전체")
        self.auto_verify = tk.BooleanVar(value=True)
        self.show_selected_only = tk.BooleanVar(value=False)
        self.extracted_data = [] 
        self.sort_column = "" 
        self.sort_reverse = False 
        
        # [NEW] 원소 함량 필터링용 상태 변수 (PMI 전용)
        self.element_filters = [] # list of dict: {'key': StringVar, 'min': StringVar, 'max': StringVar}
        # 기본 필터 추가 (Cr, Ni, Mo)
        for k in ["Cr", "Ni", "Mo"]:
            self.element_filters.append({
                'key': tk.StringVar(value=k),
                'min': tk.StringVar(),
                'max': tk.StringVar()
            })
        
        # --- PAUT State Variables ---
        self.paut_target_file_path = tk.StringVar(value=self.config.get('PAUT_TARGET_PATH', ""))
        self.paut_template_file_path = tk.StringVar(value=self.config.get('PAUT_TEMPLATE_PATH', ""))
        self.paut_manual_vars = {
            't': tk.StringVar(), 'h': tk.StringVar(), 'l': tk.StringVar(), 'd': tk.StringVar(),
            'nature': tk.StringVar(value="Slag"), 'loc': tk.StringVar(value="-")
        }
        self.paut_extracted_data = []
        
        # --- RT State Variables ---
        self.rt_target_file_path = tk.StringVar(value=self.config.get('RT_TARGET_PATH', ""))
        self.rt_template_file_path = tk.StringVar(value=self.config.get('RT_TEMPLATE_PATH', ""))
        self.rt_extracted_data = []
        
        # --- PT State Variables ---
        self.pt_target_file_path = tk.StringVar(value=self.config.get('PT_TARGET_PATH', ""))
        self.pt_template_file_path = tk.StringVar(value=self.config.get('PT_TEMPLATE_PATH', ""))
        self.pt_extracted_data = []
        self.pt_item_idx_map = []

        # --- Column Definitions ---
        self.column_keys = ["selected", "No", "Date", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Grade"]
        self.rt_column_keys = ["selected", "No", "Date", "Dwg", "Joint", "Loc", "Accept", "Reject", "Grade", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "Welder", "Remarks"]
        self.pt_column_keys = ["selected", "No", "Date", "Dwg", "Joint", "NPS", "Thk.", "Material", "Welder", "WType", "Result"]

        # 3. UI Initialization
        self.create_widgets()
        self.log("[INFO] 통합 버전을 시작했습니다.")
        
    def log(self, message):
        if hasattr(self, 'status_log'):
            self.status_log.config(state='normal')
            self.status_log.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n")
            self.status_log.see(tk.END)
            self.status_log.config(state='disabled')
            self.root.update_idletasks()
        else:
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}")

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
        # [NEW] 로고 위치 고정(Fixed)을 위해 내장된 최적 설정을 우선시합니다.
        
        # 1. 외부 설정 파일 (사용자 오버라이드 시도)
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                    self.config.update(saved_data)
                print("SUCCESS: 외부 저장된 설정을 불러왔습니다.")
                
                # [NEW] 탭별 전용 행 설정 불러오기
                for tab in ["PMI", "PAUT"]:
                    for row_key in ['START_ROW', 'DATA_END_ROW', 'PRINT_END_ROW']:
                        tab_specific_key = f"{tab}_{row_key}"
                        if tab_specific_key in saved_data:
                            # 현재 탭에 맞는 설정이면 config에 직접 반영 (PAUT는 별도 변수 PAUT_START_ROW 등을 쓰므로 통합 처리)
                            pass 

                # [NEW] 저장된 필터 불러오기 (PMI)
                filter_key = "PMI_FILTERS"
                if filter_key in saved_data:
                    saved_filters = saved_data[filter_key]
                    if isinstance(saved_filters, list):
                        self.element_filters = []
                        for f_data in saved_filters:
                            self.element_filters.append({
                                'key': tk.StringVar(value=f_data.get('key', '')),
                                'min': tk.StringVar(value=f_data.get('min', '')),
                                'max': tk.StringVar(value=f_data.get('max', ''))
                            })
            except Exception as e:
                print(f"WARNING: 외부 설정 불러오기 실패: {e}")

        # 2. 내장 설정 파일 (실행파일 내부에 번들링된 최적값 - 외부 파일보다 우선순위 높임)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            bundle_settings = os.path.join(sys._MEIPASS, "logo_settings_unified.json")
            if os.path.exists(bundle_settings):
                try:
                    with open(bundle_settings, 'r', encoding='utf-8') as f:
                        saved_data = json.load(f)
                        self.config.update(saved_data)
                    print("SUCCESS: 내장된 최적 설정을 최종 적용했습니다. (고정 모드)")
                except Exception as e:
                    print(f"WARNING: 내장 설정 불러오기 실패: {e}")

        # [NEW] PAUT 마이그레이션 (UT_* -> PAUT_*)
        for k in list(self.config.keys()):
            if k.startswith("UT_"):
                new_k = k.replace("UT_", "PAUT_")
                if new_k not in self.config:
                    self.config[new_k] = self.config[k]

    def save_settings(self):
        try:
            if hasattr(self, 'setting_vars'):
                for key, var in self.setting_vars.items():
                    val = var.get()
                    try:
                        if key.endswith(('_X', '_Y', '_W', '_H')) or 'MARGIN' in key:
                            self.config[key] = float(val)
                        elif 'SCALE' in key or key.endswith('_ROW'):
                            self.config[key] = int(float(val))
                        else:
                            self.config[key] = str(val)
                    except: pass

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
            self.config["PMI_DATA_END_ROW"] = self.config.get("DATA_END_ROW", 45)
            self.config["PMI_PRINT_END_ROW"] = self.config.get("PRINT_END_ROW", 47)

            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            self.log("[SUCCESS] 설정이 파일에 저장되었습니다.")
        except Exception as e:
            self.log(f"[WARNING] 설정 저장 실패: {e}")

    def evaluate_paut_flaw(self, t, h, l, depth, flaw_nature):
        """
        ASME B31.1 (2024) - UT Acceptance Criteria Logic (Ported from archive)
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

        # 1. Immediate Rejection (Crack, LOF, IP)
        unacceptable_types = ['crack', 'lof', 'lack of fusion', 'ip', 'incomplete penetration']
        if str(flaw_nature).strip().lower() in unacceptable_types:
            return "Reject (Fatal Flaw Type)", loc
        
        if l_val <= 0 or h_val <= 0 or t_val <= 0:
            return "Error (Zero/Negative Value)", loc

        # 1.1 Special Rules for 6mm <= t < 13mm
        if 6 <= t_val < 13:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            if t_val < 10: h_surf_max, h_sub_max = 0.95, 0.96
            elif t_val < 12: h_surf_max, h_sub_max = 1.04, 1.04
            else: h_surf_max, h_sub_max = 1.13, 1.14
                
            limit = h_surf_max if loc == "Surface" else h_sub_max
            if h_val > limit: return f"Reject ({loc} h: {h_val} > {limit}mm)", loc
            return "Accept", loc

        # 1.2 Special Rules for 13mm <= t < 25.4mm
        if 13 <= t_val < 25.4:
            if l_val > 6.4: return f"Reject (L: {l_val} > 6.4mm)", loc
            actual_h_t = h_val / t_val
            allowed_h_t = 0.087 if loc == "Surface" else 0.143 
            if actual_h_t > allowed_h_t: return f"Reject ({loc} h/t: {actual_h_t:.3f} > {allowed_h_t:.3f})", loc
            return "Accept", loc
                
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
        if actual_a_t <= allowed_a_t: return "Accept", loc
        return f"Reject ({loc} a/t: {actual_a_t:.3f} > {allowed_a_t:.3f})", loc

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#f9fafb")
        style.configure("TLabel", background="#f9fafb", font=("Malgun Gothic", 10))
        style.configure("TLabelframe", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
        style.configure("TLabelframe.Label", background="#f9fafb", font=("Malgun Gothic", 10, "bold"))
        style.configure("Action.TButton", font=("Malgun Gothic", 11, "bold"), padding=10)
        
        # [NEW] 탭 스타일 설정
        style.configure("Main.TNotebook.Tab", font=("Malgun Gothic", 11, "bold"), padding=[15, 5])
        
        style.map("TEntry", 
                  selectbackground=[('focus', '#3b82f6'), ('!focus', '#3b82f6')],
                  selectforeground=[('focus', 'white'), ('!focus', 'white')])

        # --- Main Scrollable Container ---
        self.canvas = tk.Canvas(self.root, background="#f9fafb", highlightthickness=0, yscrollincrement=40)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, background="#f9fafb")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_frame_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_frame_window, width=e.width))
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.root.option_add('*selectBackground', '#3b82f6')
        self.root.option_add('*selectForeground', 'white')

        # Treeview Style
        style.configure("Treeview", rowheight=28, font=("Malgun Gothic", 10))
        style.map("Treeview", background=[('selected', '#3b82f6')], foreground=[('selected', 'white')])
        style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"))

        # [NEW] 전역 엔트리 우클릭 메뉴
        self._create_entry_context_menu()

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

        # Main Container
        main_container = tk.Frame(self.scrollable_frame, background="#f9fafb", padx=20, pady=20)
        main_container.pack(fill='both', expand=True)
        
        # --- Top Notebook (Multi-Mode) ---
        self.mode_notebook = ttk.Notebook(main_container, style="Main.TNotebook")
        self.mode_notebook.pack(fill='both', expand=True, pady=(0, 10))
        
        self.pmi_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.rt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.pt_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        self.paut_mode_frame = tk.Frame(self.mode_notebook, background="#f9fafb")
        
        self.mode_notebook.add(self.pmi_mode_frame, text="  PMI (성분 분석)  ")
        self.mode_notebook.add(self.rt_mode_frame, text="  RT (방사선 투과)  ")
        self.mode_notebook.add(self.pt_mode_frame, text="  PT (침투 탐상)  ")
        self.mode_notebook.add(self.paut_mode_frame, text="  PAUT (ASME B31.1)  ")
        
        # Setup each mode
        self._setup_pmi_ui(self.pmi_mode_frame)
        self._setup_rt_ui(self.rt_mode_frame)
        self._setup_pt_ui(self.pt_mode_frame)
        self._setup_paut_ui(self.paut_mode_frame)

        # Bottom Section (Common)
        bottom_frame = tk.Frame(main_container, background="#f9fafb")
        bottom_frame.pack(fill='x', side='bottom')

        self.progress = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill='x', pady=(10, 5))

        log_frame = ttk.LabelFrame(bottom_frame, text=" 작업 로그 (Status Log) ", padding=10)
        log_frame.pack(fill='both', expand=True)

        self.status_log = tk.Text(log_frame, height=8, font=("Consolas", 9), state='disabled', background="#000000", foreground="#10b981", padx=5, pady=5)
        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.status_log.yview)
        self.status_log.configure(yscrollcommand=vsb.set)
        self.status_log.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        # Dummy frame for scroll padding
        tk.Frame(self.scrollable_frame, height=50, background="#f9fafb").pack(side='bottom', fill='x')

    def _setup_pmi_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb", padx=10, pady=10)
        container.pack(fill='both', expand=True)

        tk.Label(container, text="PMI 성적서 생성 및 관리", font=("Malgun Gothic", 14, "bold"), background="#f9fafb", foreground="#111827").pack(pady=(0, 15), anchor='w')

        # 1. File Selection
        file_frame = ttk.LabelFrame(container, text=" 파일 및 폴더 선택 ", padding=15)
        file_frame.pack(fill='x', pady=(0, 20))

        def _add_file_row(parent_frame, label, var, row, is_dir=False, types=None):
            ttk.Label(parent_frame, text=label).grid(row=row, column=0, sticky='e', padx=5, pady=5)
            ttk.Entry(parent_frame, textvariable=var, width=50, exportselection=False).grid(row=row, column=1, padx=5, pady=5)
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_frame, text="찾기", command=cmd).grid(row=row, column=2, padx=5, pady=5)

        _add_file_row(file_frame, "로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "RFI 데이터:", self.target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "성적서 양식:", self.template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # 2. Configuration Tabs
        config_frame = ttk.LabelFrame(container, text=" 리포트 세부 설정 ", padding=5)
        config_frame.pack(fill='both', expand=False, pady=(0, 20))

        self.tab_notebook = ttk.Notebook(config_frame)
        self.tab_notebook.pack(fill='both', expand=True)

        self.setting_vars = {}
        tab_cover = ttk.Frame(self.tab_notebook, padding=10)
        tab_data = ttk.Frame(self.tab_notebook, padding=10)
        tab_rows = ttk.Frame(self.tab_notebook, padding=10)

        self.tab_notebook.add(tab_cover, text="갑지 (Cover)")
        self.tab_notebook.add(tab_data, text="을지 (Data)")
        self.tab_notebook.add(tab_rows, text="행 설정 (Rows)")

        self.tab_preview = ttk.Frame(self.tab_notebook, padding=5)
        self.tab_notebook.add(self.tab_preview, text="미리보기 (Preview)")
        self._create_preview_ui(self.tab_preview)

        self._create_setting_grid(tab_cover, "COVER")
        self._create_setting_grid(tab_data, "DATA")
        self._create_margin_settings(tab_cover, "COVER")
        self._create_margin_settings(tab_data, "DATA")
        self._create_row_settings(tab_rows)

        # 3. Action Section
        action_frame = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb", padx=15, pady=15)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=(0, 10))
        ttk.Label(filter_row, text="특정 순번(NO)만 추출 (예: 1, 3, 5-10):", background="#ffffff").pack(side='left', padx=(0, 10))
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=30, exportselection=False).pack(side='left')

        verify_row = tk.Frame(action_frame, background="#ffffff")
        verify_row.pack(fill='x', pady=(0, 10))
        tk.Checkbutton(verify_row, text="재질 자동 판정 (성분 비교 알고리즘 적용, 10% 허용오차)", variable=self.auto_verify, background="#ffffff", font=("Malgun Gothic", 10)).pack(side='left')
        
        # [NEW] 원소 함량 필터 UI (PMI 전용)
        filter_ui_frame = tk.Frame(action_frame, background="#ffffff", pady=5)
        filter_ui_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(filter_ui_frame, text="원소/컬럼 함량 필터 (Min~Max):", background="#ffffff", font=("Malgun Gothic", 10, "bold")).pack(side='left', padx=(0, 10))
        
        self.filter_container = tk.Frame(filter_ui_frame, background="#ffffff")
        self.filter_container.pack(side='left')
        
        def refresh_filter_ui():
            for widget in self.filter_container.winfo_children():
                widget.destroy()
            
            for i, f_item in enumerate(self.element_filters):
                unit_frame = tk.Frame(self.filter_container, background="#ffffff", padx=5)
                unit_frame.pack(side='left')
                
                ttk.Entry(unit_frame, textvariable=f_item['key'], width=5, exportselection=False).pack(side='left')
                ttk.Label(unit_frame, text=":", background="#ffffff").pack(side='left')
                ttk.Entry(unit_frame, textvariable=f_item['min'], width=5, exportselection=False).pack(side='left')
                ttk.Label(unit_frame, text="~", background="#ffffff").pack(side='left')
                ttk.Entry(unit_frame, textvariable=f_item['max'], width=5, exportselection=False).pack(side='left')
                
                # 삭제 버튼
                btn_del = tk.Button(unit_frame, text="×", command=lambda idx=i: remove_filter(idx), 
                                    relief='flat', background="#ffffff", foreground="red", font=("Arial", 10, "bold"))
                btn_del.pack(side='left', padx=(2, 0))

        def add_filter():
            self.element_filters.append({'key': tk.StringVar(), 'min': tk.StringVar(), 'max': tk.StringVar()})
            refresh_filter_ui()

        def remove_filter(idx):
            if len(self.element_filters) > 0:
                self.element_filters.pop(idx)
                refresh_filter_ui()

        ttk.Button(filter_ui_frame, text="+ 필터 추가", command=add_filter, width=10).pack(side='left', padx=10)
        refresh_filter_ui()
        
        mode_row = tk.Frame(action_frame, background="#ffffff")
        mode_row.pack(fill='x', pady=(0, 10))
        ttk.Label(mode_row, text="추출 방식 (Extraction Method):", background="#ffffff").pack(side='left', padx=(0, 10))
        mode_combo = ttk.Combobox(mode_row, textvariable=self.extraction_mode, state="readonly", width=25)
        mode_combo['values'] = ("전체", "SS304 만", "SS316 만", "DUPLEX 만", "SS310 만", "미분류(기타) 만")
        mode_combo.pack(side='left')

        btn_frame = tk.Frame(action_frame, background="#ffffff")
        btn_frame.pack(side='right')

        ttk.Button(btn_frame, text="데이터 추출 (Extract)", style="Action.TButton", command=self.extract_only).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="성적서 생성 시작 (Generate)", style="Action.TButton", command=self.run_process).pack(side='left', padx=5)

    def _setup_rt_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb", padx=10, pady=10)
        container.pack(fill='both', expand=True)

        tk.Label(container, text="RT 성적서 생성 및 관리", font=("Malgun Gothic", 14, "bold"), background="#f9fafb", foreground="#111827").pack(pady=(0, 15), anchor='w')

        # 1. File Selection
        file_frame = ttk.LabelFrame(container, text=" 파일 및 폴더 선택 ", padding=15)
        file_frame.pack(fill='x', pady=(0, 20))

        def _add_file_row(parent_frame, label, var, row, is_dir=False, types=None):
            ttk.Label(parent_frame, text=label).grid(row=row, column=0, sticky='e', padx=5, pady=5)
            # RT 탭에서도 엔트리 우클릭 메뉴를 지원하도록 exportselection=False 설정 유지
            ttk.Entry(parent_frame, textvariable=var, width=50, exportselection=False).grid(row=row, column=1, padx=5, pady=5)
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_frame, text="찾기", command=cmd).grid(row=row, column=2, padx=5, pady=5)

        _add_file_row(file_frame, "로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "RT 데이터 (Excel):", self.rt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "RT 성적서 양식:", self.rt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # 2. Configuration Tabs
        rt_config_frame = ttk.LabelFrame(container, text=" RT 리포트 세부 설정 ", padding=5)
        rt_config_frame.pack(fill='both', expand=False, pady=(0, 20))

        self.rt_tab_notebook = ttk.Notebook(rt_config_frame)
        self.rt_tab_notebook.pack(fill='both', expand=True)

        rt_tab_cover = ttk.Frame(self.rt_tab_notebook, padding=10)
        rt_tab_data = ttk.Frame(self.rt_tab_notebook, padding=10)
        rt_tab_rows = ttk.Frame(self.rt_tab_notebook, padding=10)

        self.rt_tab_notebook.add(rt_tab_cover, text="갑지 (Cover)")
        self.rt_tab_notebook.add(rt_tab_data, text="을지 (Data)")
        self.rt_tab_notebook.add(rt_tab_rows, text="행 설정 (Rows)")

        self.rt_tab_preview = ttk.Frame(self.rt_tab_notebook, padding=5)
        self.rt_tab_notebook.add(self.rt_tab_preview, text="미리보기 (Preview)")
        self._create_rt_preview_ui(self.rt_tab_preview)

        self._create_setting_grid(rt_tab_cover, "COVER")
        self._create_setting_grid(rt_tab_data, "DATA")
        self._create_margin_settings(rt_tab_cover, "COVER")
        self._create_margin_settings(rt_tab_data, "DATA")
        self._create_row_settings(rt_tab_rows)

        # 3. Action Section
        action_frame = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb", padx=15, pady=15)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=(0, 10))
        ttk.Label(filter_row, text="특정 순번(NO)만 추출 (예: 1, 3, 5-10):", background="#ffffff").pack(side='left', padx=(0, 10))
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=30, exportselection=False).pack(side='left')

        btn_frame = tk.Frame(action_frame, background="#ffffff")
        btn_frame.pack(side='right')

        ttk.Button(btn_frame, text="데이터 추출 (Extract)", style="Action.TButton", command=self.extract_only).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="성적서 생성 (Generate)", style="Action.TButton", command=self.run_process).pack(side='left', padx=5)

    def _setup_pt_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb", padx=10, pady=10)
        container.pack(fill='both', expand=True)

        tk.Label(container, text="PT 성적서 생성 및 관리", font=("Malgun Gothic", 14, "bold"), background="#f9fafb", foreground="#111827").pack(pady=(0, 15), anchor='w')

        # 1. File Selection
        file_frame = ttk.LabelFrame(container, text=" 파일 및 폴더 선택 ", padding=15)
        file_frame.pack(fill='x', pady=(0, 20))

        def _add_file_row(parent_frame, label, var, row, is_dir=False, types=None):
            ttk.Label(parent_frame, text=label).grid(row=row, column=0, sticky='e', padx=5, pady=5)
            # PT 탭에서도 엔트리 우클릭 메뉴를 지원하도록 exportselection=False 설정 유지
            ttk.Entry(parent_frame, textvariable=var, width=50, exportselection=False).grid(row=row, column=1, padx=5, pady=5)
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_frame, text="찾기", command=cmd).grid(row=row, column=2, padx=5, pady=5)

        _add_file_row(file_frame, "로고 폴더:", self.logo_folder_path, 0, is_dir=True)
        _add_file_row(file_frame, "PT 데이터 (Excel):", self.pt_target_file_path, 1, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "PT 성적서 양식:", self.pt_template_file_path, 2, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # 2. Configuration Tabs
        pt_config_frame = ttk.LabelFrame(container, text=" PT 리포트 세부 설정 ", padding=5)
        pt_config_frame.pack(fill='both', expand=False, pady=(0, 20))

        self.pt_tab_notebook = ttk.Notebook(pt_config_frame)
        self.pt_tab_notebook.pack(fill='both', expand=True)

        pt_tab_cover = ttk.Frame(self.pt_tab_notebook, padding=10)
        pt_tab_data = ttk.Frame(self.pt_tab_notebook, padding=10)
        pt_tab_rows = ttk.Frame(self.pt_tab_notebook, padding=10)

        self.pt_tab_notebook.add(pt_tab_cover, text="갑지 (Cover)")
        self.pt_tab_notebook.add(pt_tab_data, text="을지 (Data)")
        self.pt_tab_notebook.add(pt_tab_rows, text="행 설정 (Rows)")

        self.pt_tab_preview = ttk.Frame(self.pt_tab_notebook, padding=5)
        self.pt_tab_notebook.add(self.pt_tab_preview, text="미리보기 (Preview)")
        self._create_pt_preview_ui(self.pt_tab_preview)

        self._create_setting_grid(pt_tab_cover, "COVER")
        self._create_setting_grid(pt_tab_data, "DATA")
        self._create_margin_settings(pt_tab_cover, "COVER")
        self._create_margin_settings(pt_tab_data, "DATA")
        self._create_row_settings(pt_tab_rows)

        # 3. Action Section
        action_frame = tk.Frame(container, background="#ffffff", highlightthickness=1, highlightbackground="#e5e7eb", padx=15, pady=15)
        action_frame.pack(fill='x', pady=(0, 10))

        filter_row = tk.Frame(action_frame, background="#ffffff")
        filter_row.pack(fill='x', pady=(0, 10))
        ttk.Label(filter_row, text="특정 순번(NO)만 추출 (예: 1, 3, 5-10):", background="#ffffff").pack(side='left', padx=(0, 10))
        ttk.Entry(filter_row, textvariable=self.sequence_filter, width=30, exportselection=False).pack(side='left')

        btn_frame = tk.Frame(action_frame, background="#ffffff")
        btn_frame.pack(side='right')

        ttk.Button(btn_frame, text="데이터 추출 (Extract)", style="Action.TButton", command=self.extract_only).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="성적서 생성 (Generate)", style="Action.TButton", command=self.run_process).pack(side='left', padx=5)

    def _create_pt_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        self.pt_display_cols = ["V", "No", "Date", "ISO Drawing No.", "Joint", "NPS", "Thk.", "Material", "Welder", "Weld Type", "Result"]
        self.pt_widths = {"V": 40, "No": 50, "Date": 90, "ISO Drawing No.": 200, "Joint": 100, "NPS": 80, "Thk.": 80, "Material": 100, "Welder": 100, "Weld Type": 100, "Result": 80}

        tree_frame = tk.Frame(container, background="#f9fafb")
        tree_frame.pack(side='left', fill='both', expand=True)

        self.pt_preview_tree = ttk.Treeview(tree_frame, columns=self.pt_display_cols, show='headings', height=10, selectmode='extended')
        for col in self.pt_preview_tree["columns"]:
            self.pt_preview_tree.heading(col, text=col)
            self.pt_preview_tree.column(col, width=self.pt_widths.get(col, 100), anchor='center', stretch=False)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.pt_preview_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.pt_preview_tree.xview)
        self.pt_preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.pt_preview_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._setup_preview_sidebar(self.pt_preview_tree, container, mode="PT")

    def _setup_paut_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb", padx=10, pady=10)
        container.pack(fill='both', expand=True)

        tk.Label(container, text="PAUT (ASME B31.1) 성적서 생성 및 판정", font=("Malgun Gothic", 14, "bold"), background="#f9fafb", foreground="#111827").pack(pady=(0, 15), anchor='w')

        # 1. File Selection
        file_frame = ttk.LabelFrame(container, text=" 파일 및 폴더 선택 ", padding=15)
        file_frame.pack(fill='x', pady=(0, 15))

        def _add_file_row(parent_frame, label, var, row, is_dir=False, types=None):
            ttk.Label(parent_frame, text=label).grid(row=row, column=0, sticky='e', padx=5, pady=5)
            ttk.Entry(parent_frame, textvariable=var, width=50, exportselection=False).grid(row=row, column=1, padx=5, pady=5)
            cmd = (lambda: self._browse_dir(var)) if is_dir else (lambda: self._browse_file(var, types))
            ttk.Button(parent_frame, text="찾기", command=cmd).grid(row=row, column=2, padx=5, pady=5)

        _add_file_row(file_frame, "PAUT 데이터 (Excel):", self.paut_target_file_path, 0, types=[("Excel Source", "*.xls;*.xlsx;*.xlsm")])
        _add_file_row(file_frame, "성적서 양식 (Template):", self.paut_template_file_path, 1, types=[("Excel Template", "*.xlsx;*.xlsm")])

        # 2. Vertical Layout: Manual Eval (Top) & Preview/Batch (Bottom)
        # 2.1 Manual Evaluation (Top)
        manual_frame = ttk.LabelFrame(container, text=" 개별 판정 (Manual Evaluation) ", padding=15)
        manual_frame.pack(fill='x', pady=(0, 15))
        
        # Grid inside manual_frame
        m_inputs = [
            ("모재 두께 (t):", "t"), ("결함 높이 (h):", "h"), ("결함 길이 (l):", "l"), 
            ("결함 깊이 (d):", "d"), ("결함 종류:", "nature")
        ]
        
        for i, (lbl, key) in enumerate(m_inputs):
            ttk.Label(manual_frame, text=lbl).grid(row=0, column=i*2, sticky='w', padx=(10, 2), pady=5)
            if key == "nature":
                cb = ttk.Combobox(manual_frame, textvariable=self.paut_manual_vars[key], values=["Crack", "LOF", "IP", "Slag", "Porosity", "Others"], width=10)
                cb.grid(row=0, column=i*2+1, padx=5, sticky='w')
            else:
                ent = ttk.Entry(manual_frame, textvariable=self.paut_manual_vars[key], width=10)
                ent.grid(row=0, column=i*2+1, padx=5, sticky='w')
                if key != "l": 
                    ent.bind("<KeyRelease>", lambda e: self._update_paut_auto_loc())

        ttk.Label(manual_frame, text="판정 위치:").grid(row=1, column=0, sticky='w', padx=(10, 2), pady=10)
        ttk.Label(manual_frame, textvariable=self.paut_manual_vars['loc'], font=("Malgun Gothic", 10, "bold"), foreground="#3b82f6").grid(row=1, column=1, padx=5, sticky='w')

        btn_eval = ttk.Button(manual_frame, text="판정 실행", command=self._run_manual_paut_eval)
        btn_eval.grid(row=1, column=2, padx=20)

        self.paut_res_label = tk.Label(manual_frame, text="데이터를 입력하세요", font=("Malgun Gothic", 11, "bold"), background="#f3f4f6", height=2)
        self.paut_res_label.grid(row=1, column=3, columnspan=7, sticky='ew', padx=10)

        # 2.2 Preview & Batch (Bottom)
        batch_frame = ttk.LabelFrame(container, text=" 일괄 판정 및 미리보기 (Batch & Preview) ", padding=10)
        batch_frame.pack(fill='both', expand=True)

        # PAUT Preview Tree
        self.paut_preview_tree = ttk.Treeview(batch_frame, columns=("V", "No", "ISO/DWG", "Joint", "t", "h", "l", "d", "Location", "Nature", "Result"), show='headings', height=10)
        paut_widths = {"V": 30, "No": 40, "ISO/DWG": 120, "Joint": 80, "t": 40, "h": 40, "l": 40, "d": 40, "Location": 70, "Nature": 70, "Result": 80}
        for col in self.paut_preview_tree["columns"]:
            self.paut_preview_tree.heading(col, text=col)
            self.paut_preview_tree.column(col, width=paut_widths.get(col, 60), anchor='center')
        
        self.paut_preview_tree.pack(side='left', fill='both', expand=True)
        
        scroll = ttk.Scrollbar(batch_frame, orient="vertical", command=self.paut_preview_tree.yview)
        scroll.pack(side='right', fill='y')
        self.paut_preview_tree.configure(yscrollcommand=scroll.set)

        # Buttons
        btn_box = tk.Frame(container, background="#f9fafb")
        btn_box.pack(fill='x', pady=10)
        
        ttk.Button(btn_box, text="데이터 추출", command=self._extract_paut_data).pack(side='left', padx=5)
        ttk.Button(btn_box, text="일괄 판정 실행", command=self._run_batch_paut_eval).pack(side='left', padx=5)
        ttk.Button(btn_box, text="성적서 생성", command=self._generate_paut_report).pack(side='right', padx=5)

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

    def _run_manual_paut_eval(self):
        try:
            t = float(self.paut_manual_vars['t'].get() or 0)
            h = float(self.paut_manual_vars['h'].get() or 0)
            l = float(self.paut_manual_vars['l'].get() or 0)
            d = float(self.paut_manual_vars['d'].get() or 0)
            nat = self.paut_manual_vars['nature'].get()
            
            res, loc = self.evaluate_paut_flaw(t, h, l, d, nat)
            self.paut_manual_vars['loc'].set(loc)
            if res == "Accept":
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#27ae60")
            else:
                self.paut_res_label.config(text=f"{res} (위치: {loc})", fg="white", bg="#e74c3c")
        except Exception as e:
            messagebox.showerror("입력 오류", f"입력값을 확인해주세요: {e}")

    def _populate_paut_preview(self, data):
        self.paut_preview_tree.delete(*self.paut_preview_tree.get_children())
        for i, item in enumerate(data):
            v = "√" if item.get('selected', True) else ""
            self.paut_preview_tree.insert("", "end", values=(
                v, i + 1, item.get('ISO', ''), item.get('Joint', ''),
                item.get('t', ''), item.get('h', ''), item.get('l', ''),
                item.get('d', ''), item.get('Location', ''), item.get('Nature', ''), item.get('Result', '')
            ))

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
                "t": next((c for c in cols if any(x in c.upper() for x in ["THK", "두께", "THICKNESS"])), None),
                "h": next((c for c in cols if any(x in c.upper() for x in ["HEIGHT", "높이", "FLAW_HEIGHT"])), None),
                "l": next((c for c in cols if any(x in c.upper() for x in ["LENGTH", "길이", "FLAW_LENGTH"])), None),
                "d": next((c for c in cols if any(x in c.upper() for x in ["DEPTH", "깊이", "FLAW_DEPTH"])), None),
                "nature": next((c for c in cols if any(x in c.upper() for x in ["NAT", "종류", "FLAW_NATURE"])), None),
                "iso": next((c for c in cols if any(x in c.upper() for x in ["ISO", "DWG", "DRAWING"])), None),
                "joint": next((c for c in cols if any(x in c.upper() for x in ["JOINT", "WELD"])), None)
            }
            
            if not all([mapping["t"], mapping["h"], mapping["l"], mapping["d"]]):
                self.log("⚠️ 필수 컬럼을 모두 찾을 수 없어 자동 매핑에 실패했습니다.")
                
            self.paut_extracted_data = []
            for _, row in df.iterrows():
                item = {
                    'selected': True,
                    'ISO': str(row.get(mapping["iso"], "")) if mapping["iso"] else "",
                    'Joint': str(row.get(mapping["joint"], "")) if mapping["joint"] else "",
                    't': row.get(mapping["t"], 0),
                    'h': row.get(mapping["h"], 0),
                    'l': row.get(mapping["l"], 0),
                    'd': row.get(mapping["d"], 0),
                    'Nature': str(row.get(mapping["nature"], "Slag")) if mapping["nature"] else "Slag",
                    'Result': ""
                }
                self.paut_extracted_data.append(item)
            
            self.progress['value'] = 100
            self._populate_paut_preview(self.paut_extracted_data)
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
            res, loc = self.evaluate_paut_flaw(item['t'], item['h'], item['l'], item['d'], item['Nature'])
            item['Result'] = res
            item['Location'] = loc
            if res == "Accept": count_ok += 1
            else: count_ng += 1
            
        self._populate_paut_preview(self.paut_extracted_data)
        self.log(f"✅ 판정 완료: 총 {len(self.paut_extracted_data)} 건 (합격: {count_ok}, 불합격: {count_ng})")

    def _generate_paut_report(self):
        template_path = self.paut_template_file_path.get()
        if not template_path or not os.path.exists(template_path):
            messagebox.showwarning("파일 미선택", "PAUT 성적서 양식 파일을 선택해주세요.")
            return

        final_list = [d for d in self.paut_extracted_data if d.get('selected', True)]
        if not final_list:
            messagebox.showwarning("항목 미선택", "선택된 데이터가 없습니다.")
            return

        self.log(f"🚀 PAUT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        
        try:
            wb = openpyxl.load_workbook(template_path, keep_vba=True)
            ws = wb.worksheets[0]
            ws.title = f"PAUT_Report_001"
            
            self.add_logos_to_sheet(ws, is_cover=False)
            self.force_print_settings(ws, context="DATA")
            
            start_row = int(self.config.get('PAUT_START_ROW', 11))
            end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
            
            curr_row = start_row
            data_font = Font(size=9)
            
            for i, item in enumerate(final_list):
                if curr_row > end_row:
                    ws = self.prepare_next_paut_sheet(wb, ws.title, i//(end_row-start_row+1))
                    curr_row = start_row
                
                mapping = {
                    1: item.get('No', i+1), 2: item.get('ISO', ''), 3: item.get('Joint', ''),
                    4: item.get('t', ''), 5: item.get('h', ''), 6: item.get('l', ''),
                    7: item.get('d', ''), 8: item.get('Nature', ''), 9: item.get('Result', ''),
                    10: item.get('Location', '')
                }
                
                for col_idx, val in mapping.items():
                    cell = ws.cell(row=curr_row, column=col_idx, value=val)
                    cell.font = data_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    # Set border if needed
                
                curr_row += 1
                self.progress['value'] = (i+1) / len(final_list) * 100
            
            out_name = f"PAUT_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            out_path = os.path.join(BASE_DIR, "output", out_name) if os.path.exists(os.path.join(BASE_DIR, "output")) else os.path.join(BASE_DIR, out_name)
            
            if not os.path.exists(os.path.dirname(out_path)): os.makedirs(os.path.dirname(out_path))
            wb.save(out_path)
            self.log(f"✅ PAUT 성적서 생성 완료: {out_name}")
            os.startfile(os.path.dirname(out_path))
            
        except Exception as e:
            self.log(f"❌ PAUT 성적서 생성 오류: {e}")

    def prepare_next_paut_sheet(self, wb, prev_title, page_num):
        source_sheet = wb[prev_title]
        new_sheet = wb.copy_worksheet(source_sheet)
        new_sheet.title = f"PAUT_Report_{page_num+1:03d}"
        
        start_row = int(self.config.get('PAUT_START_ROW', 11))
        end_row = int(self.config.get('PAUT_DATA_END_ROW', 40))
        for r in range(start_row, end_row + 1):
            for c in range(1, 11): new_sheet.cell(row=r, column=c).value = None
        
        self.add_logos_to_sheet(new_sheet, is_cover=False)
        return new_sheet

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

    def _create_setting_grid(self, parent, context):
        # [REFINED] More organized layout for logo settings
        items = [
            ("SITCO 로고", f"SITCO_{context}"), 
            ("서울검사 로고", f"SEOUL_{context}"), 
            ("바닥글 우측 (Footer Right)", f"FOOTER_{context}"), 
            ("바닥글 좌측 (Footer Left)", f"FOOTER_PT_{context}")
        ]
        
        for i, (label, key_prefix) in enumerate(items):
            # Block Container
            block = tk.LabelFrame(parent, text=f" {label} ", padx=10, pady=10, background="#ffffff", font=("Malgun Gothic", 9, "bold"))
            block.grid(row=i, column=0, sticky='ew', pady=5, padx=2)
            parent.columnconfigure(0, weight=1)
            
            # Row 1: File Selection
            file_row = tk.Frame(block, background="#ffffff")
            file_row.pack(fill='x', pady=(0, 5))
            tk.Label(file_row, text="파일:", width=5, anchor='e', background="#ffffff").pack(side='left')
            v_path = tk.StringVar(value=self.config.get(f"{key_prefix}_PATH", ""))
            ttk.Entry(file_row, textvariable=v_path, exportselection=False).pack(side='left', fill='x', expand=True, padx=5)
            ttk.Button(file_row, text="찾기", width=5, command=lambda v=v_path: self._browse_file(v, [("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])).pack(side='right')
            self.setting_vars[f"{key_prefix}_PATH"] = v_path
            
            # Row 2: Coordinates
            coord_row = tk.Frame(block, background="#ffffff")
            coord_row.pack(fill='x')
            
            # Anchor
            tk.Label(coord_row, text="셀:", width=5, anchor='e', background="#ffffff").pack(side='left')
            v_a = tk.StringVar(value=self.config.get(f"{key_prefix}_ANCHOR", ""))
            ttk.Entry(coord_row, textvariable=v_a, width=6, exportselection=False).pack(side='left', padx=(0, 10))
            self.setting_vars[f"{key_prefix}_ANCHOR"] = v_a
            
            for coord, key_suffix in [("X", "X"), ("Y", "Y"), ("W", "W"), ("H", "H")]:
                tk.Label(coord_row, text=f"{coord}:", width=3, anchor='e', background="#ffffff").pack(side='left')
                v = tk.StringVar(value=str(self.config.get(f"{key_prefix}_{key_suffix}", "0.0")))
                ttk.Entry(coord_row, textvariable=v, width=6, exportselection=False).pack(side='left', padx=(0, 8))
                self.setting_vars[f"{key_prefix}_{key_suffix}"] = v
            
    def _create_margin_settings(self, parent, context):
        # [NEW] 여백 및 배율 설정용 UI (탭 하단에 배치)
        frame = ttk.LabelFrame(parent, text=" 인쇄 및 여백 설정 (Print & Margins) ", padding=10)
        frame.grid(row=5, column=0, columnspan=12, sticky='ew', pady=(20, 0))
        
        m_items = [("위(Top)", "TOP"), ("아래(Bottom)", "BOTTOM"), ("왼쪽(Left)", "LEFT"), ("오른쪽(Right)", "RIGHT")]
        for i, (label, key) in enumerate(m_items):
            ttk.Label(frame, text=label + ":").grid(row=0, column=i*2, sticky='e', padx=(10, 2))
            v = tk.StringVar(value=str(self.config.get(f"MARGIN_{context}_{key}", "0.2")))
            ttk.Entry(frame, textvariable=v, width=6).grid(row=0, column=i*2+1, sticky='w')
            self.setting_vars[f"MARGIN_{context}_{key}"] = v
            
        ttk.Label(frame, text="배율(%):").grid(row=0, column=8, sticky='e', padx=(20, 2))
        v_s = tk.StringVar(value=str(self.config.get(f"PRINT_SCALE_{context}", "95")))
        ttk.Entry(frame, textvariable=v_s, width=6).grid(row=0, column=9, sticky='w')
        self.setting_vars[f"PRINT_SCALE_{context}"] = v_s

        # [NEW] 선택적 행/열 조절 UI (추가 행 배치)
        sub_frame = ttk.Frame(frame)
        sub_frame.grid(row=1, column=0, columnspan=12, sticky='ew', pady=(10, 0))
        
        ttk.Label(sub_frame, text="행 높이 조절 (예: 1-10, 5):").grid(row=0, column=0, sticky='e', padx=(0, 2))
        v_rr = tk.StringVar(value=self.config.get(f"CUSTOM_ROWS_{context}", ""))
        ttk.Entry(sub_frame, textvariable=v_rr, width=15).grid(row=0, column=1, sticky='w')
        self.setting_vars[f"CUSTOM_ROWS_{context}"] = v_rr
        
        ttk.Label(sub_frame, text="높이:").grid(row=0, column=2, sticky='e', padx=(10, 2))
        v_rh = tk.StringVar(value=str(self.config.get(f"CUSTOM_ROW_HEIGHT_{context}", "16.5")))
        ttk.Entry(sub_frame, textvariable=v_rh, width=6).grid(row=0, column=3, sticky='w')
        self.setting_vars[f"CUSTOM_ROW_HEIGHT_{context}"] = v_rh

        ttk.Label(sub_frame, text="열 너비 조절 (예: A-C, E):").grid(row=0, column=4, sticky='e', padx=(20, 2))
        v_cr = tk.StringVar(value=self.config.get(f"CUSTOM_COLS_{context}", ""))
        ttk.Entry(sub_frame, textvariable=v_cr, width=15).grid(row=0, column=5, sticky='w')
        self.setting_vars[f"CUSTOM_COLS_{context}"] = v_cr
        
        ttk.Label(sub_frame, text="너비:").grid(row=0, column=6, sticky='e', padx=(10, 2))
        v_cw = tk.StringVar(value=str(self.config.get(f"CUSTOM_COL_WIDTH_{context}", "10.0")))
        ttk.Entry(sub_frame, textvariable=v_cw, width=6).grid(row=0, column=7, sticky='w')
        self.setting_vars[f"CUSTOM_COL_WIDTH_{context}"] = v_cw

        pass


    def _create_row_settings(self, parent):
        # PMI Rows
        pmi_frame = ttk.LabelFrame(parent, text=" PMI 행 설정 ", padding=10)
        pmi_frame.pack(fill='x', pady=5)
        
        pmi_rows = [
            ("데이터 시작 행", "START_ROW", "PMI 데이터 시작"), 
            ("데이터 종료 행", "DATA_END_ROW", "PMI 데이터 종료"), 
            ("인쇄 영역 종료 행", "PRINT_END_ROW", "PMI 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(pmi_rows):
            ttk.Label(pmi_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(pmi_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(pmi_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        # PAUT Rows
        ut_row_frame = ttk.LabelFrame(parent, text=" PAUT 행 설정 ", padding=10)
        ut_row_frame.pack(fill='x', pady=5)
        
        ut_rows = [
            ("PAUT 데이터 시작 행", "UT_START_ROW", "PAUT 데이터 시작"), 
            ("PAUT 데이터 종료 행", "UT_DATA_END_ROW", "PAUT 데이터 종료"), 
            ("PAUT 인쇄 영역 종료 행", "UT_PRINT_END_ROW", "PAUT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(ut_rows):
            ttk.Label(ut_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(ut_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(ut_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

        # RT Rows
        rt_row_frame = ttk.LabelFrame(parent, text=" RT 행 설정 ", padding=10)
        rt_row_frame.pack(fill='x', pady=5)
        
        rt_rows = [
            ("RT 데이터 시작 행", "RT_START_ROW", "RT 데이터 시작"), 
            ("RT 데이터 종료 행", "RT_DATA_END_ROW", "RT 데이터 종료"), 
            ("RT 인쇄 영역 종료 행", "RT_PRINT_END_ROW", "RT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(rt_rows):
            ttk.Label(rt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 0)))
            ttk.Entry(rt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(rt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')
            
        # PT Rows
        pt_row_frame = ttk.LabelFrame(parent, text=" PT 행 설정 ", padding=10)
        pt_row_frame.pack(fill='x', pady=5)
        
        pt_rows = [
            ("PT 데이터 시작 행", "PT_START_ROW", "PT 데이터 시작"), 
            ("PT 데이터 종료 행", "PT_DATA_END_ROW", "PT 데이터 종료"), 
            ("PT 인쇄 영역 종료 행", "PT_PRINT_END_ROW", "PT 인쇄 끝")
        ]
        for i, (label, key, tip) in enumerate(pt_rows):
            ttk.Label(pt_row_frame, text=label + ":").grid(row=i, column=0, sticky='e', padx=5, pady=5)
            v = tk.StringVar(value=str(self.config.get(key, 18)))
            ttk.Entry(pt_row_frame, textvariable=v, width=10).grid(row=i, column=1, sticky='w', padx=5)
            self.setting_vars[key] = v
            ttk.Label(pt_row_frame, text=f"({tip})", foreground="gray", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky='w')

    def _create_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        self.preview_tree = ttk.Treeview(container, columns=("V", "No", "Date", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"), show='headings', height=10, selectmode='extended')
        widths = {"V": 40, "No": 50, "Date": 90, "ISO/DWG": 180, "Joint No": 100, "Test Location": 100, "Ni": 60, "Cr": 60, "Mo": 60, "Grade": 100}
        for col in self.preview_tree["columns"]:
            self.preview_tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.preview_tree.column(col, width=widths.get(col, 80), anchor='center')
        
        self._setup_preview_sidebar(self.preview_tree, container, mode="PMI")

    def _create_rt_preview_ui(self, parent):
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        self.rt_display_cols = ["V", "No", "Date", "Drawing No.", "Film Ident. No.", "Film Location", "Acc", "Rej", "Deg", "① Crack", "② IP", "③ LF", "④ Slag", "⑤ Por", "⑥ U/C", "⑦ RUC", "⑧ BT", "⑨ TI", "⑩ CP", "⑪ RC", "⑫ Mis", "⑬ EP", "⑭ SD", "⑮ Oth", "Welder No", "Remarks"]
        self.rt_widths = {"V": 40, "No": 50, "Date": 90, "Drawing No.": 150, "Film Ident. No.": 120, "Film Location": 100, "Acc": 40, "Rej": 40, "Deg": 40, "Welder No": 100, "Remarks": 120}

        # Inner frame for horizontal/vertical scroll
        tree_frame = tk.Frame(container, background="#f9fafb")
        tree_frame.pack(side='left', fill='both', expand=True)

        self.rt_preview_tree = ttk.Treeview(tree_frame, columns=self.rt_display_cols, show='headings', height=10, selectmode='extended')
        for col in self.rt_preview_tree["columns"]:
            self.rt_preview_tree.heading(col, text=col)
            self.rt_preview_tree.column(col, width=self.rt_widths.get(col, 60), anchor='center', stretch=False)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.rt_preview_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.rt_preview_tree.xview)
        self.rt_preview_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.rt_preview_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._setup_preview_sidebar(self.rt_preview_tree, container, mode="RT")

    def _setup_preview_sidebar(self, tree, container, mode):
        # 1. Event Bindings
        self.drag_start_item = None
        
        def _on_tree_press(event, t=tree, m=mode):
            t.focus_set()
            item_id = t.identify_row(event.y)
            column = t.identify_column(event.x)
            
            if item_id:
                try:
                    col_idx = int(column.replace("#", "")) - 1
                    keys = self.rt_column_keys if m == "RT" else self.column_keys
                    data = self.rt_extracted_data if m == "RT" else self.extracted_data
                    idx_map = self.rt_item_idx_map if m == "RT" else self.item_idx_map
                    
                    if 0 <= col_idx < len(keys):
                        key = keys[col_idx]
                        view_idx = t.index(item_id)
                        actual_idx = idx_map[view_idx]
                        
                        # (1) Selection Toggle
                        if key == "selected":
                            data[actual_idx]['selected'] = not data[actual_idx].get('selected', True)
                            self.populate_preview(data, switch_tab=False, mode=m)
                            return "break"
                            
                        # (2) RT Specific Defect/Result Toggle
                        elif m == "RT" and ((key.startswith("D") and key[1:].isdigit()) or (key in ["Accept", "Reject"])):
                            old_v = data[actual_idx].get(key, "")
                            new_v = "√" if old_v == "" else ""
                            data[actual_idx][key] = new_v
                            if key in ["Accept", "Reject"] and new_v == "√":
                                other = "Reject" if key == "Accept" else "Accept"
                                data[actual_idx][other] = ""
                            self.populate_preview(data, switch_tab=False, mode="RT")
                            return "break"
                except: pass

                self.drag_start_item = item_id
                if not (event.state & 0x0001 or event.state & 0x0004):
                    t.selection_set(item_id)

        def _on_tree_drag(event, t=tree):
            if not self.drag_start_item: return
            curr_item = t.identify_row(event.y)
            if not curr_item: return
            all_items = t.get_children('')
            try:
                low = min(all_items.index(self.drag_start_item), all_items.index(curr_item))
                high = max(all_items.index(self.drag_start_item), all_items.index(curr_item))
                t.selection_set(all_items[low:high+1])
            except: pass

        def _on_tree_release(event):
            self.drag_start_item = None

        tree.bind("<Button-1>", _on_tree_press)
        tree.bind("<B1-Motion>", _on_tree_drag)
        tree.bind("<ButtonRelease-1>", _on_tree_release)
        tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        tree.bind("<Double-1>", lambda e, m=mode: self.on_tree_double_click(e, m))
        tree.bind("<Button-3>", lambda e, m=mode: self.show_context_menu(e, m))
        tree.bind("<Control-c>", lambda e, m=mode: self.copy_cell(m))
        tree.bind("<Control-v>", lambda e, m=mode: self.paste_cell(m))

        # Basic Tree Scroll (if not already handled in grid)
        if mode == "PMI":
            scroll_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scroll_y.set)
            tree.pack(side='left', fill='both', expand=True)
            scroll_y.pack(side='left', fill='y')

        # 2. Sidebar buttons
        sidebar = tk.Frame(container, background="#f9fafb", padx=10)
        sidebar.pack(side='right', fill='y')

        ttk.Label(sidebar, text="날짜 선택 필터", font=("Malgun Gothic", 9, "bold")).pack(pady=(5, 2))
        date_frame = tk.Frame(sidebar, background="#f9fafb")
        date_frame.pack(fill='x', pady=5)
        
        # Mode specific date listbox
        listbox = tk.Listbox(date_frame, selectmode='single', height=6, exportselection=False, font=("Malgun Gothic", 9))
        listbox.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(date_frame, orient='vertical', command=listbox.yview)
        sb.pack(side='right', fill='y')
        listbox.config(yscrollcommand=sb.set)
        
        if mode == "RT": self.rt_date_listbox = listbox
        else: self.date_listbox = listbox
        
        def _toggle_date(event, lb=listbox):
            idx = lb.nearest(event.y)
            if idx >= 0:
                val = lb.get(idx)
                if val.startswith("[v]"):
                    lb.delete(idx); lb.insert(idx, val.replace("[v] ", "[ ] "))
                elif val.startswith("[ ]"):
                    lb.delete(idx); lb.insert(idx, val.replace("[ ] ", "[v] "))
                lb.selection_clear(0, tk.END)
        
        listbox.bind("<ButtonRelease-1>", _toggle_date)
        
        def _apply_date_filter(m=mode, lb=listbox):
            sel_dates = [lb.get(i).replace("[v] ", "").replace("[ ] ", "") for i in range(lb.size()) if lb.get(i).startswith("[v]")]
            data = self.rt_extracted_data if m == "RT" else self.extracted_data
            for item in data:
                item['date_filtered'] = (item.get('Date', '') in sel_dates)
            self.populate_preview(data, switch_tab=False, mode=m)
            
        ttk.Button(sidebar, text="선택 날짜 적용", command=_apply_date_filter).pack(fill='x', pady=(0, 10))
        
        tk.Checkbutton(sidebar, text="선택 항목만 보기", variable=self.show_selected_only, 
                       command=lambda: self.populate_preview(self.rt_extracted_data if mode=="RT" else self.extracted_data, switch_tab=False, mode=mode),
                       background="#f9fafb", font=("Malgun Gothic", 9)).pack(pady=(0, 5), anchor='w')

        ttk.Button(sidebar, text="전체 선택", width=15, command=lambda: self.select_all(mode)).pack(pady=2)
        ttk.Button(sidebar, text="선택 해제", width=15, command=lambda: self.deselect_all(mode)).pack(pady=(2, 5))
        
        ttk.Button(sidebar, text="선택 항목 병합", width=15, command=lambda: self.merge_selected_iso(mode)).pack(pady=2)
        ttk.Button(sidebar, text="일괄 변경 (Bulk)", width=15, command=lambda: self.show_bulk_update_dialog(mode)).pack(pady=(2, 10))

        tk.Frame(sidebar, height=1, background="#e5e7eb").pack(fill='x', pady=5)
        ttk.Button(sidebar, text="▲ 위로", width=10, command=lambda: self.move_item(-1, mode)).pack(pady=2)
        ttk.Button(sidebar, text="▼ 아래로", width=10, command=lambda: self.move_item(1, mode)).pack(pady=2)
        tk.Frame(sidebar, height=10, background="#f9fafb").pack()
        ttk.Button(sidebar, text="선택 삭제", width=12, command=lambda: self.delete_item(mode)).pack(pady=2)
        ttk.Button(sidebar, text="전체 초기화", width=12, command=lambda: self.clear_all(mode)).pack(pady=(2, 10))

        tk.Frame(sidebar, height=1, background="#e5e7eb").pack(fill='x', pady=5)
        ttk.Button(sidebar, text="💾 현재 내용 저장", width=18, command=lambda: self.save_preview_data(mode)).pack(pady=2)
        ttk.Button(sidebar, text="📂 저장된 내용 열기", width=18, command=lambda: self.load_preview_data(mode)).pack(pady=2)
        ttk.Button(sidebar, text="📊 엑셀 파일로 추출", width=18, command=lambda: self.export_to_excel(mode)).pack(pady=(10, 2))

        # 3. Context Menu Config
        for t in ["evenrow", "oddrow", "group_even", "group_odd", "grouped_even", "grouped_odd", "item_even", "item_odd"]:
            tree.tag_configure(t, background="", foreground="black")
        tree.tag_configure('group_even', font=("Malgun Gothic", 10, "bold"))
        tree.tag_configure('group_odd', font=("Malgun Gothic", 10, "bold"))


    def on_tree_select(self, event):
        """항목 선택 시 체크박스 상태와 동기화 (현재는 드래그 안정성을 위해 드래그 중에는 무시)"""
        pass

    def toggle_item_selection(self, item_id, mode="PMI"):
        """단일 항목의 체크 상태를 토글"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
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
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
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
        """#1(토글), ISO/DWG(#4), Joint No(#5), Test Location(#6), 또는 Grade(#10) 컬럼 더블 클릭 처리"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        region = tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y)
        if not item_id: return

        if column == "#1": # Checkbox Column
            self.toggle_item_selection(item_id, mode)
            return

        # Editable columns mapping
        if mode == "RT":
            # RT: No(#2), Date(#3), Drawing No(#4), Film Ident(#5), Film Loc(#6), Deg(#9), Welder No(#25), Remarks(#26)
            key_map = {
                "#2": "No", "#3": "Date", "#4": "Dwg", "#5": "Joint", "#6": "Loc",
                "#9": "Grade", "#25": "Welder", "#26": "Remarks"
            }
        else:
            # PMI: No, Date, Dwg, Joint, Loc, Ni, Cr, Mo, Grade
            key_map = {
                "#2": "No", "#3": "Date", "#4": "Dwg", "#5": "Joint", "#6": "Loc",
                "#7": "Ni", "#8": "Cr", "#9": "Mo", "#10": "Grade"
            }
        
        key = key_map.get(column)
        if not key: return 
        
        x, y, w, h = tree.bbox(item_id, column)
        
        view_idx = tree.index(item_id)
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map

        if 0 <= view_idx < len(idx_map):
            actual_idx = idx_map[view_idx]
            old_val = data[actual_idx].get(key, "")
            if key in ["Ni", "Cr", "Mo"] and isinstance(old_val, (int, float)) and old_val > 0:
                old_val = f"{old_val:.2f}"
        else:
            old_val = tree.set(item_id, column)
        
        entry = ttk.Entry(tree, exportselection=True)
        entry.insert(0, old_val)
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        
        def finish_edit(event=None):
            new_val = entry.get()
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                if key in ["Ni", "Cr", "Mo"]:
                    data[actual_idx][key] = self.to_float(new_val)
                else:
                    data[actual_idx][key] = str(new_val).strip()
                self.populate_preview(data, switch_tab=False, mode=mode)
            entry.destroy()
            
        entry.bind("<Return>", finish_edit)
        entry.bind("<FocusOut>", finish_edit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def show_context_menu(self, event, mode="PMI"):
        """우클릭 컨텍스트 메뉴 표시 및 마지막 클릭 컬럼 기록"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        col = tree.identify_column(event.x)
        self.last_clicked_col = col # Track for copy/paste
        item_id = tree.identify_row(event.y)
        if item_id:
            if item_id not in tree.selection():
                tree.selection_set(item_id)
            
            # Update menu commands for current mode
            self.ctx_menu.entryconfigure("선택 항목 체크/해제 토글 (Toggle Check)", command=lambda: self.toggle_selected_items(mode))
            self.ctx_menu.entryconfigure("셀 내용 복사 (Copy)", command=lambda: self.copy_cell(mode))
            self.ctx_menu.entryconfigure("셀 내용 붙여넣기 (Paste)", command=lambda: self.paste_cell(mode))
            
            # Merge commands (Hide for RT if not applicable, but user might want them)
            self.ctx_menu.entryconfigure("선택 항목 ISO 병합 (Merge ISO)", command=lambda: self.merge_selected_iso(mode))
            self.ctx_menu.entryconfigure("선택 항목 Joint 병합 (Merge Joint)", command=lambda: self.merge_selected_joint(mode))
            self.ctx_menu.entryconfigure("선택 항목 Joint 시각적 그룹화 (Group Joint)", command=lambda: self.group_selected_joint(mode))
            self.ctx_menu.entryconfigure("선택 항목 Joint 그룹 해제 (Ungroup Joint)", command=lambda: self.ungroup_selected_joint(mode))
            self.ctx_menu.entryconfigure("선택 항목 일괄 변경 (Bulk Update)", command=lambda: self.show_bulk_update_dialog(mode))
            
            self.ctx_menu.post(event.x_root, event.y_root)


    def update_date_listbox(self, mode="PMI"):
        listbox = self.rt_date_listbox if mode == "RT" else self.date_listbox
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        if not listbox: return
        
        listbox.delete(0, tk.END)
        # [NEW] 데이터 기반으로 날짜별 선택 상태(date_filtered) 수집
        date_status = {}
        for item in data:
            dt = item.get('Date', 'N/A')
            if dt not in date_status:
                date_status[dt] = item.get('date_filtered', True)
        
        dates = sorted(list(date_status.keys()))
        for d in dates:
            is_v = date_status.get(d, True)
            prefix = "[v]" if is_v else "[ ]"
            listbox.insert(tk.END, f"{prefix} {d}")
            if is_v: listbox.select_set(tk.END)

    def copy_cell(self, mode="PMI"):
        """선택된 영역(드래그된 컬럼 범위)의 내용을 클립보드에 복사 (엑셀 스마트 영역 복사)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if not selected: return
        
        # Standardized column index to data key mapping
        if mode == "RT":
            key_map = {f"#{i+1}": k for i, k in enumerate(self.rt_column_keys)}
        else:
            key_map = {f"#{i+1}": k for i, k in enumerate(self.column_keys)}
        
        # [SMART AREA] 드래그 시작열과 종료열 사이의 범위를 계산
        try:
            s_col = getattr(self, 'start_col', '#4')
            e_col = getattr(self, 'end_col', s_col)
            
            s_idx = int(s_col.replace('#', '')) - 1
            e_idx = int(e_col.replace('#', '')) - 1
            
            col_start = min(s_idx, e_idx)
            col_end = max(s_idx, e_idx)
            
            last_col = getattr(self, 'last_clicked_col', None)
            if last_col and s_idx == e_idx and last_col.startswith('#'):
                l_idx = int(last_col.replace('#', '')) - 1
                col_start = min(col_start, l_idx)
                col_end = max(col_end, l_idx)
        except:
            col_start = col_end = 3 # ISO column (#4)
            
        target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
        
        copied_rows = []
        sorted_selected = sorted(list(selected), key=lambda x: tree.index(x))
        
        for item_id in sorted_selected:
            view_idx = tree.index(item_id)
            if not (0 <= view_idx < len(idx_map)):
                continue
            actual_idx = idx_map[view_idx]
            item = data[actual_idx]
            
            row_vals = []
            for c_id in target_col_ids:
                if c_id == "#1": # Select column (v/o)
                    status = "●" if item.get('selected', True) else "○"
                    row_vals.append(status)
                    continue

                key = key_map.get(c_id)
                if key:
                    val = item.get(key, "")
                    if mode == "PMI" and key in ["Ni", "Cr", "Mo"]:
                        f_val = self.to_float(val)
                        val = f"{f_val:.2f}" if f_val > 0 else ""
                    row_vals.append(str(val))
                else: 
                    row_vals.append(str(tree.set(item_id, c_id)))
            
            copied_rows.append("\t".join(row_vals))

        final_string = "\n".join(copied_rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(final_string)
        self.root.update()
        self.log(f"📋 '{mode}' 엑셀 영역 복사 완료: {len(sorted_selected)}행 x {len(target_col_ids)}열")

    def paste_cell(self, mode="PMI"):
        """클립보드 내용을 선택된 영역(드래그된 컬럼 범위)에 스마트하게 붙여넣기 (멀티행/아래로 채우기 지원)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        try: 
            clipboard_val = self.root.clipboard_get()
            if not clipboard_val: return
        except: return
        
        selected = list(tree.selection())
        if not selected: return
        
        # [SMART SORT] 시각적 순서(위에서 아래)대로 정렬
        selected.sort(key=lambda x: tree.index(x))
        
        # [SMART AREA] 드래그된 컬럼 범위를 계산
        try:
            s_idx = int(getattr(self, 'start_col', '#4').replace('#', '')) - 1
            e_idx = int(getattr(self, 'end_col', '#4').replace('#', '')) - 1
            col_start = min(s_idx, e_idx)
            col_end = max(s_idx, e_idx)
        except:
            col_start = col_end = 3 # Default ISO column (#4)
            
        target_col_ids = [f"#{i+1}" for i in range(col_start, col_end + 1)]
        if mode == "RT":
            key_map = {f"#{i+1}": k for i, k in enumerate(self.rt_column_keys)}
        else:
            key_map = {f"#{i+1}": k for i, k in enumerate(self.column_keys)}
        
        # 클립보드 데이터를 행/열 그리드로 파싱 (TAB/NEWLINE)
        paste_rows = [line.split("\t") for line in clipboard_val.strip().split("\n")]
        
        # [FILL DOWN] 한 줄만 선택했는데 클립보드가 여러 줄이면 아래로 자동 확장
        all_children = tree.get_children('')
        if len(selected) == 1 and len(paste_rows) > 1:
            curr_idx = all_children.index(selected[0])
            for i in range(1, len(paste_rows)):
                if curr_idx + i < len(all_children):
                    selected.append(all_children[curr_idx + i])
        
        for r_idx, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if not (0 <= view_idx < len(idx_map)):
                continue
            actual_idx = idx_map[view_idx]
            
            # 클립보드 데이터 매칭
            source_row = paste_rows[min(r_idx, len(paste_rows)-1)]
            
            for c_idx, c_id in enumerate(target_col_ids):
                cell_val = source_row[min(c_idx, len(source_row)-1)].strip()
                
                key = key_map.get(c_id)
                if key:
                    if mode == "PMI" and key in ["Ni", "Cr", "Mo"]:
                        data[actual_idx][key] = self.to_float(cell_val)
                    else:
                        data[actual_idx][key] = str(cell_val).strip()
        
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"📋 {mode} 엑셀 스마트 붙여넣기 완료: {len(selected)}행 x {len(target_col_ids)}열")

    def merge_selected_iso(self, mode="PMI"):
        """선택된 항목들의 ISO 번호를 첫 번째 항목의 것으로 통일 (병합 효과)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        # [SMART SORT] 시각적 순서대로 정렬하여 첫 번째 항목 식별
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            first_iso = data[actual_idx_0].get('Dwg', '')
        else:
            first_iso = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['Dwg'] = first_iso
                if i > 0: data[actual_idx]['is_merged_iso'] = True
                else: data[actual_idx].pop('is_merged_iso', None)
        
        self.populate_preview(data, switch_tab=False, mode=mode)
        self.log(f"🔗 {mode} {len(selected)}개 항목 ISO 병합 완료: {first_iso}")

    def merge_selected_joint(self, mode="PMI"):
        """선택된 항목들의 Joint No를 첫 번째 항목의 것으로 통일 (데이터 완전 변경)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
        selected = tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("알림", "병합할 항목을 2개 이상 선택해주세요.")
            return
        
        # [SMART SORT]
        selected = sorted(list(selected), key=lambda x: tree.index(x))
        
        view_idx_0 = tree.index(selected[0])
        actual_idx_0 = idx_map[view_idx_0] if 0 <= view_idx_0 < len(idx_map) else None
        
        if actual_idx_0 is not None:
            first_joint = data[actual_idx_0].get('Joint', '')
        else:
            first_joint = ""
        
        for i, item_id in enumerate(selected):
            view_idx = tree.index(item_id)
            if 0 <= view_idx < len(idx_map):
                actual_idx = idx_map[view_idx]
                data[actual_idx]['Joint'] = first_joint
                if i > 0: data[actual_idx]['is_merged_joint'] = True
                else: data[actual_idx].pop('is_merged_joint', None)
        
        self.populate_preview(data, switch_tab=False, mode=mode)
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

    def ungroup_selected_joint(self, mode="PMI"):
        """선택된 항목들의 시각적 그룹화 해제"""
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
        self.log(f"🔗 {mode} {len(selected)}개 항목 Joint 시각적 그룹화 해제 완료")

    def select_all(self, mode="PMI"):
        """모든 항목 체크"""
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        for item in data:
            if item.get('date_filtered', True):
                item['selected'] = True
        self.populate_preview(data, switch_tab=False, mode=mode)

    def deselect_all(self, mode="PMI"):
        """모든 항목 체크 해제"""
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        for item in data:
            if item.get('date_filtered', True):
                item['selected'] = False
        self.populate_preview(data, switch_tab=False, mode=mode)

    def move_item(self, direction, mode="PMI"):
        """선택된 아이템의 순서를 위/아래로 이동 (필터 상태 대응)"""
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
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
        tree = self.rt_preview_tree if mode == "RT" else self.preview_tree
        idx_map = self.rt_item_idx_map if mode == "RT" else self.item_idx_map
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        
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

    def save_preview_data(self, mode="PMI"):
        """현재 추출/편집된 데이터를 JSON 파일로 저장"""
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        if not data:
            messagebox.showwarning("알림", f"{mode} 저장할 데이터가 없습니다.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            initialfile=f"{mode}_Data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            title=f"{mode} 미리보기 데이터 저장"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                self.log(f"💾 {mode} 데이터 저장 완료: {os.path.basename(file_path)}")
                messagebox.showinfo("완료", "데이터가 성공적으로 저장되었습니다.")
            except Exception as e:
                self.log(f"❌ {mode} 저장 오류: {e}")
                messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {e}")

    def load_preview_data(self, mode="PMI"):
        """저장된 JSON 데이터를 불러와서 목록에 채움"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json")],
            title=f"{mode} 미리보기 데이터 불러오기"
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                    
                if not isinstance(loaded_data, list):
                    raise ValueError("올바른 데이터 형식이 아닙니다.")
                
                data = self.rt_extracted_data if mode == "RT" else self.extracted_data
                
                if data and not messagebox.askyesno("확인", f"현재 작업 중인 {mode} 데이터가 있습니다. 불러온 데이터로 덮어쓰시겠습니까?\n(아니오를 선택하면 기존 데이터에 추가됩니다.)"):
                    data.extend(loaded_data)
                else:
                    if mode == "RT": self.rt_extracted_data = loaded_data
                    else: self.extracted_data = loaded_data
                    data = loaded_data
            
                # [FIX] 필수 속성 보정
                for i, item in enumerate(data):
                    if 'date_filtered' not in item: item['date_filtered'] = True
                    if 'selected' not in item: item['selected'] = True
                    if 'order_index' not in item: item['order_index'] = i
                    
                self.update_date_listbox(mode)
                self.populate_preview(data, mode=mode)
                self.log(f"📂 {mode} 데이터 불러오기 완료: {os.path.basename(file_path)}")
                messagebox.showinfo("완료", f"데이터를 성공적으로 불러왔습니다.\n(총 {len(data)} 건)")
            except Exception as e:
                self.log(f"❌ {mode} 불러오기 오류: {e}")
                messagebox.showerror("오류", f"파일을 불러오는 중 오류가 발생했습니다: {e}")

    def clear_all(self, mode="PMI"):
        """모든 데이터 초기화"""
        if messagebox.askyesno("확인", f"모든 {mode} 데이터를 초기화하시겠습니까?"):
            if mode == "RT": self.rt_extracted_data = []
            else: self.extracted_data = []
            self.update_date_listbox(mode)
            self.populate_preview([], mode=mode)
            self.log(f"🧹 모든 {mode} 데이터 초기화 완료")

    def export_to_excel(self, mode="PMI"):
        """현재 미리보기 목록(필터링/선택 반영)을 엑셀 파일로 내보냄 (서식 및 병합 시인성 유지)"""
        data = self.rt_extracted_data if mode == "RT" else self.extracted_data
        if not data:
            messagebox.showwarning("알림", f"{mode} 내보낼 데이터가 없습니다.")
            return

        # 1. 엑셀에 들어갈 데이터 가공 (화면 표시 로직 재현)
        export_rows = []
        filter_enabled = self.show_selected_only.get()
        
        last_iso = None
        last_joint = None
        for idx, item in enumerate(data):
            if not item.get('date_filtered', True): continue
            if filter_enabled and not item.get('selected', True): continue
            
            curr_iso = item.get('Dwg', '')
            norm_iso = self.normalize_iso(curr_iso)
            curr_joint = item.get('visual_group_joint', item.get('Joint', ''))
            
            is_new_iso = (last_iso is None or self.normalize_iso(last_iso) != norm_iso)
            is_new_joint = (last_joint is None or last_joint != curr_joint)
            
            # PMI: 3-row block layout logic, RT: simple 1-row pagination usually
            is_block_start = (mode == "PMI" and len(export_rows) % 3 == 0)
            
            is_show = is_new_iso or is_new_joint or is_block_start
            
            d_iso = curr_iso if is_show else ""
            d_joint = item.get('Joint', '') if is_show else ""
            
            if item.get('is_merged_iso') and not (is_new_iso or is_new_joint): d_iso = ""
            if item.get('is_merged_joint') and not (is_new_iso or is_new_joint): d_joint = ""
            
            last_iso = curr_iso
            last_joint = curr_joint
            
            if mode == "RT":
                row_dict = {k: item.get(k, "") for k in self.rt_column_keys}
                # Overwrite ISO/Joint for visual merging
                row_dict['Dwg'] = d_iso
                row_dict['Joint'] = d_joint
                export_rows.append(row_dict)
            else:
                ni_val = self.to_float(item.get('Ni'))
                cr_val = self.to_float(item.get('Cr'))
                mo_val = self.to_float(item.get('Mo'))

                export_rows.append({
                    '순번': item.get('No', ''),
                    '날짜': item.get('Date', ''),
                    '도면번호': d_iso,
                    '조인트': d_joint,
                    '위치': item.get('Loc', ''),
                    'Ni': f"{ni_val:.2f}" if ni_val > 0 else "",
                    'Cr': f"{cr_val:.2f}" if cr_val > 0 else "",
                    'Mo': f"{mo_val:.2f}" if mo_val > 0 else "",
                    '판정': item.get('Grade', '')
                })

        if not export_rows:
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

            headers = list(export_rows[0].keys())
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')

            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row_idx, row_data in enumerate(export_rows, 2):
                for col_idx, key in enumerate(headers, 1):
                    val = row_data[key]
                    if mode == "RT" and key.startswith('D') and key[1:].isdigit():
                        # Keep check symbols as is
                        pass
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border

            if mode == "PMI":
                column_widths = {1: 8, 2: 12, 3: 25, 4: 15, 5: 15, 6: 8, 7: 8, 8: 8, 9: 15}
            else:
                # RT Column Widths
                column_widths = {1: 8, 2: 12, 3: 25, 4: 15, 5: 15, 6: 10} # Simplified for now
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

    def sort_by_column(self, col):
        """특정 컬럼을 기준으로 데이터 전체 정렬 (문자열/숫자 구분 및 헤더 표시)"""
        if not self.extracted_data: return
        
        # 1. 정렬 기준 컬럼 매핑 (Treeview 헤더 이름 -> 데이터 키)
        key_map = {
            "V": "selected", "No": "No", "Date": "Date", "ISO/DWG": "Dwg",
            "Joint No": "Joint", "Test Location": "Loc", "Ni": "Ni", 
            "Cr": "Cr", "Mo": "Mo", "Grade": "Grade"
        }
        data_key = key_map.get(col)
        if not data_key: return

        # 2. 정렬 방향 결정 (기존 방향과 같으면 토글)
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False # 기본 오름차순 (A~Z, 1~9)

        # 3. 데이터 정렬 (지능형 복합 자연어 정렬: Interleaved Natural Sort)
        def get_natural_key(val):
            """문자열을 텍스트와 숫자로 조각내어 에러 없이 통합된 자연어 정렬 수행"""
            if val is None: return [(2, "")]
            
            s_val = str(val).strip().lower()
            if not s_val: return [(2, "")]
            
            # [ (타입, 값), (타입, 값)... ] 구조로 반환
            # (0, 숫자) 가 (1, 문자) 보다 우선순위가 높음 (0 < 1)
            # 이를 통해 "1" -> [(0, 1)], "1A" -> [(0, 1), (1, 'a')], "2" -> [(0, 2)] 순서 보장
            def segment_to_tuple(text):
                if text.isdigit():
                    return (0, int(text))
                return (1, text)
            
            return [segment_to_tuple(c) for c in re.split(r'(\d+)', s_val) if c]

        def get_value(item, key):
            val = item.get(key, "")
            
            # 선택(V) 컬럼 전용 처리
            if key == "selected":
                return [0, 0 if val is True else 1]
            
            return get_natural_key(val)

        # 정렬 수행: 계층적 정렬 (사용자 요청 기반)
        try:
            # 1순위 ISO(Dwg), 2순위 Joint No 를 기본 뼈대로 사용
            k_iso, k_joint = "Dwg", "Joint"
            
            if data_key in [k_iso, k_joint]:
                # ISO나 Joint를 클릭하면 ISO > Joint > order_index 순서로 안정적 정렬
                sort_key = lambda x: (get_value(x, k_iso), get_value(x, k_joint), x.get('order_index', 0))
            else:
                # 그 외 컬럼 클릭 시 해당 컬럼이 1순위, 그 뒤에 ISO > Joint > order_index 보조
                sort_key = lambda x: (get_value(x, data_key), get_value(x, k_iso), get_value(x, k_joint), x.get('order_index', 0))
            
            self.extracted_data.sort(key=sort_key, reverse=self.sort_reverse)
        except Exception as e:
            self.log(f"⚠️ 정렬 오류: {e}")

        # 4. 헤더 텍스트 갱신 (정렬 화살표 표시)
        for c in ["V", "No", "Date", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"]:
            prefix = ""
            if c == col:
                prefix = "▼ " if self.sort_reverse else "▲ "
            self.preview_tree.heading(c, text=prefix + c)

        # 5. 화면 갱신 및 로그
            self.populate_preview(self.extracted_data, switch_tab=False)
        self.log(f"📊 '{col}' 기준 {'내림차순' if self.sort_reverse else '오름차순'} 정렬 완료")



    def populate_preview(self, data_list, switch_tab=True, mode="PMI"):
        """추출된 데이터를 미리보기 표에 채움 (필터 반영 및 그룹 색상 적용)"""
        if mode == "RT":
            tree = self.rt_preview_tree
            self.rt_item_idx_map = []
            idx_map = self.rt_item_idx_map
        elif mode == "PT":
            tree = self.pt_preview_tree
            self.pt_item_idx_map = []
            idx_map = self.pt_item_idx_map
        else:
            tree = self.preview_tree
            self.item_idx_map = []
            idx_map = self.item_idx_map

        for item in tree.get_children():
            tree.delete(item)
        
        filter_enabled = self.show_selected_only.get()
        
        last_iso = None
        last_joint = None
        current_tag = "group_even"
        
        for idx, item in enumerate(data_list):
            if not item.get('date_filtered', True):
                continue
            
            is_selected = item.get('selected', True)
            if filter_enabled and not is_selected:
                continue
            
            idx_map.append(idx)
            v_mark = "●" if is_selected else "○"
            
            # ISO/DWG 번호가 바뀌면 배경색 태그 교체
            curr_iso = item.get('Dwg', '')
            norm_iso = self.normalize_iso(curr_iso)
            
            # [NEW] 시각적 병합 (visual_group_joint) 값이 존재하면 해당 값으로 그룹화 계산
            curr_joint = item.get('visual_group_joint', item.get('Joint', ''))
            
            # [FIX] Joint-Aware: Joint가 바뀌었는지도 함께 체크
            is_new_iso = (last_iso is None or self.normalize_iso(last_iso) != norm_iso)
            is_new_joint = (last_joint is None or last_joint != curr_joint)
            
            if is_new_iso:
                current_tag = "group_odd" if current_tag == "group_even" else "group_even"

            # [REFINED] 하이브리드 표시 로직: 블록 시작점 정렬이거나 "새로운 그룹(ISO or Joint)" 시작점일 때 표시
            display_count = len(idx_map) - 1 
            is_block_start = (display_count % 3 == 0)
            
            # ISO나 Joint가 바뀌었거나 3줄 단위 시작이면 표시
            is_show = is_new_iso or is_new_joint or is_block_start

            display_iso = curr_iso if is_show else ""
            display_joint = item.get('Joint', '') if is_show else ""
            
            # [FIX] 명시적으로 병합된 경우라도 '새 그룹 시작'이면 표시 강제
            if item.get('is_merged_iso') and not (is_new_iso or is_new_joint): display_iso = ""
            if item.get('is_merged_joint') and not (is_new_iso or is_new_joint): display_joint = ""

            last_iso = curr_iso
            last_joint = curr_joint
            
            # Row Values based on mode
            if mode == "RT":
                row_vals = [v_mark]
                row_vals.extend([
                    item.get('No', ''), item.get('Date', ''),
                    display_iso, display_joint, item.get('Loc', ''),
                    item.get('Accept', ''), item.get('Reject', ''), item.get('Grade', '')
                ])
                # Add D1 to D15
                for di in range(1, 16):
                    row_vals.append(item.get(f'D{di}', ''))
                row_vals.extend([item.get('Welder', ''), item.get('Remarks', '')])
            elif mode == "PT":
                row_vals = [
                    v_mark, item.get('No', ''), item.get('Date', ''),
                    item.get('Dwg', ''), item.get('Joint', ''),
                    item.get('NPS', ''), item.get('Thk.', ''),
                    item.get('Material', ''), item.get('Welder', ''),
                    item.get('WType', ''), item.get('Result', '')
                ]
            else:
                row_vals = [
                    v_mark, item.get('No', ''), item.get('Date', ''),
                    display_iso, display_joint, item.get('Loc', ''),
                    f"{self.to_float(item.get('Ni')): .2f}" if self.to_float(item.get('Ni')) > 0 else "",
                    f"{self.to_float(item.get('Cr')): .2f}" if self.to_float(item.get('Cr')) > 0 else "",
                    f"{self.to_float(item.get('Mo')): .2f}" if self.to_float(item.get('Mo')) > 0 else "",
                    item.get('Grade', '')
                ]

            row_tags = [str(idx), current_tag]
            tree.insert("", "end", values=tuple(row_vals), tags=tuple(row_tags))
            
        if switch_tab:
            if mode == "RT":
                self.rt_tab_notebook.select(self.rt_tab_preview)
            elif mode == "PT":
                self.pt_tab_notebook.select(self.pt_tab_preview)
            else:
                self.tab_notebook.select(self.tab_preview)

    def _browse_dir(self, var):
        path = filedialog.askdirectory(initialdir=var.get() or RESOURCE_DIR)
        if path: var.set(path)

    def _browse_file(self, var, types):
        path = filedialog.askopenfilename(initialdir=os.path.dirname(var.get() or BASE_DIR), filetypes=types)
        if path: var.set(path)

    def show_bulk_update_dialog(self):
        """[NEW] 선택된 항목들을 필터 조건에 따라 일괄 변경하는 다이얼로그 표시"""
        if not self.extracted_data:
            messagebox.showwarning("알림", "처리할 데이터가 없습니다.")
            return

        # 체크된 항목 추출
        selected_indices = [idx for idx, item in enumerate(self.extracted_data) if item.get('selected', True)]
        if not selected_indices:
            messagebox.showwarning("항목 미선택", "일괄 변경할 항목을 체크(선택)해주세요.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("선택 항목 일괄 변경")
        dialog.geometry("450x550")
        dialog.configure(background="#f9fafb")
        dialog.transient(self.root)
        dialog.grab_set()

        # Center the dialog
        dialog.geometry(f"+{self.root.winfo_x() + 100}+{self.root.winfo_y() + 50}")

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="1. 대상 필터링 (선택 사항)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 10))
        
        filter_frame = ttk.Frame(main_frame)
        filter_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(filter_frame, text="필터 열:").grid(row=0, column=0, sticky='w')
        col_var = tk.StringVar(value="Test Location")
        col_combo = ttk.Combobox(filter_frame, textvariable=col_var, state='readonly', width=15,
                                 values=["전체(필터 없음)", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"])
        col_combo.grid(row=0, column=1, padx=5, sticky='w')
        
        ttk.Label(filter_frame, text="필터 값:").grid(row=1, column=0, sticky='w', pady=5)
        val_var = tk.StringVar(value="WELD")
        ttk.Entry(filter_frame, textvariable=val_var, width=18).grid(row=1, column=1, padx=5, sticky='w', pady=5)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)

        ttk.Label(main_frame, text="2. 변경할 값 입력 (입력된 항목만 반영)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 10))
        
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill='both', expand=True)
        
        entries = {}
        row_idx = 0
        for label, key in [("Ni (%):", "Ni"), ("Cr (%):", "Cr"), ("Mo (%):", "Mo"), 
                           ("Location:", "Loc"), ("Grade:", "Grade"), ("Date:", "Date")]:
            ttk.Label(fields_frame, text=label).grid(row=row_idx, column=0, sticky='e', pady=3, padx=5)
            entries[key] = ttk.Entry(fields_frame, width=20)
            entries[key].grid(row=row_idx, column=1, sticky='w', pady=3)
            row_idx += 1

        def apply_action():
            filter_col = col_var.get()
            filter_val = val_var.get().strip().lower()
            
            # Update mappings for filtering
            key_map = {"ISO/DWG": "Dwg", "Joint No": "Joint", "Test Location": "Loc", "Ni": "Ni", "Cr": "Cr", "Mo": "Mo", "Grade": "Grade"}
            target_key = key_map.get(filter_col)
            
            update_spec = {}
            for key, entry in entries.items():
                val = entry.get().strip()
                if val:
                    if key in ["Ni", "Cr", "Mo"]:
                        # Support relative values like +1.0 or -0.5
                        if val.startswith(('+', '-')):
                            try:
                                update_spec[key] = ('relative', float(val))
                            except: update_spec[key] = ('absolute', self.to_float(val))
                        else:
                            update_spec[key] = ('absolute', self.to_float(val))
                    else:
                        update_spec[key] = ('absolute', val)
            
            if not update_spec:
                messagebox.showwarning("입력 부족", "변경할 값을 하나 이상 입력해주세요.")
                return

            # Apply updates
            count = 0
            for idx in selected_indices:
                item = self.extracted_data[idx]
                
                # Check filter
                match = True
                if target_key:
                    curr_val = str(item.get(target_key, "")).strip().lower()
                    if filter_val not in curr_val: # Partial match for convenience
                        match = False
                
                if match:
                    for k, spec in update_spec.items():
                        mode, val = spec
                        if mode == 'relative':
                            item[k] = self.to_float(item.get(k, 0)) + val
                        else:
                            item[k] = val
                    count += 1
            
            self.populate_preview(self.extracted_data, switch_tab=False)
            messagebox.showinfo("일괄 변경 완료", f"총 {count}개의 항목이 업데이트되었습니다.")
            dialog.destroy()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(20, 0))
        ttk.Button(btn_frame, text="적용 (Apply)", command=apply_action).pack(side='right', padx=5)
        ttk.Button(btn_frame, text="취소 (Cancel)", command=dialog.destroy).pack(side='right')

    # --- Integrated Verification Logic ---

    def to_float(self, val):
        if pd.isna(val): return 0.0
        s = str(val).upper().replace("%", "").strip()
        if "<" in s or "ND" in s or s == "": return 0.0
        try: return float(s)
        except: return 0.0

    def check_material_grade(self, row_data):
        """jjchRFIPMI.py에서 이식된 10% 여유치 판정 로직"""
        cr = row_data.get('Cr', 0.0)
        ni = row_data.get('Ni', 0.0)
        mo = row_data.get('Mo', 0.0)
        mn = row_data.get('Mn', 0.0)
        
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

    # --- Excel Helper Logic ---

    def find_image_smart(self, keyword, exclude_keyword=None):
        def _search_in_folder(folder_path):
            if not folder_path or not os.path.exists(folder_path): return None
            candidates = glob.glob(os.path.join(folder_path, "*.*"))
            valid_extensions = ['.PNG', '.JPG', '.JPEG', '.BMP', '.GIF']
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
        
        # 2. PyInstaller 묶음(실행파일 내부 임시 폴더)에서 검색 (Standalone 지원)
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            found = _search_in_folder(sys._MEIPASS)
            if found: return found
            
        return None

    def place_image_freely(self, ws, img_path, anchor_cell_str, w, h, x_offset, y_offset):
        try:
            if not img_path or not os.path.exists(img_path): return
            # [NEW] Check for valid anchor cell string
            if not anchor_cell_str or not isinstance(anchor_cell_str, str) or len(anchor_cell_str) < 2:
                self.log(f"[WARNING] {os.path.basename(img_path)} 배치 실패: 잘못된 앵커 셀 ({anchor_cell_str})")
                return

            original = PILImage.open(img_path).convert("RGBA")
            # [NEW] Ensure w, h are positive
            w = max(1, int(w)); h = max(1, int(h))
            resized = original.resize((w, h), PILImage.Resampling.LANCZOS)
            temp_name = f"temp_{os.path.basename(img_path)}"
            temp_full_path = os.path.join(tempfile.gettempdir(), temp_name)
            resized.save(temp_full_path)
            
            img = XLImage(temp_full_path); img.width = w; img.height = h
            col_str, row_num = coordinate_from_string(anchor_cell_str)
            col_idx = max(0, column_index_from_string(col_str) - 1)
            row_idx = max(0, row_num - 1) 
            # [FIX] EMU offsets and dimensions MUST be non-negative
            emu_x = max(0, int(x_offset * 9525)); emu_y = max(0, int(y_offset * 9525))
            emu_w = max(1, int(w * 9525)); emu_h = max(1, int(h * 9525))
            
            marker = AnchorMarker(col=col_idx, colOff=emu_x, row=row_idx, rowOff=emu_y)
            size = XDRPositiveSize2D(cx=emu_w, cy=emu_h)
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(img)
        except Exception as e: self.log(f"[WARNING] {os.path.basename(img_path)} 배치 실패: {e}")

    def add_logos_to_sheet(self, ws, is_cover=False, clear_existing=True):
        if clear_existing:
            try: ws._images = [] 
            except: pass
        ctx = "COVER" if is_cover else "DATA"
        def _get_params(key_prefix):
            return (self.config.get(f"{key_prefix}_{ctx}_PATH", ""), self.config[f"{key_prefix}_{ctx}_ANCHOR"], self.config[f"{key_prefix}_{ctx}_W"], self.config[f"{key_prefix}_{ctx}_H"], self.config[f"{key_prefix}_{ctx}_X"], self.config[f"{key_prefix}_{ctx}_Y"])
        
        # 1. SITCO
        path, anchor, w, h, x, y = _get_params("SITCO")
        if not path or not os.path.exists(path): path = self.find_image_smart("SITCO")
        self.place_image_freely(ws, path, anchor, w, h, x, y)
        
        # 2. SEOUL
        path, anchor, w, h, x, y = _get_params("SEOUL")
        if not path or not os.path.exists(path): path = self.find_image_smart("서울검사")
        self.place_image_freely(ws, path, anchor, w, h, x, y)
        
        # 3. FOOTER
        path, anchor, w, h, x, y = _get_params("FOOTER")
        if not path or not os.path.exists(path):
            path = self.find_image_smart("바닥글", exclude_keyword="PMI-1")
            if not path: path = self.find_image_smart("PMI", exclude_keyword="PMI-1")
        self.place_image_freely(ws, path, anchor, w, h, x, y)
        
        # 4. FOOTER_PT (Left)
        path, anchor, w, h, x, y = _get_params("FOOTER_PT")
        if not path or not os.path.exists(path):
            path = self.find_image_smart("PMI갑") if is_cover else None
            if not path: path = self.find_image_smart("PMI-1")
            if not path: path = self.find_image_smart("PT")
        self.place_image_freely(ws, path, anchor, w, h, x, y)

    def force_print_settings(self, ws, context="DATA"):
        try:
            # [NEW] 갑지는 T열까지, 을지는 M열까지 인쇄 영역 설정
            ws.print_area = 'A1:T51' if context == "COVER" else 'A1:M47'
            ws.page_setup.paperSize = 9; ws.page_setup.orientation = 'portrait'
            
            # [REFINED] 사용자 설정 배율 및 여백 적용
            ws.page_setup.scale = int(float(self.config.get(f'PRINT_SCALE_{context}', 95)))
            ws.print_options.horizontalCentered = True; ws.print_options.verticalCentered = True
            
            ws.page_margins.top = float(self.config.get(f'MARGIN_{context}_TOP', 0.2))
            ws.page_margins.bottom = float(self.config.get(f'MARGIN_{context}_BOTTOM', 0.2))
            ws.page_margins.left = float(self.config.get(f'MARGIN_{context}_LEFT', 0.5))
            ws.page_margins.right = float(self.config.get(f'MARGIN_{context}_RIGHT', 0.3))
        except Exception as e:
            print(f"[WARNING] Print settings failed: {e}")

    def apply_custom_dimensions(self, ws, context):
        """[NEW] 사용자의 행/열 조절 설정을 파싱하여 적용"""
        try:
            # Row Adjustment
            row_range_str = self.config.get(f"CUSTOM_ROWS_{context}", "").strip()
            if row_range_str:
                height = float(self.config.get(f"CUSTOM_ROW_HEIGHT_{context}", 16.5))
                for part in row_range_str.split(','):
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        for r in range(start, end + 1):
                            ws.row_dimensions[r].height = height
                    else:
                        ws.row_dimensions[int(part)].height = height
            
            # Column Adjustment
            col_range_str = self.config.get(f"CUSTOM_COLS_{context}", "").strip()
            if col_range_str:
                width = float(self.config.get(f"CUSTOM_COL_WIDTH_{context}", 10.0))
                for part in col_range_str.split(','):
                    if '-' in part:
                        start_letter, end_letter = part.split('-')
                        for c in range(column_index_from_string(start_letter), column_index_from_string(end_letter) + 1):
                            col_letter = openpyxl.utils.get_column_letter(c)
                            ws.column_dimensions[col_letter].width = width
                    else:
                        ws.column_dimensions[part.strip().upper()].width = width
        except Exception as e:
            print(f"[DEBUG] Custom dimension failed: {e}")

    def safe_set_value(self, ws, coord, value, align=None):
        """병합된 셀이라도 기준 셀(Top-Left)을 찾아 안전하게 값을 입력합니다. (성능 최적화 버전)"""
        try:
            # 전달받은 인자가 이미 셀 객체인 경우
            if hasattr(coord, 'coordinate'):
                target_cell = coord
                coord_str = target_cell.coordinate
            else:
                target_cell = ws[coord]
                coord_str = coord

            # 해당 셀이 MergedCell(병합 영역의 기준 셀이 아닌 곳)인 경우에만 탐색 수행
            if isinstance(target_cell, MergedCell):
                for m_range in ws.merged_cells.ranges:
                    if coord_str in m_range:
                        target_cell = ws.cell(row=m_range.min_row, column=m_range.min_col)
                        break
            
            target_cell.value = value
            if align:
                target_cell.alignment = Alignment(horizontal=align, vertical='center')
        except Exception: pass

    def safe_merge_cells(self, ws, start_row, start_column, end_row, end_column):
        """이미 병합된 영역이 있는지 확인하고 안전하게 병합을 수행합니다."""
        try:
            # 병합하려는 영역이 이미 다른 병합 영역과 겹치는지 체크
            # openpyxl은 겹치는 병합을 허용하지 않으므로, 겹치면 해당 영역을 풀거나 스킵해야 함
            ws.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column)
        except Exception as e:
            # 겹치는 경우 기존 병합을 해제하고 시도하거나, 로그만 남김
            self.log(f"⚠️ 병합 실패 ({start_row},{start_column}~{end_row},{end_column}): {e}")

    def set_eulji_headers(self, ws):
        headers = ["NI", "CR", "MO"]
        data_font = Font(size=9); header_row = self.config['START_ROW']
        for i, val in enumerate(headers):
            col = 8 + i
            cell = ws.cell(row=header_row, column=col)
            self.safe_set_value(ws, cell.coordinate, val, align='center')
            cell.font = data_font
        
        materials = "SS304,SS304L,SS316,SS316L,SS321,SS347,SS410,SS430,DUPLEX,MONEL,INCONEL,ER308,ER308L,ER309,ER309L,ER316,ER316L,ER347,ER2209,WP316,WP316L,TP316,TP316L,F316L,A182-F316L,A312-TP316L"
        dv_q = DataValidation(type="list", formula1=f'"{materials}"', allow_blank=True)
        ws.add_data_validation(dv_q)
        for r in range(self.config['START_ROW'], self.config['DATA_END_ROW'] + 1):
            target_l = ws.cell(row=r, column=13); dv_q.add(target_l) # 12 -> 13 (M)
            target_l.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center'); target_l.font = Font(size=8.5)
    
        # [NEW] K15에 하이픈 추가
        self.safe_set_value(ws, 'K15', "-")
        ws['K15'].alignment = Alignment(horizontal='center', vertical='center')
        
    def prepare_next_sheet(self, wb, source_sheet_idx, page_num):
        source_sheet = wb.worksheets[source_sheet_idx]; new_sheet = wb.copy_worksheet(source_sheet) 
        base_title = source_sheet.title.split('_')[0]; new_sheet.title = f"{base_title[:20]}_{page_num:03d}"
        self.force_print_settings(new_sheet, context="DATA"); self.add_logos_to_sheet(new_sheet, is_cover=False)
        self.apply_custom_dimensions(new_sheet, "DATA")
        for col_letter, col_dim in source_sheet.column_dimensions.items(): new_sheet.column_dimensions[col_letter].width = col_dim.width
        data_font = Font(size=9); grade_font = Font(size=8.5)
        for r in range(self.config['START_ROW'], self.config['DATA_END_ROW'] + 1):
            rd = new_sheet.row_dimensions[r] # [REMOVED] Hardcoded 20.55 override
        
        for r in range(self.config['START_ROW'] + 1, self.config['DATA_END_ROW'] + 1):
            for c in range(1, 14):
                cell = new_sheet.cell(row=r, column=c)
                cell.font = grade_font if c == 13 else data_font
                self.safe_set_value(new_sheet, cell, None)
        merged_to_clear = [rng for rng in new_sheet.merged_cells.ranges if rng.min_row >= self.config['START_ROW'] and rng.max_row <= self.config['DATA_END_ROW']]
        for rng in merged_to_clear: new_sheet.unmerge_cells(str(rng))
        
        # [FIX] 갑지 데이터 수식으로 연결 (첫번째 시트 참조)
        try:
            ws0 = wb.worksheets[0]
            if len(wb.worksheets) > 1:
                # 엑셀 수식을 사용하여 갑지의 값이 바뀌면 자동으로 바뀌게 설정
                self.safe_set_value(new_sheet, 'K5', f"='{ws0.title}'!L5")
                self.safe_set_value(new_sheet, 'M5', f"='{ws0.title}'!N5")
                self.safe_set_value(new_sheet, 'M8', f"='{ws0.title}'!N8")
                # [FIX] K5:M10 범위 글씨체 바탕, 크기 9 적용
                for r_idx in range(5, 11):
                    for c_idx in range(11, 14): # K(11) ~ M(13)
                        cell = new_sheet.cell(row=r_idx, column=c_idx)
                        cell.font = Font(name='바탕', size=9, bold=False)
                        # K5, K8는 줄바꿈 적용, 그 외는 셀 크기에 맞춤 적용
                        if (r_idx == 5 or r_idx == 8) and c_idx == 11:
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
        except: pass
        
        # [NEW] K15에 하이픈 추가
        self.safe_set_value(new_sheet, 'K15', "-")
        new_sheet['K15'].alignment = Alignment(horizontal='center', vertical='center')
        
        new_sheet.column_dimensions['M'].width = 18.0
        return new_sheet

    def extract_only(self, show_msg=True):
        """데이터만 추출하여 리스트와 미리보기에 반영 (PMI/RT 대응)"""
        # 현재 활성 탭에 따른 모드 결정
        try:
            tab_idx = self.mode_notebook.index("current") # Main notebook (PMI, RT, PT...)
            if tab_idx == 1: mode = "RT"
            elif tab_idx == 2: mode = "PT"
            else: mode = "PMI"
        except: mode = "PMI"

        target_file = self.pt_target_file_path.get() if mode == "PT" else (self.rt_target_file_path.get() if mode == "RT" else self.target_file_path.get())
        if not target_file:
            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일을 선택해주세요.")
            return False
            
        self.log(f"🔍 {mode} 데이터 추출 시작: {os.path.basename(target_file)}")
        
        # [NEW] Extract date from filename
        fname = os.path.basename(target_file)
        # More flexible regex for dates like 2024.03.05, 24.03.05, 24_03_05, 240305, etc.
        date_match = re.search(r'(\d{4}[-._]\d{2}[-._]\d{2}|\d{2}[-._]\d{2}[-._]\d{2}|\d{8}|\d{6})', fname)
        extracted_date = date_match.group(0) if date_match else ""
        
        # Standardize formatting
        if extracted_date:
            # Remove separators like . or _
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

        self.progress['value'] = 0
        all_extracted_data = []
        
        try:
            target_input = self.sequence_filter.get().strip()
            target_no_list = [x.strip() for x in target_input.replace(',', ' ').split() if target_input and x.strip()] if target_input else []
            
            with pd.ExcelFile(target_file) as xls:
                for s_idx, sheet_name in enumerate(xls.sheet_names):
                    self.log(f"📄 시트 스캔: {sheet_name}")
                    try: temp_df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=50)
                    except: continue
                    header_idx = None
                    for i, row in temp_df.iterrows():
                        row_str = str(row.values).upper()
                        if mode == "RT":
                            if "FILM" in row_str or "DEFECT" in row_str or "IQI" in row_str:
                                header_idx = i; break
                        elif mode == "PT":
                            if (("LINE" in row_str or "ISO" in row_str or "DWG" in row_str) and ("JOINT" in row_str or "WELD" in row_str)):
                                header_idx = i; break
                        else:
                            if ("CR" in row_str and "NI" in row_str) or ("CHROMIUM" in row_str):
                                header_idx = i; break
                    if header_idx is None: continue
                    
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_idx)
                    
                    # [NEW] [Refinement] Two-line header detection and merging
                    if not df.empty and len(df) > 0:
                        row0_str = " ".join([str(x).upper() for x in df.iloc[0] if pd.notna(x)])
                        if any(k in row0_str for k in ["ORIGIN", "FACTOR", "WELDER1"]):
                            self.log(f"   ℹ️ [자동보정] 두 줄 헤더 감지됨 (시트: {sheet_name})")
                            new_cols = []
                            for col, val in zip(df.columns, df.iloc[0]):
                                c_txt = str(col) if pd.notna(col) and "Unnamed" not in str(col) else ""
                                v_txt = str(val) if pd.notna(val) else ""
                                new_cols.append(f"{c_txt} {v_txt}".strip())
                            df.columns = new_cols
                            df = df.iloc[1:].reset_index(drop=True)

                def _find_col(df, keywords, exclude=None):
                    # 1. 우선 정확히 일치하는 컬럼을 찾음
                    for col in df.columns:
                        c_up = str(col).upper().strip()
                        if exclude and any(ex in c_up for ex in exclude): continue
                        if any(k == c_up for k in keywords): return col
                    # 2. 포함된 컬럼을 찾되, NI의 경우 UNIT 등 오진 가능성 차단
                    for col in df.columns:
                        c_up = str(col).upper().strip()
                        if exclude and any(ex in c_up for ex in exclude): continue
                        if any(k in c_up for k in keywords):
                            if "NI" in keywords and ("UNIT" in c_up or "LINE" in c_up): continue
                            return col
                    return None
                if mode == "RT":
                    col_no = _find_col(df, ["NO.", "NO", "SEQ", "ITEM"])
                    col_dwg = _find_col(df, ["ISO", "DWG", "DRAWING", "LINE"])
                    col_joint = _find_col(df, ["JOINT", "WELD NO", "J/N", "FILM IDENT"])
                    col_loc = _find_col(df, ["LOCATION", "POSITION", "FILM LOC"])
                    col_grade = _find_col(df, ["GRADE", "DEG"]) # DEG for RT
                    col_welder = _find_col(df, ["WELDER", "W/N"])
                    col_remarks = _find_col(df, ["REMARKS", "REMARK", "비고"])
                    # RT Specifics
                    col_t = _find_col(df, ["T", "THICK", "THK"])
                    col_mat = _find_col(df, ["MAT", "MATERIAL"])
                    col_weld = _find_col(df, ["WELD", "TYPE"])
                    col_iqi = _find_col(df, ["IQI"])
                    col_sens = _find_col(df, ["SENS", "SENSITIVITY"])
                    col_den = _find_col(df, ["DEN", "DENSITY"])
                    col_result = _find_col(df, ["RESULT", "판정", "ACC", "REJ"])
                    defect_cols = {}
                    for i in range(1, 16):
                        c = _find_col(df, [f"D{i}", f"DEFECT{i}"])
                        if c: defect_cols[f"D{i}"] = c
                elif mode == "PT":
                    col_no = _find_col(df, ["NO.", "NO", "SEQ", "ITEM"])
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
                    col_no = _find_col(df, ["NO.", "NO", "SEQ", "NUM", "POS", "ITEM"])
                    col_joint = _find_col(df, ["JOINT", "J/N", "JOINT NO", "PUNCH", "WELD NO"])
                    col_loc = _find_col(df, ["LOCATION", "TEST POSITION", "POINT", "AREA", "POSITION"])
                    col_dwg = _find_col(df, ["ISO", "DWG", "DRAWING", "LINE"])
                    col_grade_orig = _find_col(df, ["GRADE", "MATERIAL", "SPEC", "TYPE"])

                # [NEW] Forward-fill (Fill Down) for ISO and Joint
                last_dwg = ""
                last_joint = ""

                for _, row in df.iterrows():
                    v_raw_no = str(row[col_no]).strip() if col_no is not None else str(_+1)
                    if target_no_list and v_raw_no not in target_no_list: continue
                    
                    # [NEW] Common Fill-down Logic
                    curr_dwg = str(row[col_dwg]).strip() if col_dwg is not None else ""
                    if (not curr_dwg or curr_dwg == "nan") and last_dwg: curr_dwg = last_dwg
                    if curr_dwg and curr_dwg != "nan": last_dwg = curr_dwg

                    curr_joint = str(row[col_joint]).strip() if col_joint is not None else ""
                    if (not curr_joint or curr_joint == "nan") and last_joint: curr_joint = last_joint
                    if curr_joint and curr_joint != "nan": last_joint = curr_joint
                    elif not curr_joint or curr_joint == "nan": curr_joint = v_raw_no

                    if mode == "RT":
                        # RT-specific Row Data
                        item_data = {
                            'No': v_raw_no, 'Date': extracted_date, 'Dwg': curr_dwg, 'Joint': curr_joint,
                            'Loc': str(row[col_loc]).strip() if col_loc is not None else "",
                            'Grade': str(row[col_grade]).strip() if col_grade is not None else "",
                            'Welder': str(row[col_welder]).strip() if col_welder is not None else "",
                            'Remarks': str(row[col_remarks]).strip() if col_remarks is not None else "",
                            'T': str(row[col_t]).strip() if col_t is not None else "",
                            'Mat': str(row[col_mat]).strip() if col_mat is not None else "",
                            'Weld': str(row[col_weld]).strip() if col_weld is not None else "",
                            'IQI': str(row[col_iqi]).strip() if col_iqi is not None else "",
                            'Sens': str(row[col_sens]).strip() if col_sens is not None else "",
                            'Den': str(row[col_den]).strip() if col_den is not None else "",
                            'Result': str(row[col_result]).strip() if col_result is not None else "ACC",
                            'selected': True,
                            'order_index': len(self.rt_extracted_data) + len(all_extracted_data)
                        }
                        # Defects D1-D15 (Normalization)
                        for i in range(1, 16):
                            key = f"D{i}"
                            c = defect_cols.get(key)
                            val = str(row[c]).strip() if c is not None else ""
                            if val and val.lower() in ["v", "x", "o", "1", "√", "v"]: val = "√"
                            else: val = ""
                            item_data[key] = val
                            
                        all_extracted_data.append(item_data)
                    elif mode == "PT":
                        # PT-specific Row Data (Filter for ACC only)
                        res_str = str(row[col_result]).upper() if col_result is not None else "ACC"
                        is_pass = any(k in res_str for k in ["ACC", "OK", "ACCEPT", "합격"])
                        is_fail = any(k in res_str for k in ["REJ", "RW", "FAIL", "UNACC"])
                        if is_pass and not is_fail:
                            size_v = str(row[col_size]).strip() if col_size is not None else ""
                            thk_v = str(row[col_thk]).strip() if col_thk is not None else ""
                            # SCH -> Thk 변환
                            thk_converted = convert_sch_to_thk(size_v, thk_v)
                            
                            all_extracted_data.append({
                                'No': v_raw_no, 'Date': extracted_date, 'Dwg': curr_dwg, 'Joint': self.force_two_digit(curr_joint),
                                'NPS': size_v, 'Thk.': thk_converted, 
                                'Material': self.fix_material_name(row[col_mat]) if col_mat is not None else "",
                                'Welder': str(row[col_welder]).strip() if col_welder is not None else "",
                                'WType': str(row[col_wtype]).strip() if col_wtype is not None else "",
                                'Result': "Acc",
                                'selected': True,
                                'order_index': len(self.pt_extracted_data) + len(all_extracted_data)
                            })
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

                            all_extracted_data.append({
                                'No': v_raw_no, 
                                'Joint': curr_joint,
                                'Loc': str(row[col_loc]).strip() if col_loc is not None else "",
                                'Cr': v_cr, 'Ni': v_ni, 'Mo': v_mo, 'Mn': v_mn,
                                'Grade': final_grade, 'Dwg': curr_dwg,
                                'Date': extracted_date,
                                'selected': True,
                                'order_index': len(self.extracted_data) + len(all_extracted_data)
                            })
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
                        if f_min > 0 and f_max > 0:
                            all_extracted_data = [d for d in all_extracted_data if f_min <= d.get(f_key, d.get(f_key.capitalize(), 0.0)) <= f_max]
                        elif f_min > 0:
                            all_extracted_data = [d for d in all_extracted_data if d.get(f_key, d.get(f_key.capitalize(), 0.0)) >= f_min]
                        elif f_max > 0:
                            all_extracted_data = [d for d in all_extracted_data if d.get(f_key, d.get(f_key.capitalize(), 0.0)) <= f_max]
                        
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

            # [CHANGE] Overwrite -> Accumulate
            if mode == "RT":
                self.rt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("RT")
                self.populate_preview(self.rt_extracted_data, mode="RT")
                total_count = len(self.rt_extracted_data)
            elif mode == "PT":
                self.pt_extracted_data.extend(all_extracted_data)
                self.update_date_listbox("PT")
                self.populate_preview(self.pt_extracted_data, mode="PT")
                total_count = len(self.pt_extracted_data)
            else:
                self.extracted_data.extend(all_extracted_data)
                self.update_date_listbox("PMI")
                self.sort_column = "ISO/DWG"; self.sort_reverse = False
                self.sort_by_column("ISO/DWG") # Auto sort for PMI
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

    def run_process(self):
        # 현재 활성 탭에 따른 모드 결정
        try:
            tab_idx = self.mode_notebook.index("current")
            if tab_idx == 1: mode = "RT"
            elif tab_idx == 2: mode = "PT"
            else: mode = "PMI"
        except: mode = "PMI"

        if mode == "RT":
            target_file = self.rt_target_file_path.get()
            template_path = self.rt_template_file_path.get()
            data = self.rt_extracted_data
        elif mode == "PT":
            target_file = self.pt_target_file_path.get()
            template_path = self.pt_template_file_path.get()
            data = self.pt_extracted_data
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
            
        self.save_settings()
        
        # 데이터가 비어있거나 새로 추출이 필요한 경우 수행
        if not data:
            if not self.extract_only(show_msg=False): return
            # Re-fetch data after extraction
            data = self.pt_extracted_data if mode == "PT" else (self.rt_extracted_data if mode == "RT" else self.extracted_data)
            
        # [NEW] 체크된 항목만 필터링 (기본값은 True)
        final_list = [d for d in data if d.get('selected', True) and d.get('date_filtered', True)]
        if not final_list:
            messagebox.showwarning("항목 미선택", f"선택된 {mode} 데이터가 없습니다. 미리보기에서 항목을 체크해주세요.")
            return

        if mode == "RT":
            self._run_rt_process(final_list, template_path)
        elif mode == "PT":
            self._run_pt_process(final_list, template_path)
        else:
            self._run_pmi_process(final_list, template_path)

    def _run_pmi_process(self, final_list, template_path):
        self.log(f"🚀 PMI 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0

        # No changes needed here, just marking for clarity.
        self.progress['value'] = 0
        
        try:
            all_extracted_data = final_list
            wb = openpyxl.load_workbook(template_path, keep_vba=True)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            if len(wb.worksheets) >= 1:
                ws0 = wb.worksheets[0]; self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False)
                self.force_print_settings(ws0, context="COVER") # [NEW] 갑지 전용 여백 적용
                
                # [NEW] 갑지 전용 여백 적용
                self.force_print_settings(ws0, context="COVER")
                for r in range(23, 39):
                    try:
                        cell_a = ws0.cell(row=r, column=1); eb = cell_a.border
                        cell_a.border = Border(left=medium_side, right=eb.right, top=eb.top, bottom=eb.bottom)
                    except: pass
                
                self.safe_set_value(ws0, 'I35', None) # [FIX] Border() 대신 가독성을 위해 safe_set_value 적용
                self.apply_custom_dimensions(ws0, "COVER") # [MOVED] Ensure it has the final say
            
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]; ws.title = f"{ws.title[:20]}_001"
            # 을지 기본 설정
            self.add_logos_to_sheet(ws, is_cover=False); self.force_print_settings(ws, context="DATA"); self.set_eulji_headers(ws)
            # self.apply_custom_dimensions(ws, "DATA") # [MOVED] To the end of process
            
            try:
                if len(wb.worksheets) >= 2:
                    # [FIX] 수식으로 연결 (갑지 L5 -> 을지 K5, 갑지 N5 -> 을지 M5, 갑지 N8 -> 을지 M8)
                    self.safe_set_value(ws, 'K5', f"='{ws0.title}'!L5")
                    self.safe_set_value(ws, 'M5', f"='{ws0.title}'!N5")
                    self.safe_set_value(ws, 'M8', f"='{ws0.title}'!N8")
                    # [FIX] K5:M10 범위 스타일 설정 (바탕, 크기 9)
                    for r_idx in range(5, 11):
                        for c_idx in range(11, 14):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.font = Font(name='바탕', size=9, bold=False)
                            if (r_idx == 5 or r_idx == 8) and c_idx == 11: # K5, K8 줄바꿈
                                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            else:
                                cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    ws.column_dimensions['M'].width = 18.0
            except Exception as e:
                self.log(f"갑지 데이터 복사 실패: {e}")
            
            materials = "SS304,SS304L,SS316,SS316L,SS321,SS347,SS410,SS430,DUPLEX,MONEL,INCONEL,ER308,ER308L,ER309,ER309L,ER316,ER316L,ER347,ER2209,WP316,WP316L,TP316,TP316L,F316L,A182-F316L,A312-TP316L"
            dv_q = DataValidation(type="list", formula1=f'"{materials}"', allow_blank=True)
            # [NEW] 데이터 기입 전 기존 병합 영역 정리 (A-E 등 중복 병합 방지)
            def clear_merges_in_range(sheet, start_row, end_row):
                merged_ranges = list(sheet.merged_cells.ranges)
                for r in merged_ranges:
                    # 영역이 겹치면 해제 (aggressive clearing)
                    if r.min_row <= end_row and r.max_row >= start_row:
                        try: sheet.unmerge_cells(str(r))
                        except: pass

            ws = wb.worksheets[data_sheet_id]
            clear_merges_in_range(ws, self.config['START_ROW'], self.config['DATA_END_ROW'] + 20)
            ws.add_data_validation(dv_q)

            current_row = self.config['START_ROW']; current_page = 1; data_ptr = 0
            while data_ptr < len(all_extracted_data):
                # 가용 행 수 계산 (DATA_END_ROW까지만 채움)
                rows_left = self.config['DATA_END_ROW'] - current_row + 1
                
                # 만약 공간이 전혀 없으면 새 시트로 전환
                if rows_left <= 0:
                    current_page += 1; ws = self.prepare_next_sheet(wb, data_sheet_id, current_page)
                    clear_merges_in_range(ws, self.config['START_ROW'], self.config['DATA_END_ROW'] + 20)
                    current_row = self.config['START_ROW']; ws.add_data_validation(dv_q)
                    rows_left = self.config['DATA_END_ROW'] - current_row + 1

                # 이번에 기입할 블록 크기 결정 (기본 3행, 공간 부족 시 남은 만큼)
                this_block_size = min(3, rows_left)
                
                # 현재 블록에 들어갈 데이터 (최대 this_block_size개)
                batch = all_extracted_data[data_ptr : data_ptr + this_block_size]
                
                # 블록의 실제 높이(행 수) 결정
                # 데이터가 남았으면 this_block_size만큼 차지, 마지막 데이터라면 데이터 개수만큼만 차지할 수도 있지만
                # 기존 "3개 한 블록" 스타일을 유지하기 위해 가용 공간(this_block_size)을 다 씁니다.
                actual_block_rows = this_block_size
                
                # 테두리 및 서식 적용
                d_height = float(self.config.get('ROW_HEIGHT_DATA', 20.55))
                for r_offset in range(actual_block_rows):
                    r = current_row + r_offset
                    rd = ws.row_dimensions[r]; rd.height = d_height
                    for c in range(1, 14):
                        cell = ws.cell(row=r, column=c)
                        l_s = thin_side; r_s = thin_side
                        t_s = medium_side if r == self.config['START_ROW'] else thin_side
                        b_s = medium_side if r == self.config['DATA_END_ROW'] else thin_side
                        
                        if c == 1: l_s = medium_side
                        if c == 13: r_s = medium_side
                        
                        if c <= 5: # A-E 수직 병합 구역 내부 선 제거
                            if 1 < c < 5: r_s = Side(style=None); l_s = Side(style=None)
                            elif c == 1: r_s = Side(style=None)
                            elif c == 5: l_s = Side(style=None)
                            
                            # 가변 블록 크기에 따른 병합 구역 내부 선 처리
                            if actual_block_rows > 1:
                                if r_offset == 0: b_s = Side(style=None)
                                elif r_offset == actual_block_rows - 1: t_s = Side(style=None)
                                else: t_s = Side(style=None); b_s = Side(style=None)
                                
                        cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

                # A-E 수직 병합 및 데이터 입력 (블록 전체 범위)
                if actual_block_rows > 1:
                    self.safe_merge_cells(ws, start_row=current_row, start_column=1, end_row=current_row + actual_block_rows - 1, end_column=5)
                    self.safe_merge_cells(ws, start_row=current_row, start_column=6, end_row=current_row + actual_block_rows - 1, end_column=6)
                
                # 도면번호/조인트번호 입력 (배치의 첫번째 데이터 기준)
                if batch:
                    self.safe_set_value(ws, ws.cell(row=current_row, column=1).coordinate, batch[0].get('Dwg', ''))
                    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    
                    self.safe_set_value(ws, ws.cell(row=current_row, column=6).coordinate, batch[0].get('Joint', batch[0].get('No', '')))
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center', vertical='center')

                # 개별 데이터 기입
                for i, item in enumerate(batch):
                    r = current_row + i
                    self.safe_set_value(ws, ws.cell(row=r, column=7).coordinate, item.get('Loc', ''), align='center')
                    
                    for val_key, col_idx in [('Ni', 8), ('Cr', 9), ('Mo', 10)]:
                        v = item.get(val_key, 0.0); cell = ws.cell(row=r, column=col_idx)
                        self.safe_set_value(ws, cell.coordinate, v if v > 0 else "", align='center')
                        if v > 0: cell.number_format = '0.00'
                    
                    self.safe_set_value(ws, ws.cell(row=r, column=13).coordinate, item.get('Grade', ''), align='center')
                    cell_l = ws.cell(row=r, column=13); cell_l.font = Font(size=8.5); dv_q.add(cell_l)

                data_ptr += len(batch)
                current_row += actual_block_rows
                self.progress['value'] = 30 + (data_ptr / len(all_extracted_data)) * 65

            # [NEW] 마지막 데이터 기입 직후 다음 행 H~L에 BLANK 표시 (병합 적용)
            try:
                # current_row는 이미 다음 3행 블록의 시작점이므로, 실제 마지막 데이터 r의 다음 행을 계산
                last_data_r = (current_row - actual_block_rows) + len(batch)
                if last_data_r <= self.config['DATA_END_ROW']:
                    # H(8)~L(12) 열 병합 후 BLANK 입력
                    self.safe_merge_cells(ws, start_row=last_data_r, start_column=8, end_row=last_data_r, end_column=12)
                    self.safe_set_value(ws, ws.cell(row=last_data_r, column=8).coordinate, "BLANK")
                    ws.cell(row=last_data_r, column=8).alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row=last_data_r, column=8).font = Font(size=9, bold=True)
            except: pass

            # [RE-NC] 데이터가 없는 빈 칸도 45열까지 3개 행 블록 단위로 테두리/병합 적용
            d_height = float(self.config.get('ROW_HEIGHT_DATA', 20.55))
            while current_row <= self.config['DATA_END_ROW']:
                rows_left = self.config['DATA_END_ROW'] - current_row + 1
                this_block_size = min(3, rows_left)
                
                for r_offset in range(this_block_size):
                    r = current_row + r_offset
                    rd = ws.row_dimensions[r]; rd.height = d_height
                    for c in range(1, 14):
                        cell = ws.cell(row=r, column=c)
                        l_s = thin_side; r_s = thin_side
                        t_s = medium_side if r == self.config['START_ROW'] else thin_side
                        b_s = medium_side if r == self.config['DATA_END_ROW'] else thin_side
                        
                        if c == 1: l_s = medium_side
                        if c == 13: r_s = medium_side
                        
                        if c <= 5:
                            if 1 < c < 5: r_s = Side(style=None); l_s = Side(style=None)
                            elif c == 1: r_s = Side(style=None)
                            elif c == 5: l_s = Side(style=None)
                            
                            if this_block_size > 1:
                                if r_offset == 0: b_s = Side(style=None)
                                elif r_offset == this_block_size - 1: t_s = Side(style=None)
                                else: t_s = Side(style=None); b_s = Side(style=None)
                            
                        cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)
                
                if this_block_size > 1:
                    self.safe_merge_cells(ws, start_row=current_row, start_column=1, end_row=current_row+this_block_size-1, end_column=5)
                    # [NEW] 빈 데이터 구역도 F열 병합
                    self.safe_merge_cells(ws, start_row=current_row, start_column=6, end_row=current_row+this_block_size-1, end_column=6)
                
                dv_q.add(ws.cell(row=current_row, column=13))
                current_row += this_block_size

            # [FORCE] 모든 데이터 시트의 45행 바닥선(A-L) 설정
            data_end_row = int(self.config.get('DATA_END_ROW', 45))
            for idx, s in enumerate(wb.worksheets):
                if s.max_row >= data_end_row:
                    for c in range(1, 14):
                        cell = s.cell(row=data_end_row, column=c)
                        curr_border = cell.border
                        l_s = curr_border.left; r_s = curr_border.right; t_s = curr_border.top
                        
                        if idx == 0: # 갑지
                            if c in [1, 2, 3, 11, 12, 13]: # A, B, C, K, L, M 바닥선 제거
                                b_s = Side(style=None)
                            else: # D~J 는 약한선
                                b_s = thin_side
                        else: # 을지
                            b_s = medium_side
                            
                        cell.border = Border(left=l_s, right=r_s, top=t_s, bottom=b_s)

            # [FORCE] 을지 시트 상단 외곽 테두리 제거 (1행 Top 및 A1-A4 Left, T1-T4 Right)
            for idx, s in enumerate(wb.worksheets):
                if idx > 0: # 을지 시트
                    # 1행 상단 테두리 제거 (T열까지 확장)
                    for c_idx in range(1, 21):
                        cell = s.cell(row=1, column=c_idx)
                        b = cell.border
                        cell.border = Border(left=b.left, right=b.right, top=Side(style=None), bottom=b.bottom)
                    
                    # A1-A4 좌측, M1-M4 우측, T1-T4 우측 테두리 제거
                    for r_idx in range(1, 5):
                        # A열 (좌측)
                        cell_a = s.cell(row=r_idx, column=1)
                        b_a = cell_a.border
                        cell_a.border = Border(left=Side(style=None), right=b_a.right, top=b_a.top, bottom=b_a.bottom)
                        
                        # M열 (우측)
                        cell_m = s.cell(row=r_idx, column=13)
                        b_m = cell_m.border
                        cell_m.border = Border(left=b_m.left, right=Side(style=None), top=b_m.top, bottom=b_m.bottom)
                        
                        # T열 (우측)
                        cell_t = s.cell(row=r_idx, column=20)
                        b_t = cell_t.border
                        cell_t.border = Border(left=b_t.left, right=Side(style=None), top=b_t.top, bottom=b_t.bottom)
                    
                    # [NEW] L3, M3 (페이지 번호 구역) 모든 테두리 제거
                    for c_idx in [12, 13]: # L, M
                        s.cell(row=3, column=c_idx).border = Border()
                    
                    # [NEW] 5행 위쪽 선 진한선으로 (A~M열 범위)
                    for c_idx in range(1, 14): # A~M
                        cell_5 = s.cell(row=5, column=c_idx)
                        b_5 = cell_5.border
                        cell_5.border = Border(left=b_5.left, right=b_5.right, top=medium_side, bottom=b_5.bottom)

            # [FORCE] 을지 시트의 지정 행(17~45) 글꼴 크기 통일
            for idx, s in enumerate(wb.worksheets):
                if idx > 0: # 을지 시트
                    for r_idx in range(17, data_end_row + 1):
                        for c_idx in range(1, 14):
                            cell = s.cell(row=r_idx, column=c_idx)
                            f = cell.font
                            if f:
                                cell.font = Font(name=f.name, size=10, bold=f.bold, italic=f.italic, vertAlign=f.vertAlign, underline=f.underline, strike=f.strike, color=f.color)
                            else:
                                cell.font = Font(name='맑은 고딕', size=10)

            total_p = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                # [NEW] 최종 저장 직전 사용자의 행/열 커스텀 설정을 모든 시트에 다시 한번 적용 (최우선순위)
                ctx = "COVER" if p_idx == 0 else "DATA"
                self.apply_custom_dimensions(s, ctx)
                
                try:
                    if p_idx == 0: # Gapji (Cover)
                        self.safe_merge_cells(s, 3, 14, 3, 20) # N3:T3
                        cell = s["N3"]
                    else: # Eulji (Data)
                        # [FIX] 을지 시트 페이지 번호 가독성 확보를 위해 L3:M3 병합 검토 (또는 M3 단독)
                        try: self.safe_merge_cells(s, 3, 12, 3, 13) # L3:M3
                        except: pass
                        cell = s["L3"] if "L3" in s.merged_cells else s["M3"]
                    
                    cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    cell.font = Font(name='맑은 고딕', size=11, bold=True)
                    # 단어 사이의 간격을 인위적으로 늘림
                    self.safe_set_value(s, cell.coordinate, f"Page    {page_num}    of    {total_p}")
                except Exception as e:
                    self.log(f"페이지 번호 기입 실패 ({p_idx}): {e}")

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"{os.path.splitext(os.path.basename(template_path))[0]}_Unified_{now_str}{os.path.splitext(template_path)[1]}"
            save_path = os.path.join(os.path.dirname(template_path), output_name); wb.save(save_path)
            self.progress['value'] = 100; self.log(f"✨ 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"통합 성적서 생성이 완료되었습니다.\n\n경로: {os.path.dirname(save_path)}\n파일명: {output_name}")
        except Exception as e:
            self.log(f"❌ 오류: {e}")
            traceback.print_exc()
            # [NEW] 사용자에게 친절한 팝업 알림 추가
            error_msg = f"성적서 생성 중 오류가 발생했습니다.\n\n원인: {e}\n\n* 주로 양식 파일의 시트 구조나 병합 상태가 맞지 않을 때 발생합니다. 다른 양식을 시도하거나 설정을 확인해주세요."
            messagebox.showerror("생성 실패", error_msg)
            # End of _run_pmi_process logic
        finally:
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

    def _run_pt_process(self, final_list, template_path):
        """PT 성적서 생성 (1-row-per-data 레이아웃 + ISO 병합)"""
        self.log(f"🚀 PT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        try:
            wb = openpyxl.load_workbook(template_path, keep_vba=True)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            # 갑지 (Cover)
            ws0 = wb.worksheets[0]
            self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False)
            self.force_print_settings(ws0, context="COVER")
            self.apply_custom_dimensions(ws0, "COVER")

            # 을지 (Data)
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]
            ws.title = f"{ws.title[:20]}_001"
            self.add_logos_to_sheet(ws, is_cover=False)
            self.force_print_settings(ws, context="DATA")

            # 헤더 감지 및 시작 행 결정
            start_row = int(self.config.get('PT_START_ROW', 18))
            end_row = int(self.config.get('PT_END_ROW', 37))
            
            # 스타일 설정
            thin_side = Side(style='thin')
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            # [Smart Detection] 템플릿에서 헤더를 찾아 시작 행 자동 보정
            detected_start = -1
            iso_col_idx = 2 # 기본값 B컬럼
            for r in range(1, 40):
                line_cells = [str(ws.cell(row=r, column=c).value).upper() for c in range(1, 15)]
                line_str = "".join(line_cells)
                if (("JOINT" in line_str or "WELD" in line_str) and ("ISO" in line_str or "LINE" in line_str)) or ("DWG" in line_str and "RESULT" in line_str):
                    detected_start = r + 1
                    # ISO 컬럼 위치 동적 확인
                    for c_idx, val in enumerate(line_cells):
                        if any(k in val for k in ["ISO", "LINE", "DWG", "DRAWING"]) and not any(k in val for k in ["JOINT", "WELD"]):
                            iso_col_idx = c_idx + 1
                            break
                    break
            if detected_start > 0: start_row = detected_start

            current_row = start_row
            current_page = 1
            data_ptr = 0
            
            while data_ptr < len(final_list):
                if current_row > end_row:
                    current_page += 1
                    ws = self.prepare_next_sheet(wb, data_sheet_id, current_page)
                    current_row = start_row
                
                item = final_list[data_ptr]
                
                # 데이터 기입 (B, C, D 컬럼 병합 대응)
                self.safe_set_value(ws, ws.cell(row=current_row, column=1).coordinate, item.get('No', ''))
                # ISO Drawing No. (병합된 칸 반영)
                iso_cell = ws.cell(row=current_row, column=iso_col_idx)
                self.safe_set_value(ws, iso_cell.coordinate, item.get('Dwg', ''))
                self.safe_merge_cells(ws, current_row, iso_col_idx, current_row, iso_col_idx + 2)
                
                self.safe_set_value(ws, ws.cell(row=current_row, column=5).coordinate, item.get('Joint', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=6).coordinate, item.get('NPS', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=7).coordinate, item.get('Thk.', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, item.get('Material', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=9).coordinate, item.get('Welder', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=10).coordinate, item.get('WType', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=11).coordinate, item.get('Result', 'Acc'))
                
                # 스타일링
                for c in range(1, 12):
                    cell = ws.cell(row=current_row, column=c)
                    cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    cell.font = Font(name='바탕', size=9)
                    cell.border = thin_border

                data_ptr += 1
                current_row += 1
                self.progress['value'] = (data_ptr / len(final_list)) * 95

            total_p = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                self.apply_custom_dimensions(s, "DATA" if p_idx > 0 else "COVER")
                # 페이지 번호 기입
                try:
                    p_text = f"Page    {page_num}    of    {total_p}"
                    if p_idx == 0: self.safe_set_value(s, 'O35', p_text)
                    else: self.safe_set_value(s, 'V3', p_text)
                except: pass

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"PT_Report_{now_str}.xlsx"
            save_path = os.path.join(os.path.dirname(template_path), output_name)
            wb.save(save_path)
            self.progress['value'] = 100
            self.log(f"✨ PT 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"PT 성적서 생성이 완료되었습니다.\n경로: {os.path.dirname(save_path)}")
        except Exception as e:
            self.log(f"❌ PT 생성 오류: {e}")
            traceback.print_exc()
            messagebox.showerror("생성 실패", f"PT 성적서 생성 중 오류가 발생했습니다: {e}")
        finally:
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

    def _run_rt_process(self, final_list, template_path):
        """RT 성적서 생성 (1-row-per-data 레이아웃)"""
        self.log(f"🚀 RT 성적서 생성 시작 (총 {len(final_list)} 건)...")
        self.progress['value'] = 0
        try:
            wb = openpyxl.load_workbook(template_path, keep_vba=True)
            if len(wb.worksheets) < 1:
                raise ValueError("선택한 템플릿 파일에 시트가 존재하지 않습니다.")

            # Gapji (Cover) Logic (Simplified for RT, using generic SITCO style)
            ws0 = wb.worksheets[0]
            self.add_logos_to_sheet(ws0, is_cover=True, clear_existing=False)
            self.force_print_settings(ws0, context="COVER")
            self.apply_custom_dimensions(ws0, "COVER")

            # Data Sheet Logic (1-row-per-data)
            data_sheet_id = 1 if len(wb.worksheets) >= 2 else 0
            ws = wb.worksheets[data_sheet_id]
            ws.title = f"{ws.title[:20]}_001"
            self.add_logos_to_sheet(ws, is_cover=False)
            self.force_print_settings(ws, context="DATA")
            # RT usually doesn't have the same headers as PMI
            # self.set_eulji_headers(ws) 

            # Configizable or Default RT boundaries
            start_row = int(self.config.get('RT_START_ROW', 17))
            end_row = int(self.config.get('RT_END_ROW', 41))
            rows_per_page = end_row - start_row + 1
            
            current_row = start_row
            current_page = 1
            data_ptr = 0
            
            while data_ptr < len(final_list):
                if current_row > end_row:
                    current_page += 1
                    ws = self.prepare_next_sheet(wb, data_sheet_id, current_page)
                    current_row = start_row
                
                item = final_list[data_ptr]
                
                # Column Data Mapping (standardized layout)
                self.safe_set_value(ws, ws.cell(row=current_row, column=1).coordinate, item.get('No', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=2).coordinate, item.get('Date', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=3).coordinate, item.get('Dwg', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=4).coordinate, item.get('Joint', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=5).coordinate, item.get('Loc', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=6).coordinate, item.get('T', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=7).coordinate, item.get('Mat', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=8).coordinate, item.get('Weld', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=9).coordinate, item.get('Deg', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=10).coordinate, item.get('IQI', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=11).coordinate, item.get('Sens', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=12).coordinate, item.get('Den', ''))
                
                # Defects D1-D15 (Col 13-27)
                for i in range(1, 16):
                    d_val = item.get(f'D{i}', '')
                    # Check symbol mapping
                    self.safe_set_value(ws, ws.cell(row=current_row, column=12 + i).coordinate, d_val)
                
                self.safe_set_value(ws, ws.cell(row=current_row, column=28).coordinate, item.get('Result', 'ACC'))
                self.safe_set_value(ws, ws.cell(row=current_row, column=29).coordinate, item.get('Welder', ''))
                self.safe_set_value(ws, ws.cell(row=current_row, column=30).coordinate, item.get('Remarks', ''))
                
                # Styling for individual row
                for c in range(1, 31):
                    cell = ws.cell(row=current_row, column=c)
                    cell.alignment = Alignment(horizontal='center', vertical='center', shrink_to_fit=True)
                    cell.font = Font(name='바탕', size=9)

                data_ptr += 1
                current_row += 1
                self.progress['value'] = (data_ptr / len(final_list)) * 95

            total_p = len(wb.worksheets)
            for p_idx, s in enumerate(wb.worksheets):
                page_num = p_idx + 1
                self.apply_custom_dimensions(s, "DATA" if p_idx > 0 else "COVER")
                # Page Numbers
                try:
                    if p_idx == 0: cell = s["O35"] if "O35" in s else s.cell(row=35, column=15)
                    else: cell = s["V3"] if "V3" in s else s.cell(row=3, column=22)
                    self.safe_set_value(s, cell.coordinate, f"Page   {page_num}   of   {total_p}")
                except: pass

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"RT_Report_{now_str}.xlsx"
            save_path = os.path.join(os.path.dirname(template_path), output_name)
            wb.save(save_path)
            self.progress['value'] = 100
            self.log(f"✨ RT 완료! 저장됨: {output_name}")
            messagebox.showinfo("성공", f"RT 성적서 생성이 완료되었습니다.\n경로: {os.path.dirname(save_path)}")
        except Exception as e:
            self.log(f"❌ RT 생성 오류: {e}")
            traceback.print_exc()
            messagebox.showerror("생성 실패", f"RT 성적서 생성 중 오류가 발생했습니다: {e}")
        finally:
            for f in glob.glob(os.path.join(tempfile.gettempdir(), "temp_*.png")):
                try: os.remove(f)
                except: pass

if __name__ == "__main__":
    root = tk.Tk()
    PMIReportApp(root)
    root.mainloop()
