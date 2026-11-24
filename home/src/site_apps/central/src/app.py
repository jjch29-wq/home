### VERSION: BUDGET_SYNC_FIXED_FINAL_V8_1WON_BALANCED ###
import mimetypes
import math
import os
from site_apps.central.src.data_sync import get_history_path, get_process_photos_dir
# [OPTIMIZATION] Prevent slow Windows registry scan for mimetypes
if os.name == 'nt':
    mimetypes.init(files=[])

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sys
import subprocess
import os
import time
import datetime
import json
import ctypes
import re
import traceback
import sys
import subprocess
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..')))
from site_apps.central.src.utils.helpers import install_and_import, normalize_id
from site_apps.central.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.central.src.views.tab_config_manager import TabConfigMixin
from site_apps.central.src.views.inout_view import InoutViewMixin
from site_apps.central.src.views.app_dialogs import AppDialogsMixin
# [FIX] Bypass slow pytz timezone loading on network drives (Google Drive)
import builtins
import inspect

_orig_open = builtins.open
class _PytzDummyFile:
    def close(self): pass

def _fast_open(file, mode='r', *args, **kwargs):
    if isinstance(file, str) and 'pytz' in file and 'zoneinfo' in file and mode == 'rb':
        try:
            frame = inspect.currentframe()
            if frame and frame.f_back and frame.f_back.f_code.co_name == 'open_resource':
                if frame.f_back.f_back and frame.f_back.f_back.f_code.co_name == 'resource_exists':
                    return _PytzDummyFile()
        except Exception:
            pass
    return _orig_open(file, mode, *args, **kwargs)

builtins.open = _fast_open

try:
    # Pre-import critical libraries
    pd = install_and_import('pandas')
    np = install_and_import('numpy')
    install_and_import('openpyxl')
    install_and_import('tkcalendar')
    install_and_import('xlsxwriter')
    install_and_import('pillow', 'PIL')
    from PIL import Image, ImageTk
finally:
    builtins.open = _orig_open

from site_apps.central.src.daily_work_report_manager import DailyWorkReportManager
from site_apps.central.src.ndt_billing_tab import NDTCalculatorTab
from site_apps.central.src.daily_work_log_tab import DailyWorkLogTab
import site_apps.central.src.daily_work_report_manager as daily_work_report_manager
print(f"DEBUG: daily_work_report_manager path: {daily_work_report_manager.__file__}")

# Pre-compiled regex for performance
from site_apps.central.src.views.components import *

# Auto-install additional dependencies
from tkcalendar import DateEntry, Calendar

# [FIX] Patch tkcalendar for stability in ko_KR locale / Python 3.14.
# 각 패치를 독립적으로 적용 - 한 패치 실패가 나머지를 무효화하지 않도록 분리.
import tkinter as _tk_ref

# 1) Calendar.__init__ : select_on_nav=False 기본값 + 헤더 너비 조정
try:
    _orig_cal_init = Calendar.__init__
    def _patched_cal_init(self, *args, **kwargs):
        if 'select_on_nav' not in kwargs:
            kwargs['select_on_nav'] = False
        _orig_cal_init(self, *args, **kwargs)
        def widen():
            if hasattr(self, '_header_month'):
                try: self._header_month.configure(width=25)
                except: pass
            if hasattr(self, '_header_year'):
                try: self._header_year.configure(width=10)
                except: pass
        self.after_idle(widen)
        self._widen_headers = widen
    Calendar.__init__ = _patched_cal_init
except Exception:
    pass

# 2) Calendar._display_calendar : 헤더 너비 재조정
try:
    _orig_cal_display = Calendar._display_calendar
    def _patched_cal_display(self, *args, **kwargs):
        _orig_cal_display(self, *args, **kwargs)
        if hasattr(self, '_widen_headers'):
            self.after_idle(self._widen_headers)
    Calendar._display_calendar = _patched_cal_display
except Exception:
    pass

# 3) DateEntry._show_calendar : tkcalendar < 1.6 버전 호환 (1.6.1에서는 없음 - 무시)
try:
    if hasattr(DateEntry, '_show_calendar'):
        _orig_de_show_cal = DateEntry._show_calendar
        def _patched_de_show_cal(self):
            try: self._last_known_date = self.get_date()
            except: self._last_known_date = None
            try: _orig_de_show_cal(self)
            except BaseException: pass
            if hasattr(self, '_calendar'):
                try: self._calendar.configure(select_on_nav=False)
                except: pass
        DateEntry._show_calendar = _patched_de_show_cal
except Exception:
    pass

# --- GLOBAL UTILITY FUNCTIONS MOVED TO utils/helpers.py ---


# 4) DateEntry.drop_down : tkcalendar 1.6.1+ 날짜박스 클릭 시 튕김 방지 (핵심 패치)
try:
    if hasattr(DateEntry, 'drop_down'):
        _orig_drop_down = DateEntry.drop_down
        def _patched_drop_down(self):
            try: self._last_known_date = self.get_date()
            except: self._last_known_date = None
            try:
                _orig_drop_down(self)
                if hasattr(self, '_calendar'):
                    try: self._calendar.configure(select_on_nav=False)
                    except: pass
            except BaseException:
                pass
        DateEntry.drop_down = _patched_drop_down
except Exception:
    pass

# 5) DateEntry._on_b1_press : 클릭 이벤트 처리 중 예외 방지
try:
    if hasattr(DateEntry, '_on_b1_press'):
        _orig_b1_press = DateEntry._on_b1_press
        def _patched_b1_press(self, event):
            try: _orig_b1_press(self, event)
            except BaseException: pass
        DateEntry._on_b1_press = _patched_b1_press
except Exception:
    pass

# 6) DateEntry._on_calendar_selection
try:
    if hasattr(DateEntry, '_on_calendar_selection'):
        _orig_de_on_sel = DateEntry._on_calendar_selection
        def _patched_de_on_sel(self, event):
            try:
                new_date = self._calendar.selection_get()
                if hasattr(self, '_last_known_date') and self._last_known_date == new_date:
                    pass
            except: pass
            _orig_de_on_sel(self, event)
        DateEntry._on_calendar_selection = _patched_de_on_sel
except Exception:
    pass

# 7) DateEntry._setup_style : 스타일 설정 중 update_idletasks 일시 차단
try:
    if hasattr(DateEntry, '_setup_style'):
        _orig_setup_style = DateEntry._setup_style
        def _patched_setup_style(self):
            _orig_update = self.update_idletasks
            self.update_idletasks = lambda: None
            try:
                _orig_setup_style(self)
            except BaseException:
                pass
            finally:
                self.update_idletasks = _orig_update
        DateEntry._setup_style = _patched_setup_style
except Exception:
    pass

# 8) DateEntry.update_idletasks : Python 3.14 BaseException 안전화 (클래스 레벨 영구 적용)
try:
    def _safe_de_update_idletasks(self):
        try:
            _tk_ref.Misc.update_idletasks(self)
        except BaseException:
            pass
    DateEntry.update_idletasks = _safe_de_update_idletasks
except Exception:
    pass

# 9) DateEntry._determine_downarrow_name : Configure/Map 이벤트 콜백 예외 방지
try:
    if hasattr(DateEntry, '_determine_downarrow_name'):
        _orig_det = DateEntry._determine_downarrow_name
        def _patched_det(self, event=None):
            try: _orig_det(self, event)
            except BaseException: pass
        DateEntry._determine_downarrow_name = _patched_det
except Exception:
    pass

# 10) Calendar._setup_style : ThemeChanged 시 예외 방지
try:
    if hasattr(Calendar, '_setup_style'):
        _orig_cal_setup_style = Calendar._setup_style
        def _patched_cal_setup_style(self, event=None):
            try: _orig_cal_setup_style(self, event)
            except BaseException: pass
        Calendar._setup_style = _patched_cal_setup_style
except Exception:
    pass

# 11) Calendar._prev_year / _next_year : 년도 이동
# [FIX] Exception 수준에서만 잡고, 날짜도 롤백.
try:
    if hasattr(Calendar, '_prev_year'):
        _orig_prev_year = Calendar._prev_year
        def _patched_prev_year(self):
            _saved = self._date
            try:
                _orig_prev_year(self)
            except Exception:
                self._date = _saved   # 실패 시 날짜 롤백
        Calendar._prev_year = _patched_prev_year
except Exception:
    pass

try:
    if hasattr(Calendar, '_next_year'):
        _orig_next_year = Calendar._next_year
        def _patched_next_year(self):
            _saved = self._date
            try:
                _orig_next_year(self)
            except Exception:
                self._date = _saved   # 실패 시 날짜 롤백
        Calendar._next_year = _patched_next_year
except Exception:
    pass

# 12) Calendar._prev_month / _next_month : 월 이동도 동일하게 안전화
try:
    if hasattr(Calendar, '_prev_month'):
        _orig_prev_month = Calendar._prev_month
        def _patched_prev_month(self):
            _saved = self._date
            try:
                _orig_prev_month(self)
            except Exception:
                self._date = _saved
        Calendar._prev_month = _patched_prev_month
except Exception:
    pass

try:
    if hasattr(Calendar, '_next_month'):
        _orig_next_month = Calendar._next_month
        def _patched_next_month(self):
            _saved = self._date
            try:
                _orig_next_month(self)
            except Exception:
                self._date = _saved
        Calendar._next_month = _patched_next_month
except Exception:
    pass

# 12.5) tk.Misc.destroy : Python 3.13+ 종료 시 TclError ("can't delete Tcl command") 예외 방지
try:
    _orig_misc_destroy = tk.Misc.destroy
    def _patched_misc_destroy(self):
        try:
            _orig_misc_destroy(self)
        except tk.TclError:
            pass
        except Exception:
            pass
    tk.Misc.destroy = _patched_misc_destroy
except Exception:
    pass

# 13) DateEntry._on_focus_out_cal : 년도/월 네비게이션 버튼 클릭 시 달력 닫힘 방지
# [근본 원인] 년도 버튼(_l_year/_r_year) 클릭 → Calendar에 <FocusOut> 발생
# → focus_get()이 버튼 위젯(DateEntry가 아닌 _top_cal 자식)을 반환
# → 기존 코드: "focus != self(DateEntry)" 이므로 else: _top_cal.withdraw() → 달력 닫힘
# → 사용자 입장: 년도 버튼 누를 때마다 달력이 꺼짐 ("화면 튕김")
# [수정] focus가 _top_cal 내부 위젯으로 이동한 경우 → 달력 유지
try:
    if hasattr(DateEntry, '_on_focus_out_cal'):
        _orig_on_focus_out_cal = DateEntry._on_focus_out_cal
        def _patched_on_focus_out_cal(self, event):
            fw = self.focus_get()
            if fw is not None:
                try:
                    # focus가 _top_cal 내부(년도·월 버튼 등)로 이동한 경우 → 유지
                    if str(fw).startswith(str(self._top_cal)):
                        return
                except Exception:
                    pass
            try:
                _orig_on_focus_out_cal(self, event)
            except BaseException:
                pass
        DateEntry._on_focus_out_cal = _patched_on_focus_out_cal
except Exception:
    pass

# --- Custom Draggable Messagebox Implementation ---

# Injection: Replace standard messagebox methods to enable draggable behavior globally
if not hasattr(messagebox, 'showerror_orig'):
    messagebox.showerror_orig = messagebox.showerror
    messagebox.showwarning_orig = messagebox.showwarning
    messagebox.showinfo_orig = messagebox.showinfo
    
    messagebox.showerror = DraggableMessagebox.showerror
    messagebox.showwarning = DraggableMessagebox.showwarning
    messagebox.showinfo = DraggableMessagebox.showinfo
















        
        # self.wait_window() removed from here to allow caller to set vars before blocking
        # self.wait_window() removed from here to allow caller to set vars before blocking

class MaterialManager(TabConfigMixin, InoutViewMixin, AppDialogsMixin):
    def __init__(self, root):
        self.root = root

        # 공정별 관경에 따른 검사길이 (m) 룩업 테이블
        self.SIZE_LENGTH = {
            '1100A': 3.511,  '1000A': 3.1919, '900A': 2.8727, '850A': 2.7131,
            '800A':  2.5535, '750A':  2.3939, '700A': 2.2343, '650A': 2.0747,
            '600A':  1.9151, '550A':  1.7555, '500A': 1.5959, '450A': 1.4363,
            '400A':  1.2767, '350A':  1.1172, '300A': 1.0006, '250A': 0.8401,
            '200A':  0.6795, '150A':  0.519,  '125A': 0.4392, '100A': 0.3591,
            '80A':   0.2799, '65A':   0.2397, '50A':  0.1901, '40A':  0.1527,
            '32A':   0.1341, '25A':   0.1068, '20A':  0.0855,
        }
        
        # High DPI awareness
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
                
        self.root.title("자재 및 소모품 관리 시스템 (Material Manager)")
        self.root.geometry("1600x900")
        try:
            self.root.state('zoomed') # Maximize on Windows
        except:
            pass
        
        # Configure overall style
        
        # Global Scroll Handler for all canvases in the app
        def _on_global_mousewheel(event):
            try:
                widget = event.widget.winfo_containing(event.x_root, event.y_root)
                if not widget: return
                
                # Do not interfere with native scrolling widgets
                if isinstance(widget, (tk.Text, ttk.Treeview, tk.Listbox)):
                    return
                    
                parent = widget
                while parent:
                    if isinstance(parent, tk.Canvas):
                        # Allow scrolling only if content is taller than the canvas
                        bbox = parent.bbox("all")
                        if bbox:
                            content_height = bbox[3] - bbox[1]
                            if content_height > parent.winfo_height():
                                parent.yview_scroll(int(-1 * (event.delta / 120)), "units")
                        return "break" # [FIX] Prevent scroll from reaching Combobox/Entry under mouse

                    if hasattr(parent, 'master') and parent.master:
                        parent = parent.master
                    else:
                        break
            except Exception:
                pass
            return "break" # Default to break if we are over a canvas-descendant to be safe
                
        self.root.bind_all("<MouseWheel>", _on_global_mousewheel)
        
        # [FIX] Globally disable MouseWheel on Combobox and Entry to prevent unintentional value changes during scrolling
        self.root.bind_class("TCombobox", "<MouseWheel>", lambda e: "break")
        self.root.bind_class("TEntry", "<MouseWheel>", lambda e: "break")
        
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam') # Use 'clam' theme for better grid line visibility
        except:
            pass
            
        self.style.configure(".", font=('Malgun Gothic', 11))
        self.style.configure("Treeview.Heading",
                             font=('Malgun Gothic', 10, 'bold'),
                             background='#eef2f6',
                             foreground='#263238',
                             relief='raised',
                             padding=[6, 4])
        self.style.map("Treeview.Heading",
                       background=[('active', '#dbe3eb'), ('pressed', '#cfd8dc')])
        self.style.configure("Treeview", font=('Malgun Gothic', 10), rowheight=32)
        
        # Detect system background color for tk widgets (Canvas, Text)
        self.theme_bg = self.root.cget('bg')
        if not self.theme_bg or self.theme_bg == 'SystemButtonFace':
             # Fallback/Modern check for Windows
             try:
                 self.theme_bg = self.style.lookup('TFrame', 'background')
             except:
                 self.theme_bg = '#f0f0f0' # Standard Windows gray
        
        # --- 앱 탭별/위젯별 모던 파스텔 아기자기 공통 스타일 정의 ---
        try:
            # 1. 메인 및 서브 탭 디자인 (도톰하고 세련된 라운드 탭 스타일)
            self.style.configure("TNotebook", background=self.theme_bg, tabmargins=[3, 3, 3, 0])
            self.style.configure("TNotebook.Tab",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 padding=[12, 5],
                                 background='#e2e5e8',
                                 foreground='#555555',
                                 borderwidth=1)
            self.style.map("TNotebook.Tab",
                           background=[("selected", "#ffffff"), ("active", "#f0f4f8")],
                           foreground=[("selected", "#1565c0"), ("active", "#0d47a1")])

            # 2. 그룹 상자 (LabelFrame) 제목 및 테두리 정돈
            self.style.configure("TLabelframe", background=self.theme_bg, bordercolor='#cfd8dc')
            self.style.configure("TLabelframe.Label",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 foreground='#37474f',
                                 background=self.theme_bg)

            # 3. 기본 버튼 공통 패딩
            self.style.configure("TButton", font=('Malgun Gothic', 10), padding=[8, 3])

            # 4. Primary 버튼 (산뜻한 소프트 스카이블루 - 작업 추가 등)
            self.style.configure("Primary.TButton",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 background='#e3f2fd',
                                 foreground='#1565c0',
                                 bordercolor='#90caf9',
                                 padding=[12, 4])
            self.style.map("Primary.TButton",
                           background=[('active', '#bbdefb'), ('pressed', '#90caf9')],
                           foreground=[('active', '#0d47a1')])

            # 5. Action / Success 버튼 (산뜻한 에메랄드/그린 - 최종 저장용)
            self.style.configure("Action.TButton",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 background='#e8f5e9',
                                 foreground='#2e7d32',
                                 bordercolor='#a5d6a7',
                                 padding=[14, 5])
            self.style.map("Action.TButton",
                           background=[('active', '#c8e6c9'), ('pressed', '#a5d6a7')],
                           foreground=[('active', '#1b5e20')])

            # 6. Warning 버튼 (따뜻한 피치/오렌지 - 자동산출, 비우기 등)
            self.style.configure("Warning.TButton",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 background='#fff3e0',
                                 foreground='#e65100',
                                 bordercolor='#ffcc80',
                                 padding=[10, 4])
            self.style.map("Warning.TButton",
                           background=[('active', '#ffe0b2'), ('pressed', '#ffcc80')])

            # 7. Info 버튼 (부드러운 라벤더/퍼플 - 수정, 조회 등)
            self.style.configure("Info.TButton",
                                 font=('Malgun Gothic', 10, 'bold'),
                                 background='#ede7f6',
                                 foreground='#4527a0',
                                 bordercolor='#b39ddb',
                                 padding=[10, 4])
            self.style.map("Info.TButton",
                           background=[('active', '#d1c4e9'), ('pressed', '#b39ddb')])

            # 8. Danger 버튼 (부드러운 로즈/소프트레드 - 삭제용)
            self.style.configure("Danger.TButton",
                                 font=('Malgun Gothic', 10),
                                 background='#ffebee',
                                 foreground='#c62828',
                                 bordercolor='#ef9a9a',
                                 padding=[8, 3])
            self.style.map("Danger.TButton",
                           background=[('active', '#ffcdd2'), ('pressed', '#ef9a9a')])

            # 9. Secondary 버튼 (단정한 슬레이트펄 - 보조 도구/안내용)
            self.style.configure("Secondary.TButton",
                                 font=('Malgun Gothic', 10),
                                 background='#f5f7f8',
                                 foreground='#37474f',
                                 bordercolor='#cfd8dc',
                                 padding=[8, 3])
            self.style.map("Secondary.TButton",
                           background=[('active', '#e4e7eb'), ('pressed', '#cfd8dc')])
        except Exception:
            pass

        # Detachable Windows State (pop-up support)
        self.detached_windows = {}
        
        # Determine base directory and bundle directory for portability
        if getattr(sys, 'frozen', False):
            # If running as an executable
            self.app_dir = os.path.dirname(sys.executable)
            self.bundle_dir = getattr(sys, '_MEIPASS', self.app_dir)
            # For exe, use user's Documents folder for config to ensure write permissions
            documents_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'MaterialManager', 'central')
            if not os.path.exists(documents_dir):
                os.makedirs(documents_dir, exist_ok=True)
            self.config_path = os.path.join(documents_dir, 'Material_Manager_Config.json')
        else:
            # 스크립트 실행 모드: src/ 우선, 없으면 ../data/ 탐색
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
            self.bundle_dir = self.app_dir
            self.config_path = os.path.join(self.app_dir, 'Material_Manager_Config.json')
            _db_name = 'Material_Inventory.xlsx'
            _data_dir = os.path.join(os.path.dirname(self.app_dir), 'data')
            _candidates = [
                os.path.join(self.app_dir, _db_name),   # src/Material_Inventory.xlsx
                os.path.join(_data_dir, _db_name),      # data/Material_Inventory.xlsx
            ]
            self.db_path = next((p for p in _candidates if os.path.exists(p)), _candidates[0])

        if getattr(sys, 'frozen', False):
            # exe 실행 모드: exe와 같은 폴더에 DB 저장 (어디서 실행해도 동일)
            self.db_path = os.path.join(self.app_dir, 'Material_Inventory.xlsx')

        # 중앙지사에서 선택한 클라우드 재고 파일을 우선 사용한다.
        from site_apps.central.src.data_sync import get_business_data_dir
        cloud_inventory = os.path.join(get_business_data_dir(), 'Material_Inventory.xlsx')
        if os.path.isfile(cloud_inventory):
            self.db_path = cloud_inventory

        # db_path 확정
        
        self.sites = [] # Initialize site list
        self.daily_units = ['EA', 'CAN', 'SET', 'KG', 'M', '매', 'I/D', 'P,M,I/D', 'M,I/D', 'Point', 'Meter', 'Inch', 'Dia']
        self.users = [
            "계장 박상태", "계장 박재욱", "계장 손주연", "계장 이경재", "계장 이은수", "계장 이준우", "계장 장승대", "계장 한지원",
            "과장 유상훈", "과장 주영광", "대리 김성현", "대리 박만영", "대리 우명광", "대리 윤석현",
            "부장 곽상돈", "부장 김규년", "부장 노현태", "부장 박광복", "부장 이병국", "부장 이봉주", "부장 주동기", "부장 주진철", "부장 형윤철",
            "이사 곽재운", "이사 박홍규", "주임 김성렬", "주임 김진환", "주임 김태현", "주임 황조현",
            "차장 권인경", "차장 김춘호", "차장 박원준"
        ] # Initialize worker/name list
        self.warehouses = [] # Initialize warehouse list
        self.equipments = [] # Initialize equipment list
        self.aux_equipments = ['JIREH MICROBE BTG006', 'COBRA Manual Scanner'] # Initialize auxiliary equipment list
        self.vehicles = [] # Initialize vehicle list
        self.vehicle_inspections = {} # [NEW] Track active inspection widgets
        self.vehicle_boxes = [] # [NEW] Track active inspection widgets for automated save
        self.companies = ['GS네오텍', '다림건설(주)', '한국지역난방공사', '한국지역난방공사 중앙지사'] # Initialize companies list
        self.worktimes = [] # Initialize worktimes list
        self.ot_times = [] # Initialize ot_times list
        self.test_items = [] # [NEW] Initialize test_items
        self.applied_codes = [] # [NEW] Initialize applied_codes
        
        # [NEW] Pre-initialize autocomplete/display lists to prevent AttributeError during load_data/refresh_filters
        self.materials_display_list = []
        self.co_code_list = []
        self.eq_code_list = []
        self.item_name_list = []
        self.class_list = []
        self.spec_list = []
        self.unit_list = []
        self.mfr_list = []
        self.origin_list = []
        self.sn_list = []
        self.model_list = []
        self.supplier_list = []
        self.equipment_suggestions = []
        self.test_methods = [" ", "RT", "PAUT", "UT", "MT", "PT", "PMI", "ETC"]  # [NEW] Initialize test methods
        
        # [NEW] Centralized NDT Consumable Definitions
        self.ndt_groups = {
            'PT약품': ['세척제', '침투제', '현상제', '형광침투제'],
            'MT약품': ['백색페인트', '흑색자분', '형광자분']
        }
        self.ndt_materials_all = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        
        self.budget_sites = []  # [NEW] 공사실행예산서 전용 현장 목록 (budget_df 현장만)
        self.hidden_sites = ["초안", "롯데현장"]  # [NEW] Default hidden sites as requested
        
        self.load_data()
        
        # Centralized list of Carestream films for suggestions
        self.carestream_films = [
            "Carestream AA400-3⅓*12\"",
            "Carestream AA400-3⅓*17\"", 
            "Carestream AA400-4½*12\"",
            "Carestream AA400-10*12\"",
            "Carestream M100-3⅓*12\"",
            "Carestream M100-10*12\"",
            "Carestream M100-14*17\"",
            "Carestream MX125-3⅓*6\"",
            "Carestream MX125-3⅓*12\"",
            "Carestream MX125-4½*12\"",
            "Carestream MX125-10*12\"",
            "Carestream MX125-14*17\"",
            "Carestream T200-3⅓*12\"",
            "Carestream T200-3⅓*17\"",
            "Carestream T200-4½*12\"",
            "Carestream T200-10*12\"",
            "Carestream T200-14*17\""
        ]
        
        # Core draggable widget keys that should never be hidden
        self.CORE_DRAGGABLE_KEYS = [
            'form_box_geometry', 
            'ndt_usage_box_geometry', 
            'rtk_usage_box_geometry', 
            'save_btn_geometry', 
            'workers_box_geometry'
        ]
        
        # Registry for draggable items {config_key: widget_instance}
        self.draggable_items = {}
        self.memos = {} # key -> {'container': container, 'text_widget': text, 'title_entry': entry}
        self.checklists = {} # key -> {'container': container, 'title_entry': entry, 'item_frame': frame, 'items': []}
        self.vehicle_inspections = {} # key -> widget
        self.layout_locked = False
        self.daily_usage_sash_locked = False
        self._last_motion_time = 0 # For performance throttling
        self.is_ready = False  # Suppress saves until fully loaded

        # [CRITICAL FIX] Pre-load configuration synchronously to ensure UI starts in correct state
        self.preload_config_locks()

        self.create_widgets()
        self.update_registration_combos()
        
        # Enable keyboard navigation
        self.setup_keyboard_shortcuts()

        # NDT 변수 초기화 및 추적 바인딩 추가
        self._init_ndt_vars_and_traces()

    def _init_ndt_vars_and_traces(self):
        """NDT 관련 변수(관경, 조인트수)를 생성하고 자동계산 이벤트를 바인딩합니다."""
        # 동적 바인딩에 대비해 명시적으로 StringVar 초기화
        for var_name in ['ndt_report_pipe_var', 'ndt_ori_joint_var', 'ndt_rep_joint_var', 'ndt_ori_qty_var', 'ndt_rep_qty_var', 'ndt_rej_joint_var', 'ndt_inspection_type_var']:
            if not hasattr(self, var_name):
                setattr(self, var_name, tk.StringVar(value=""))

        # 콜백 함수: 관경과 조인트수가 변경되면 수량을 자동계산
        def on_ndt_change(*args):
            try:
                if not hasattr(self, 'cb_daily_test_method'): return
                method = self.cb_daily_test_method.get().strip().upper()
                
                if method not in ('PAUT', 'PT', 'MT', 'RT'): return
                
                pipe_val = getattr(self, 'ndt_report_pipe_var').get().strip().upper()
                
                # 숫자만 입력된 경우 'A'를 붙여서 처리 (예: "600" -> "600A")
                if pipe_val and pipe_val.isdigit():
                    pipe_val += 'A'
                    
                if self.ndt_loc_type_var.get() == '핫태핑':
                    if method not in ('MT', 'PT'): return
                    from site_apps.central.src.hot_tapping import inspection_length
                    length = inspection_length(self.hot_tap_size_var.get(), self.hot_tap_stage_var.get())
                else:
                    from site_apps.central.src.services.inspection_quantity import quantity_per_joint
                    length = quantity_per_joint(method, pipe_val, self.SIZE_LENGTH)
                    if length is None: return
                
                # 원본 조인트수
                ori_j = getattr(self, 'ndt_ori_joint_var').get().strip()
                # 재촬영 조인트수
                rep_j = getattr(self, 'ndt_rep_joint_var').get().strip()
                if self.ndt_loc_type_var.get() == '핫태핑' or method == 'RT':
                    if not ori_j: self.ndt_ori_qty_var.set('')
                    if not rep_j: self.ndt_rep_qty_var.set('')
                
                total_joints = 0.0
                
                if ori_j:
                    try:
                        joints = float(ori_j.replace(',', ''))
                        total_joints += joints
                        getattr(self, 'ndt_ori_qty_var').set(f"{joints * length:g}" if method == 'RT' else f"{joints * length:.4f}")
                    except ValueError: pass
                    
                if rep_j:
                    try:
                        joints = float(rep_j.replace(',', ''))
                        total_joints += joints
                        getattr(self, 'ndt_rep_qty_var').set(f"{joints * length:g}" if method == 'RT' else f"{joints * length:.4f}")
                    except ValueError: pass
                    
                # 메인 폼의 '수량' 칸(ent_daily_test_amount)에 실시간 자동 기입
                if (total_joints > 0 or self.ndt_loc_type_var.get() == '핫태핑' or method == 'RT') and hasattr(self, 'ent_daily_test_amount'):
                    total_qty = total_joints * length
                    if method == 'RT':
                        total_qty = float(self.ndt_ori_qty_var.get() or 0)
                    self.ent_daily_test_amount.delete(0, tk.END)
                    self.ent_daily_test_amount.insert(0, f"{total_qty:g}" if method == 'RT' else f"{total_qty:.4f}")
                    if method == 'RT' and hasattr(self, '_calculate_daily_ndt_fee'):
                        self._calculate_daily_ndt_fee()
                    
            except Exception as e:
                print(f"NDT auto-calc error: {e}")

        # 변수에 추적(trace) 추가
        getattr(self, 'ndt_report_pipe_var').trace_add('write', on_ndt_change)
        getattr(self, 'ndt_ori_joint_var').trace_add('write', on_ndt_change)
        getattr(self, 'ndt_rep_joint_var').trace_add('write', on_ndt_change)
        self.ndt_rej_joint_var.trace_add('write', on_ndt_change)
        for name in ('ndt_loc_type_var', 'hot_tap_stage_var', 'hot_tap_size_var'):
            getattr(self, name).trace_add('write', on_ndt_change)
        self._refresh_ndt_quantity = on_ndt_change

    def preload_config_locks(self):
        """Pre-load critical lock states synchronously before UI creation"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.layout_locked = config.get('layout_locked', False)
                    self.daily_usage_sash_locked = config.get('daily_usage_sash_locked', False)
                    print(f"DEBUG: Pre-loaded locks - Layout: {self.layout_locked}, Sash: {self.daily_usage_sash_locked}")
        except Exception as e:
            print(f"DEBUG: Failed to pre-load locks: {e}")

    def _ensure_canvas_scroll_region(self):
        """Update canvas scroll region based on content height and width (ensures full visibility)"""
        try:
            if hasattr(self, 'entry_canvas') and self.entry_canvas:
                # Remove update_idletasks() to prevent infinite <Configure> loops 
                # Instead, use an after() callback to defer the bbox calculation if needed, or just calculate as-is
                
                def _do_update():
                    if not self.entry_canvas.winfo_exists(): return
                    
                    # Dynamically adjust the canvas window height to accommodate new/removed widgets
                    if hasattr(self, 'entry_inner_frame') and hasattr(self, 'entry_canvas_window'):
                        req_h = self.entry_inner_frame.winfo_reqheight()
                        canvas_h = self.entry_canvas.winfo_height()
                        target_h = max(req_h, canvas_h)
                        self.entry_canvas.itemconfig(self.entry_canvas_window, height=target_h)
                        
                    bbox = self.entry_canvas.bbox("all")
                    if bbox:
                        canvas_h = self.entry_canvas.winfo_height()
                        canvas_w = self.entry_canvas.winfo_width()
                        bottom = max(bbox[3], canvas_h)
                        right = max(bbox[2], canvas_w, 1100)
                        self.entry_canvas.configure(scrollregion=(0, 0, right, bottom))
                        
                # Debounce the update to prevent recursive looping
                if hasattr(self, '_scroll_update_id'):
                    self.entry_canvas.after_cancel(self._scroll_update_id)
                self._scroll_update_id = self.entry_canvas.after(50, _do_update)
                
        except Exception as e:
            print(f"DEBUG: Scroll region update error: {e}")

    def _on_daily_usage_sash_changed(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_daily_usage_sash_changed_impl
        return _on_daily_usage_sash_changed_impl(self, *args, **kwargs)

    def toggle_resolution_lock(self):
        """Toggle window resolution lock"""
        try:
            self.resolution_locked = not self.resolution_locked
            
            if self.resolution_locked:
                self.locked_width = self.root.winfo_width()
                self.locked_height = self.root.winfo_height()
                self.root.resizable(False, False)
                if hasattr(self, 'btn_resolution_lock'):
                    self.btn_resolution_lock.config(text="🔒 해상도 고정됨")
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['resolution_locked'] = True
                self.tab_config['locked_width'] = self.locked_width
                self.tab_config['locked_height'] = self.locked_height
                print(f"Resolution locked at: {self.locked_width}x{self.locked_height}")
            else:
                self.root.resizable(True, True)
                if hasattr(self, 'btn_resolution_lock'):
                    self.btn_resolution_lock.config(text="🔓 해상도 고정")
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['resolution_locked'] = False
                print("Resolution unlocked")
                
            self.save_tab_config(force=True)
        except Exception as e:
            print(f"Error toggling resolution lock: {e}")

    # Removed force_save_config - using save_tab_config(force=True) instead

    def extract_sn_from_model(self, model_name, current_sn):
        """Extract SN from model name if 'S/N.' is present and update current_sn if it's empty"""
        if pd.isna(model_name):
            return model_name, current_sn
        
        model_str = str(model_name)
        if 'S/N.' in model_str:
            parts = model_str.split('S/N.', 1)
            new_model = parts[0].strip()
            # Remove trailing underscore or dot if present before S/N.
            if new_model.endswith('_') or new_model.endswith('.'):
                new_model = new_model[:-1].strip()
            
            extracted_sn = parts[1].strip()
            
            # Use extracted SN if current SN is empty or matches part of it
            if not str(current_sn).strip() or str(current_sn) == 'nan':
                return new_model, extracted_sn
            else:
                # If current SN exists, we still clean up the model name
                return new_model, current_sn
        
        return model_name, current_sn

    def _sync_dataframe_schema(self, *args, **kwargs):
        from site_apps.central.src.models.material_model import _sync_dataframe_schema_impl
        return _sync_dataframe_schema_impl(self, *args, **kwargs)

    def load_data(self, *args, **kwargs):
        if "export" in "load_data" or "excel" in "load_data":
            from site_apps.central.src.services.excel_exporter import load_data_impl
        else:
            from site_apps.central.src.services.data_loader import load_data_impl
        return load_data_impl(self, *args, **kwargs)


    
    def migrate_old_schema(self):
        """Migrate data from old schema to new schema"""
        old_df = self.materials_df.copy()
        self.materials_df = pd.DataFrame(columns=[
            'MaterialID', '회사코드', '관리품번', '품목명', 'SN', '창고',
            '모델명', '규격', '품목군코드', '제조사', '제조국', 
            '가격', '원가', '관리단위', '수량'
        ])
        
        # Headers are already cleaned by the global cleanup in load_data/init

        for _, row in old_df.iterrows():
            new_row = {
                'MaterialID': row.get('MaterialID', ''),
                '회사코드': '',
                '관리품번': row.get('Equipment Code', ''),
                '품목명': row.get('Item Name', row.get('Name', '')),
                'SN': row.get('SN', ''),
                '창고': '',
                '모델명': '',
                '규격': row.get('Specification', ''),
                '품목군코드': '',
                '제조사': row.get('Manufacturer', ''),
                '제조국': '',
                '가격': 0,
                '관리단위': row.get('Unit', 'EA'),
                '수량': row.get('Current Stock', row.get('Initial Stock', 0))
            }
            self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_row])], ignore_index=True)
        
        self.save_data()
        messagebox.showinfo("마이그레이션 완료", "기존 데이터가 새로운 형식으로 변환되었습니다.")

    def register_new_material(self, name, model='', sn='', persist=True, **kwargs):
        """Helper to register a new material in the master list summerly"""
        try:
            # Robustly calculate NEW ID even if mixed types or NaNs exist
            valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna()
            new_id = int(valid_ids.max() + 1) if not valid_ids.empty else 10001
        except:
            new_id = 10001 # Fallback
        
        # Determine defaults
        new_mat = {
            'MaterialID': new_id,
            '회사코드': kwargs.get('co_code', ''),
            '관리품번': '',
            '품목명': name,
            'SN': sn,
            '창고': kwargs.get('warehouse', ''),
            '모델명': model or name,
            '규격': '',
            '품목군코드': '',
            '공급업체': '',
            '제조사': '',
            '제조국': '',
            '가격': 0,
            '원가': 0,
            '관리단위': 'EA',
            '수량': 0,
            '재고하한': 10,
            'Active': 1
        }
        # Update with any explicit kwargs
        for k, v in kwargs.items():
            if k in new_mat: new_mat[k] = v
            
        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_mat])], ignore_index=True)
        if persist:
            self.save_data()
            self.update_material_combo()
        return new_id
        
        self.save_data()
        messagebox.showinfo("마이그레이션 완료", "기존 데이터가 새로운 형식으로 변환되었습니다.")
    
    def enable_autocomplete(self, combobox, values_list_attr=None, values_list=None, prefix_list=None):
        """Standardize to Native Dropdown: Filter values smoothly without UI blocking."""
        if not hasattr(self, '_autocomplete_timers'):
            self._autocomplete_timers = {}

        def perform_filter(event_widget, force_all=False):
            """Filter combobox values based on typed text."""
            typed = str(event_widget.get()).strip()
            
            if not typed and not force_all:
                return  # Don't filter on empty input unless forced
            
            if values_list_attr:
                base_values = getattr(self, values_list_attr, [])
            elif values_list is not None:
                base_values = values_list
            else:
                base_values = []
            
            all_values = (prefix_list + base_values) if prefix_list else base_values
            
            if force_all or not typed:
                new_values = all_values
            elif prefix_list and typed in prefix_list:
                new_values = all_values
            else:
                new_values = [v for v in all_values if typed.lower() in str(v).lower()]
            
            # Update values if changed
            if list(event_widget['values']) != list(new_values):
                event_widget['values'] = new_values

        def on_keyrelease(event):
            """Filter values as user types smoothly."""
            if event.keysym in ("Left", "Right", "Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
                return
            
            # Cancel any previous timer for this combobox
            if combobox in self._autocomplete_timers:
                self.root.after_cancel(self._autocomplete_timers[combobox])
            
            def delayed_filter():
                try:
                    perform_filter(combobox, force_all=False)
                except Exception:
                    pass
                
            self._autocomplete_timers[combobox] = self.root.after(120, delayed_filter)

        combobox.bind('<KeyRelease>', on_keyrelease, add='+')

        # Guard against sticky dropdown after selection
        def _on_selected(e=None):
            combobox._just_selected = True
            def _reset(): combobox._just_selected = False
            combobox.after(300, _reset)
            
        combobox.bind('<<ComboboxSelected>>', _on_selected, add='+')

        def on_focus_in(event):
            try:
                combobox.selection_range(0, tk.END)
            except Exception:
                pass
            perform_filter(combobox, force_all=True)
            
        combobox.bind('<FocusIn>', on_focus_in, add=True)
        
        # Click to show full list
        combobox.bind('<Button-1>', lambda e: perform_filter(combobox, force_all=True), add=True)


    def apply_autocomplete_to_all_comboboxes(self):
        """Map specific comboboxes to their data lists for autocomplete"""
        # Mapping: {widget_attr_name: values_list_attr_name}
        mappings = {
            'cb_material': 'materials_display_list',
            'cb_trans_filter_mat': 'materials_display_list',
            'cb_trans_filter_site': 'sites',
            'cb_budget_site': 'budget_sites',
            'cb_daily_filter_worker': 'users',
            'cb_daily_filter_vehicle': 'vehicles',
            'cb_daily_test_method': 'test_methods',  # Test method list
            'cb_daily_co_code': 'co_code_list',  # [NEW] Company code for MT/PT
            'cb_trans_filter_vehicle': 'vehicles',
            'cb_sales_filter_site': 'sites',
            'cb_filter_co': 'co_code_list',
            'cb_filter_class': 'class_list',
            'cb_filter_mfr': 'mfr_list',
            'cb_filter_name': 'consumable_display_list', # [NEW] Use consumable-only list for Stock View
            'cb_filter_sn': 'sn_list',
            'cb_filter_model': 'model_list',
            'cb_filter_eq': 'eq_code_list',
            'cb_co_code': 'co_code_list',
            'cb_eq_code': 'eq_code_list',
            'cb_item_name': 'item_name_list',
            'cb_eq_name': 'equipments',
            'cb_model': 'model_list',
            'cb_class': 'class_list',
            'cb_spec': 'spec_list',
            'cb_unit': 'unit_list',
            'cb_supplier': 'supplier_list',
            'cb_mfr': 'mfr_list',
            'cb_origin': 'origin_list',
            'cb_trans_site': 'sites',
            'cb_warehouse': 'warehouses',
            'ent_user': 'users',
            # [NEW] Additional comboboxes for autocomplete
            'cb_name': 'users',  # Worker name
            'ent_worktime': 'worktimes',  # Work time
            'cb_vehicle_info': 'vehicles',  # Vehicle info
        }
        
        for widget_attr, list_attr in mappings.items():
            if hasattr(self, widget_attr):
                widget = getattr(self, widget_attr)
                if isinstance(widget, ttk.Combobox):
                    # For filtering, we might want to include '전체' dynamically
                    current_values = list(widget['values'])
                    if '전체' in current_values:
                        self.enable_autocomplete(widget, values_list_attr=list_attr, prefix_list=['전체'])
                    else:
                        self.enable_autocomplete(widget, values_list_attr=list_attr)

    def _safe_format_datetime(self, val, format_str='%Y-%m-%d %H:%M'):
        if pd.isna(val) or val is None or str(val).strip().lower() in ['nan', 'none', '']:
            return ""
        try:
            return pd.to_datetime(val).strftime(format_str)
        except:
            return str(val)

    def clean_nan(self, *args, **kwargs):
        from site_apps.central.src.utils.helpers import clean_nan_impl
        return clean_nan_impl(self, *args, **kwargs)

    def normalize_id(self, val):
        """Robustly normalize IDs: handle NaN, trailing .0, and whitespace."""
        if pd.isna(val) or val == '' or str(val).lower() == 'nan': return ""
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        return s

    def save_data(self, *args, **kwargs):
        if "export" in "save_data" or "excel" in "save_data":
            from site_apps.central.src.services.excel_exporter import save_data_impl
        else:
            from site_apps.central.src.services.data_loader import save_data_impl
        return save_data_impl(self, *args, **kwargs)
            
    def get_base_salaries(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import get_base_salaries_impl
        return get_base_salaries_impl(self, *args, **kwargs)

    def get_material_defaults(self, *args, **kwargs):
        from site_apps.central.src.models.material_model import get_material_defaults_impl
        return get_material_defaults_impl(self, *args, **kwargs)

    def get_expense_defaults(self, actual=False):
        """Extract site expense defaults from settings_df"""
        if not hasattr(self, 'settings_df') or self.settings_df.empty:
            return [
                ("차량유지비", "주유, 수리, 통행, 주차 등", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 200000 if not actual else 6667),
                ("소모품비", "장갑,일회용 작업복외", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 100000 if not actual else 3333),
                ("복리후생비", "생수, 음료 외 기타", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 50000 if not actual else 1667)
            ]
        df = self.settings_df[self.settings_df['Category'] == 'Expense']
        if df.empty:
            return [
                ("차량유지비", "주유, 수리, 통행, 주차 등", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 200000 if not actual else 6667),
                ("소모품비", "장갑,일회용 작업복외", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 100000 if not actual else 3333),
                ("복리후생비", "생수, 음료 외 기타", "N/A", 12 if not actual else 1, "개월" if not actual else "일", 50000 if not actual else 1667)
            ]
        # Return in (cat, cont, ppl, qty, unit, price) format as expected by _add_row_s1
        result = []
        for x in df[['Name', 'Spec', 'Unit', 'Rate']].values:
            name = x[0]
            spec = x[1]
            unit = x[2]
            rate = x[3]
            qty = 12 if not actual else 1
            
            # 사전원가 기준: 개월 단위, 월별 금액
            # 사후원가 기준: 일 단위, 일별 금액 (월 금액 / 30)
            if name == '차량유지비':
                unit = '개월' if not actual else '일'
                rate = 200000 if not actual else 6667
            elif name == '소모품비':
                unit = '개월' if not actual else '일'
                rate = 100000 if not actual else 3333
            elif name == '복리후생비':
                unit = '개월' if not actual else '일'
                rate = 50000 if not actual else 1667
            elif name == 'Se-175':
                unit = '식' if not actual else '일'
                rate = 0 if not actual else 35714
                
            result.append((name, spec, "N/A", qty, unit, rate))
        return result

    def get_outsource_defaults(self):
        """Extract outsource defaults from settings_df"""
        if not hasattr(self, 'settings_df') or self.settings_df.empty:
            return [("케이엔디이", "방사선투과검사", 0, 15000)]
        df = self.settings_df[self.settings_df['Category'] == 'Outsource']
        if df.empty:
            return [("케이엔디이", "방사선투과검사", 0, 15000)]
        return [tuple(x) for x in df[['Name', 'Spec', 'Unit', 'Rate']].values]

    def create_widgets(self):
        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Tab 1: Current Stock (현재 재고 현황)
        self.tab_stock = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stock, text='현재 재고 현황')
        self.setup_stock_tab()

        # Tab 2: Import/Export (데이터 가져오기/내보내기)
        self.tab_import = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_import, text='데이터 가져오기/내보내기')
        self.setup_import_tab()
        
        # Tab 3: Register/Transaction (입출고 관리)
        self.tab_inout = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_inout, text='입출고 관리')
        self.setup_inout_tab()

        # Tab 4: Daily Usage Entry by Site (현장별 일일 사용량 기입 - 시작 기본 탭)
        self.tab_daily_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage, text='현장별 일일 사용량 기입')
        self.setup_daily_usage_tab()
        
        # Tab 5: Daily Usage Query (현장 일일기록 조회 및 관리)
        self.tab_daily_usage_query = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage_query, text='현장 일일기록 조회 및 관리')
        self.setup_daily_usage_query_tab()
        
        # Tab 6: Monthly Usage Entry (월별 집계)
        self.tab_monthly_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_monthly_usage, text='월별 집계')
        self.setup_monthly_usage_tab()

        # Tab 7: Daily Work Log (작업/감독일보)
        self.tab_daily_work_log = DailyWorkLogTab(self.notebook)
        self.tab_daily_work_log.main_app = self
        self.notebook.add(self.tab_daily_work_log, text='작업/감독일보')

        # Tab 8: NDT Billing Calculator (기성 정산 (NDT))
        self.tab_ndt_billing = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_ndt_billing, text='기성 정산 (NDT)')
        self.setup_ndt_billing_tab()

        # Tab 9: Project Execution Budget (공사실행예산서)
        self.tab_budget = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_budget, text='공사실행예산서')
        self.setup_budget_tab()

        # Tab 10: NDT Inspection Report Generator (비파괴검사 성적서/보고서)
        try:
            from site_apps.central.src.views.ndt_report_tab import NDTReportTab
            self.tab_ndt_report = NDTReportTab(self.notebook, main_app=self)
            self.notebook.add(self.tab_ndt_report, text='비파괴검사 성적서/보고서')
        except Exception as e:
            import logging
            logging.getLogger(__name__).error(f"Failed to load NDTReportTab: {e}")
        
        # Select default tab instantly to prevent flicker: 마지막 보던 탭 또는 현장별 일일 사용량 기입
        init_tab = self.tab_daily_usage
        if hasattr(self, 'tab_config') and isinstance(self.tab_config, dict):
            saved_tab_name = self.tab_config.get('selected_tab_text')
            if saved_tab_name:
                for tid in self.notebook.tabs():
                    if self.notebook.tab(tid, "text") == saved_tab_name:
                        init_tab = tid
                        break
        try:
            self.notebook.select(init_tab)
        except:
            self.notebook.select(self.tab_daily_usage)
        
        # Initial view update (Move here to ensure all tabs are ready before refresh)
        self.refresh_inquiry_filters()
        self.update_daily_usage_view()
        
        # Bind tab events (Fixed order, disable drag reordering)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.notebook.bind("<Button-3>", self.show_tab_context_menu)
        
        # Bind main window resize to maintain sash ratios
        self.root.bind("<Configure>", self._on_main_window_resize)
        
        # Load and restore all configuration (geometry, locks, sashes, draggable items)
        self.root.after(100, self.load_tab_config)
        
        # Apply autocomplete to all comboboxes
        self.root.after(200, self.apply_autocomplete_to_all_comboboxes)
        
        # Save tab config on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    def setup_stock_tab(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import setup_stock_tab_impl
        return setup_stock_tab_impl(self, *args, **kwargs)
    
    def show_low_stock(self):
        """Show items with low stock (less than their specific reorder point)"""
        from site_apps.central.src.models.stock_calculation import build_stock_lookups
        stock_lookups = build_stock_lookups(self)
        low_stock_items = []
        for _, mat in self.materials_df.iterrows():
            if mat.get('Active', 1) == 0:
                continue
                
            current = self.calculate_current_stock(mat['MaterialID'], lookups=stock_lookups)
            reorder_point = mat.get('재고하한', 10)
            if pd.isna(reorder_point) or reorder_point <= 0:
                reorder_point = 10 # Default fallback
                
            if pd.notna(current) and current < reorder_point:
                low_stock_items.append((mat.get('품목명', ''), current, reorder_point))
        
        if not low_stock_items:
            messagebox.showinfo("재고 알림", "수량이 10개 미만인 항목이 없습니다.")
        else:
            msg = "다음 항목들의 재고가 부족합니다:\n\n"
            for item, current, reorder in low_stock_items:
                msg += f"• {item}: 현재 {current:g} (필요 수준: {reorder:g})\n"
            messagebox.showwarning("재고 부족", msg)
    
    def select_all_stock(self):
        """Select all items in the stock treeview"""
        all_items = self.stock_tree.get_children()
        self.stock_tree.selection_set(all_items)

    def delete_selected_material(self):
        """선택된 자재 항목들을 재고에서 삭제"""
        selected_items = self.stock_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("선택 오류", "삭제할 품목을 선택해주세요.")
            return
        
        # Confirm deletion
        confirm = messagebox.askyesno("삭제 확인", f"선택한 {len(selected_items)}개의 품목을 재고에서 영구히 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.")
        
        if not confirm:
            return
            
        # Get MaterialIDs to delete
        mat_ids_to_remove = []
        for item in selected_items:
            values = self.stock_tree.item(item, 'values')
            if values:
                # Ensure we match the type of MaterialID in the dataframe
                # [UX IMPROVEMENT] Get hidden MaterialID from the end of values
                mat_ids_to_remove.append(type(self.materials_df['MaterialID'].iloc[0])(values[-1]))
        
        # Soft delete: Set Active=0 instead of removing from materials_df
        if 'Active' not in self.materials_df.columns:
            self.materials_df['Active'] = 1
            
        # Standardize Active column to numeric to avoid pandas TypeError
        self.materials_df['Active'] = pd.to_numeric(self.materials_df['Active'], errors='coerce').fillna(1)
        
        initial_count = len(self.materials_df[self.materials_df['Active'] != 0])
        
        # Apply deletion (set to 0)
        mask = self.materials_df['MaterialID'].isin(mat_ids_to_remove)
        self.materials_df.loc[mask, 'Active'] = 0
        
        final_count = len(self.materials_df[self.materials_df['Active'] != 0])
        removed_count = initial_count - final_count
        
        if removed_count > 0:
            # Save data and update views
            self.save_data()
            self.update_stock_view()
            self.update_material_combo()
            
            # Optional: Clear transactions related to these materials?
            # For now, let's keep them for history unless explicitly asked.
            
            messagebox.showinfo("완료", f"{removed_count}개의 품목이 삭제되었습니다.")
        else:
            messagebox.showwarning("실패", "데이터프레임에서 항목을 삭제하지 못했습니다.")

    def reset_stock_filters(self):
        """Reset all stock filters to default"""
        self.cb_filter_co.set("전체")
        self.cb_filter_class.set("전체")
        self.cb_filter_mfr.set("전체")
        self.cb_filter_name.set("전체")
        self.cb_filter_sn.set("전체")
        self.cb_filter_model.set("전체")
        self.cb_filter_eq.set("전체")
        self.search_var.set("")
        self.update_stock_view()

    def save_material_edits(self, mat_id, new_data):
        """Save edited material data back to the database"""
        # Update DataFrame (Robust matching)
        target_id_str = self.normalize_id(mat_id)
        mask = self.materials_df['MaterialID'].apply(self.normalize_id) == target_id_str
        idx = self.materials_df.index[mask].tolist()
        if not idx:
            messagebox.showerror("오류", "자재를 찾을 수 없습니다.")
            return
            
        for col, val in new_data.items():
            if col in self.materials_df.columns:
                # Check the actual dtype of the column
                col_dtype = self.materials_df[col].dtype
                
                # Handle different data types appropriately
                if col_dtype == 'float64' or col_dtype == 'int64':
                    # Numeric columns - convert to float, empty becomes 0.0
                    try:
                        val = float(val) if str(val).strip() else 0.0
                    except (ValueError, TypeError):
                        val = 0.0
                else:
                    # String/object columns - handle empty values
                    if val == '' or val == 'nan' or pd.isna(val):
                        val = ''
                    else:
                        val = str(val)
                
                self.materials_df.at[idx[0], col] = val
        
        # Re-check SN extraction after edit
        model = self.materials_df.at[idx[0], '모델명']
        sn = self.materials_df.at[idx[0], 'SN']
        new_model, new_sn = self.extract_sn_from_model(model, sn)
        self.materials_df.at[idx[0], '모델명'] = new_model
        self.materials_df.at[idx[0], 'SN'] = new_sn
        
        # Save to Excel
        self.save_data()
        
        # Refresh everything
        self.update_stock_view()
        self.update_registration_combos()
        self.update_material_combo()
        
        messagebox.showinfo("완료", "자재 정보가 성공적으로 수정되었습니다.")

    
    def calculate_current_stock(self, *args, **kwargs):
        from site_apps.central.src.models.material_model import calculate_current_stock_impl
        return calculate_current_stock_impl(self, *args, **kwargs)

    def update_stock_view(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import update_stock_view_impl
        return update_stock_view_impl(self, *args, **kwargs)

    def reload_and_update_stock_view(self, *args, **kwargs):
        """[UX IMPROVEMENT] 엑셀/클라우드 파일을 디스크에서 다시 읽어들인 후 재고 화면을 갱신합니다."""
        try:
            self.load_data()
            if hasattr(self, 'refresh_filters'):
                self.refresh_filters()
        except Exception as e:
            print(f"[ERROR] Failed to reload data from file: {e}")
        return self.update_stock_view(*args, **kwargs)

    def open_detached_stock_view(self):
        """현재 재고 현황을 별도의 팝업창(모니터링 창)으로 엽니다."""
        if hasattr(self, 'detached_windows') and 'stock' in self.detached_windows and self.detached_windows['stock']['window'].winfo_exists():
            self.detached_windows['stock']['window'].lift()
            return
            
        if not hasattr(self, 'detached_windows'):
            self.detached_windows = {}
            
        popup = tk.Toplevel(self.root)
        popup.title("📦 현재 재고 모니터링 (팝업)")
        popup.geometry("1400x800")
        
        self.detached_windows['stock'] = {'window': popup}
        
        main_frame = ttk.Frame(popup, padding=10)
        main_frame.pack(expand=True, fill='both')
        
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(info_frame, text="💡 메인 창의 필터 및 검색어 설정이 이 팝업창에도 실시간으로 반영됩니다.", font=('Malgun Gothic', 10, 'bold'), foreground='#00529B').pack(side='left')
        ttk.Button(info_frame, text="🔄 새로고침", command=self.update_stock_view).pack(side='right')
        
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(expand=True, fill='both')
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        columns = tuple(self.stock_tree['columns'])
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.detached_windows['stock']['tree'] = tree
        
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        
        col_widths = {
            'No.': 40, '회사코드': 80, '분류': 100, '제조사': 100, '품목명': 150, 
            '모델명': 150, 'S/N': 100, '관리단위': 70, 
            '수량(현장)': 80, '단위(현장)': 70, '현재재고': 80,
            '단가': 90, '재고하한': 80, '재고금액': 120, '상태/위치': 100, '관리품번': 120
        }
        for col in columns:
            tree.heading(col, text=self.stock_tree.heading(col, 'text'), command=lambda c=col: self.treeview_sort_column(tree, c, False))
            tree.column(col, width=self.stock_tree.column(col, 'width'), anchor='center', stretch=False)
        tree['displaycolumns'] = self.stock_tree['displaycolumns']
            
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        self.enable_tree_column_drag(tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, tree))
        
        # UI 스타일
        tree.tag_configure('deployed', background='#FFF9C4')
        tree.tag_configure('in_stock', background='')
        
        # 첫 데이터 로딩
        self.update_stock_view()
        
    def treeview_sort_column(self, tv, col, reverse):
        """Standard sorting function for Treeview columns"""
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Try to sort numerically if it looks like a number
        try:
            # Clean values like "1,234" or "EA"
            l.sort(key=lambda t: float(str(t[0]).replace(',', '').strip()) if str(t[0]).replace(',', '').strip() else 0, reverse=reverse)
        except (ValueError, TypeError):
            # Fallback to string sort
            l.sort(reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Switch to reverse sort for next click
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))
        
    def format_worker_summary(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import format_worker_summary_impl
        return format_worker_summary_impl(self, *args, **kwargs)

    def clean_df_export(self, df):
        """Sanitize data for Excel: remove illegal characters, handle empty values, and protect formulas"""
        # 1. Pre-clean common empty representations
        df = df.replace(['nan', 'NaN', 'None'], "")
        
        # 2. Identify columns that have at least one meaningful value
        def is_really_empty(col):
            # A column is empty if all stripped values are "", NaN, or 0 sequences
            return df[col].astype(str).str.strip().replace(['nan', 'None', '', '0', '0.0', '0.00'], pd.NA).dropna().empty

        non_empty_cols = [col for col in df.columns if not is_really_empty(col)]
        
        # Always keep essential columns even if empty
        essential = ['날짜', '현장', '품목명', 'Date', 'Site']
        # Also keep standard RTK and NDT result columns
        rtk_cats = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "RTK총계"]
        ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        essential.extend(rtk_cats)
        essential.extend([f'RTK_{c}' for c in rtk_cats])
        essential.extend(ndt_materials)
        essential.extend([f'NDT_{"".join(m.split())}' for m in ndt_materials])
        
        for col in essential:
            if col in df.columns and col not in non_empty_cols:
                non_empty_cols.append(col)
                
        # 3. Sanitize content and protect formulas
        def sanitize_and_escape(val):
            if pd.isna(val) or val == "" or val == 0 or val == 0.0:
                return ""
            
            # String conversion and cleaning
            s_val = str(val).strip()
            if not s_val or s_val.lower() in ['nan', 'none']:
                return ""

            # Remove illegal XML control characters (ASCII < 32 except tab, LF, CR)
            # These characters cause Excel to report file corruption.
            s_val = "".join(ch for ch in s_val if ord(ch) >= 32 or ch in "\t\n\r")
            
            # Escape potential formula characters (=, +, -, @)
            # IMPORTANT: Skip escaping if the value is a valid number (e.g., -5000)
            if s_val.startswith(('=', '+', '-', '@')):
                try:
                    # If it's a number, don't prefix with '
                    float(s_val)
                    return val
                except ValueError:
                    # It's a text string starting with a formula char
                    return f"'{s_val}"
            
            return s_val

        # Apply sanitization to all columns
        for col in df.columns:
            df[col] = df[col].apply(sanitize_and_escape)
        
        # Return only non-empty columns in original order
        final_cols = [c for c in df.columns if c in non_empty_cols]
        return df[final_cols]

    def save_df_to_excel_autofit(self, *args, **kwargs):
        if "export" in "save_df_to_excel_autofit" or "excel" in "save_df_to_excel_autofit":
            from site_apps.central.src.services.excel_exporter import save_df_to_excel_autofit_impl
        else:
            from site_apps.central.src.services.data_loader import save_df_to_excel_autofit_impl
        return save_df_to_excel_autofit_impl(self, *args, **kwargs)


    
    def setup_import_tab(self):
        import_frame = ttk.LabelFrame(self.tab_import, text="데이터 관리")
        import_frame.pack(pady=20, padx=20, fill='both', expand=True)
        
        # Import Section
        ttk.Label(import_frame, text="엑셀 파일에서 자재 데이터 가져오기", font=('Arial', 11, 'bold')).pack(pady=10)
        ttk.Label(import_frame, text="형식: MaterialID, 회사코드, 관리품번, 품목명, 창고, 모델명, 규격, 품목군코드, 제조사, 제조국, 가격, 관리단위, 수량", 
                 wraplength=600).pack(pady=5)
        
        btn_import = ttk.Button(import_frame, text="엑셀 파일 가져오기", command=self.import_from_excel)
        btn_import.pack(pady=10)
        
        ttk.Separator(import_frame, orient='horizontal').pack(fill='x', pady=20)
        
        # Export Section
        ttk.Label(import_frame, text="현재 데이터 엑셀로 내보내기", font=('Arial', 11, 'bold')).pack(pady=10)
        
        btn_export_materials = ttk.Button(import_frame, text="자재 목록 내보내기", command=self.export_materials)
        btn_export_materials.pack(pady=5)
        
        btn_export_trans = ttk.Button(import_frame, text="거래 내역 내보내기", command=self.export_transactions)
        btn_export_trans.pack(pady=5)
        
        btn_export_all = ttk.Button(import_frame, text="전체 데이터 내보내기", command=self.export_all)
        btn_export_all.pack(pady=5)
    
    def import_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            # Load raw data
            raw_imported_df = pd.read_excel(file_path)
            
            # 1. Standardize column names (mapping) summerly
            col_map = {
                '품목명': ['품목명', '자재명', '품명', 'Item Name', 'Name'],
                'SN': ['SN', 'SN번호', 'Serial Number', 'S/N'],
                '규격': ['규격', 'Specification', 'Spec'],
                '수량': ['수량', 'Initial Stock', 'Current Stock', 'Qty', 'Amount'],
                '모델명': ['모델명', 'Model', 'Machine'],
                '창고': ['창고', 'Warehouse', 'Location'],
                '가격': ['가격', 'Price', 'Unit Price'],
                '제조사': ['제조사', 'Manufacturer', 'Brand'],
                '공급업체': ['공급업체', '공급업자', 'Supplier', 'Vendor'],
                '관리단위': ['관리단위', 'Unit', 'UOM'],
                '재고하한': ['재고하한', '재주문 수준', 'Reorder Point', 'Min Stock']
            }
            
            # Simple column normalization
            working_df = pd.DataFrame()
            for std_name, aliases in col_map.items():
                for alias in aliases:
                    if alias in raw_imported_df.columns:
                        working_df[std_name] = raw_imported_df[alias]
                        break
                if std_name not in working_df.columns:
                    if std_name in ['품목명', '수량']: # Critical columns defaults
                        if std_name == '수량': working_df[std_name] = 0
                        else: 
                            messagebox.showerror("오류", f"필수 컬럼 '{std_name}'을(를) 찾을 수 없습니다.")
                            return
                    else:
                        working_df[std_name] = ""

            # Standardize values: clean strings, handle NaNs summerly
            working_df['품목명'] = working_df['품목명'].astype(str).str.strip()
            working_df['SN'] = working_df['SN'].astype(str).replace('nan', '').str.strip()
            working_df['규격'] = working_df['규격'].astype(str).replace('nan', '').str.strip()
            working_df['수량'] = pd.to_numeric(working_df['수량'], errors='coerce').fillna(0)
            working_df['가격'] = pd.to_numeric(working_df['가격'], errors='coerce').fillna(0)
            
            # [Added] Extract SN from Model name if necessary summerly
            def _extract_row_sn(row):
                m, s = self.extract_sn_from_model(row['모델명'], row['SN'])
                return pd.Series([m, s])
            
            working_df[['모델명', 'SN']] = working_df.apply(_extract_row_sn, axis=1)
            
            # Filter out rows without a name
            working_df = working_df[working_df['품목명'] != 'nan']
            working_df = working_df[working_df['품목명'] != '']
            
            # 2. Interior Deduplication: merge same items within the Excel file itself summerly
            # We group by Name, SN, and Spec.
            final_excel_df = working_df.groupby(['품목명', 'SN', '규격'], as_index=False).agg({
                '수량': 'sum',
                '가격': 'max',
                '모델명': 'first',
                '창고': 'first',
                '제조사': 'first',
                '공급업체': 'first',
                '관리단위': 'first',
                '재고하한': 'first'
            })
            
            # 3. Preparation: Match with existing inventory
            count_new = 0
            count_merged = 0
            count_overwritten = 0
            count_skipped = 0
            
            duplicate_indices = [] # (excel_row_idx, db_row_idx)
            new_items = []
            
            for idx, row in final_excel_df.iterrows():
                mat_name, sn, spec = row['품목명'], row['SN'], row['규격']
                
                existing_idx = -1
                if not self.materials_df.empty:
                    mask = (self.materials_df['품목명'].astype(str) == mat_name) & \
                           (self.materials_df['SN'].astype(str).replace('nan', '') == sn) & \
                           (self.materials_df['규격'].astype(str).replace('nan', '') == spec)
                    matches = self.materials_df.index[mask].tolist()
                    if matches:
                        existing_idx = matches[0]
                
                if existing_idx != -1:
                    duplicate_indices.append((idx, existing_idx))
                else:
                    new_items.append(idx)
            
            # 4. Conflict Resolution Choice Dialog summerly
            mode = "SKIP"
            if duplicate_indices:
                choices_dlg = tk.Toplevel(self.root)
                choices_dlg.title("중복 항목 처리 선택")
                choices_dlg.geometry("450x300")
                choices_dlg.transient(self.root)
                choices_dlg.grab_set()
                
                selection = tk.StringVar(value="CANCEL")
                
                ttk.Label(choices_dlg, text=f"기존 재고 목록과 일치하는 항목 {len(duplicate_indices)}건이 발견되었습니다.", 
                          font=('Malgun Gothic', 10, 'bold')).pack(pady=15)
                ttk.Label(choices_dlg, text="어떻게 처리할까요?").pack(pady=5)
                
                def set_choice(c):
                    selection.set(c)
                    choices_dlg.destroy()
                
                btn_frame = ttk.Frame(choices_dlg, padding=10)
                btn_frame.pack(fill='both', expand=True)
                
                ttk.Button(btn_frame, text="기존 재고에 수량 합치기 (추천)", 
                           command=lambda: set_choice("MERGE"), width=40).pack(pady=5)
                ttk.Button(btn_frame, text="기존 정보 덮어쓰기 (Excel 정보로 교체)", 
                           command=lambda: set_choice("OVERWRITE"), width=40).pack(pady=5)
                ttk.Button(btn_frame, text="중복 항목 건너뛰기 (신규만 추가)", 
                           command=lambda: set_choice("SKIP"), width=40).pack(pady=5)
                ttk.Button(btn_frame, text="가져오기 취소", 
                           command=lambda: set_choice("CANCEL"), width=40).pack(pady=15)
                
                self.root.wait_window(choices_dlg)
                mode = selection.get()
                if mode == "CANCEL":
                    return

            # 5. Process Import based on Mode summerly
            # Add new items first
            new_rows = []
            for idx in new_items:
                row_data = final_excel_df.iloc[idx].to_dict()
                new_mat_id = self.materials_df['MaterialID'].max() + 1 if not self.materials_df.empty else 1
                row_data['MaterialID'] = new_mat_id
                row_data['Active'] = 1
                new_rows.append(row_data)
                count_new += 1
            
            if new_rows:
                self.materials_df = pd.concat([self.materials_df, pd.DataFrame(new_rows)], ignore_index=True)
            
            # Handle duplicates
            if mode != "SKIP":
                for ex_idx, db_idx in duplicate_indices:
                    ex_row = final_excel_df.iloc[ex_idx]
                    
                    if mode == "MERGE":
                        # Add to existing quantity
                        current_qty = pd.to_numeric(self.materials_df.at[db_idx, '수량'], errors='coerce')
                        if pd.isna(current_qty): current_qty = 0
                        self.materials_df.at[db_idx, '수량'] = current_qty + ex_row['수량']
                        # Also update Manufacturer/Supplier if empty in DB but present in Excel
                        for field in ['제조사', '공급업체', '규격', '창고']:
                            if field in self.materials_df.columns and field in ex_row:
                                if not self.materials_df.at[db_idx, field] or str(self.materials_df.at[db_idx, field]) == 'nan':
                                    self.materials_df.at[db_idx, field] = ex_row[field]
                        count_merged += 1
                        
                    elif mode == "OVERWRITE":
                        # Full replace (except ID and Active status)
                        for col in final_excel_df.columns:
                            if col in self.materials_df.columns:
                                self.materials_df.at[db_idx, col] = ex_row[col]
                        count_overwritten += 1
            else:
                count_skipped = len(duplicate_indices)

            # 6. Finalize summerly
            self.save_data()
            self.update_material_combo()
            self.update_stock_view()
            self.update_registration_combos()
            
            summary = f"자재 가져오기가 완료되었습니다.\n\n"
            summary += f"• 신규 등록: {count_new}건\n"
            if mode == "MERGE":
                summary += f"• 재고량 합산: {count_merged}건\n"
            elif mode == "OVERWRITE":
                summary += f"• 정보 덮어쓰기: {count_overwritten}건\n"
            elif mode == "SKIP":
                summary += f"• 중복 건너뜀: {count_skipped}건\n"
            
            messagebox.showinfo("가져오기 완료", summary)
            
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("오류", f"파일을 가져오는데 실패했습니다: {e}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일을 가져오는데 실패했습니다: {e}")
    
    def export_materials(self, *args, **kwargs):
        if "export" in "export_materials" or "excel" in "export_materials":
            from site_apps.central.src.services.excel_exporter import export_materials_impl
        else:
            from site_apps.central.src.services.data_loader import export_materials_impl
        return export_materials_impl(self, *args, **kwargs)
    
    def export_transactions(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Transactions_Export.xlsx",
            title="거래 내역 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                self.transactions_df.to_excel(save_path, index=False)
                messagebox.showinfo("완료", "거래 내역이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    
    def delete_transaction_entry(self):
        """Delete selected transaction from history and refresh"""
        selection = self.inout_tree.selection()
        if not selection:
            messagebox.showwarning("선택 오류", "삭제할 기록을 선택해주세요.")
            return
            
        targets = []
        try:
            for item in selection:
                tags = self.inout_tree.item(item, 'tags')
                if len(tags) < 2 or tags[1] not in ('TRANS', 'DAILY'):
                    raise ValueError('기록의 원본 정보를 확인할 수 없습니다.')
                idx, source = int(tags[0]), tags[1]
                frame = self.transactions_df if source == 'TRANS' else self.daily_usage_df
                if idx not in frame.index:
                    raise ValueError('기록이 변경되었습니다. 새로고침 후 다시 선택해주세요.')
                targets.append((source, idx))
        except (ValueError, TypeError) as error:
            messagebox.showwarning('삭제 오류', str(error), parent=self.root)
            return
        daily_count = len({idx for source, idx in targets if source == 'DAILY'})
        prompt = '선택한 기록을 삭제할까요? 삭제 시 재고 계산에 반영됩니다.'
        if daily_count:
            prompt += (f'\n\n현장 사용 {daily_count}건은 원본 일일 작업 기록 전체가 삭제됩니다.'
                       '\n검사량·작업자·작업시간·메모도 함께 삭제되어 조회 및 집계에 영향을 줍니다.')
        if not messagebox.askyesno('삭제 확인', prompt, parent=self.root):
            return
        old_transactions = self.transactions_df.copy(deep=True)
        old_daily = self.daily_usage_df.copy(deep=True)
        try:
            self.transactions_df = self.transactions_df.drop(
                list({idx for source, idx in targets if source == 'TRANS'}))
            self.daily_usage_df = self.daily_usage_df.drop(
                list({idx for source, idx in targets if source == 'DAILY'}))
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.transactions_df = old_transactions
            self.daily_usage_df = old_daily
            messagebox.showerror('삭제 저장 실패', f'{error}\n화면의 원본 기록을 복원했습니다.', parent=self.root)
            return
        self.update_transaction_view()
        self.update_stock_view()
        
        if daily_count:
            self.update_recent_entries_view()
            self.update_daily_usage_view()
            self.update_monthly_usage_view()
        if hasattr(self, 'budget_view_tree'):
            self.update_budget_site_view()

        messagebox.showinfo("완료", "거래 기록이 삭제되었습니다.")

    def export_transaction_history(self):
        """Export the detailed In/Out history to Excel summerly"""
        history_data = []
        columns = [self.inout_tree.heading(col, 'text') for col in self.inout_tree['columns']]

        for item in self.inout_tree.get_children():
            values = self.inout_tree.item(item, 'values')
            row_data = {}
            for i, col in enumerate(columns):
                row_data[col] = values[i] if i < len(values) else ''
            history_data.append(row_data)

        if not history_data:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return

        # Prepare filename
        today = datetime.datetime.now().strftime('%Y%m%d')
        filename = f"입출고내역_{today}.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            title="입출고 내역 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                history_df = pd.DataFrame(history_data)
                history_df = self.clean_df_export(history_df)
                self.save_df_to_excel_autofit(history_df, save_path, "입출고내역")
                messagebox.showinfo("완료", "입출고 내역이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")


    def _migrate_worktimes(self, input_list):
        """Helper to standardize work times with 익일 marker and sort logically"""
        default_worktimes = [
            "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00", "09:00~22:00",
            "09:00~23:00", "09:00~24:00", "09:00~익일01:00", "09:00~익일02:00", "09:00~익일03:00",
            "09:00~익일04:00", "09:00~익일05:00", "09:00~익일06:00", "09:00~익일07:00", "09:00~익일08:00", "09:00~익일09:00",
            "10:00~17:00", "10:00~19:00", "12:00~익일08:00", "13:00~18:00", "13:00~22:00", "13:00~익일02:00",
            "18:00~익일01:00", "18:00~익일02:00", "18:00~익일03:00", "18:00~익일04:00", "18:00~익일05:00", "18:00~익일06:00", "18:00~익일07:00", "18:00~익일08:00",
            "20:00~익일02:00", "20:00~익일04:00", "20:00~익일05:00", "20:00~익일06:00",
            "21:00~익일03:00", "21:00~익일05:00", "21:00~익일06:00", "21:00~익일07:00",
            "22:00~익일01:00", "22:00~익일02:00", "22:00~익일03:00", "22:00~익일04:00", "22:00~익일05:00", "22:00~익일06:00", "22:00~익일07:00", "22:00~익일08:00", "22:00~익일09:00"
        ]
        
        processed_times = set()
        for t in list(input_list) + default_worktimes:
            if not t or '~' not in t: continue
            
            clean_t = str(t).replace('익일', '').strip()
            try:
                start, end = clean_t.split('~')
                s_h = int(start.split(':')[0])
                e_h = int(end.split(':')[0])
                
                if e_h < s_h and '익일' not in str(t):
                    processed_times.add(f"{start}~익일{end}")
                else:
                    processed_times.add(str(t).strip())
            except:
                processed_times.add(str(t).strip())

        def time_sort_key(t):
            has_ikil = 1 if '익일' in t else 0
            clean = t.replace('익일', '')
            return (has_ikil, clean)

        return sorted(list(processed_times), key=time_sort_key)
        
    def export_all(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Complete_Export.xlsx",
            title="전체 데이터 저장",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    self.materials_df.to_excel(writer, sheet_name='Materials', index=False)
                    self.transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
                messagebox.showinfo("완료", "전체 데이터가 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")
    

    def view_monthly_usage(self):
        """Display monthly usage by item name in the treeview"""
        # Clear current view
        for item in self.usage_tree.get_children():
            self.usage_tree.delete(item)
        
        year = int(self.cb_year.get())
        month = int(self.cb_month.get())
        
        # Filter transactions for the selected month
        month_mask = (self.transactions_df['Date'].dt.year == year) & \
                     (self.transactions_df['Date'].dt.month == month) & \
                     (self.transactions_df['Type'] == 'OUT')
        monthly_trans = self.transactions_df[month_mask]
        
        # Filter transactions for cumulative (from start of year to selected month)
        cumulative_mask = (self.transactions_df['Date'].dt.year == year) & \
                          (self.transactions_df['Date'].dt.month <= month) & \
                          (self.transactions_df['Type'] == 'OUT')
        cumulative_trans = self.transactions_df[cumulative_mask]
        
        # Build usage data for each material
        usage_data = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['MaterialID']
            
            # Calculate monthly usage
            month_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()
            
            # Calculate cumulative usage (year-to-date)
            cumulative_usage = cumulative_trans[cumulative_trans['MaterialID'] == mat_id]['Quantity'].sum()
            
            # Only show items with usage
            if month_usage > 0 or cumulative_usage > 0:
                usage_data.append({
                    '품목명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
                    '관리품번': mat.get('관리품번', ''),
                    '규격': mat.get('규격', ''),
                    '단위': mat.get('관리단위', 'EA'),
                    '월사용량': month_usage,
                    '누계사용량': cumulative_usage
                })
        
        # Sort by item name
        usage_data.sort(key=lambda x: x['품목명'])
        
        # Display in treeview
        for data in usage_data:
            self.usage_tree.insert('', tk.END, values=(
                data['품목명'],
                data['관리품번'],
                data['규격'],
                data['단위'],
                f"{data['월사용량']:.1f}",
                f"{data['누계사용량']:.1f}"
            ))
        
        # Show message if no data
        if not usage_data:
            messagebox.showinfo("알림", f"{year}년 {month}월에 사용 내역이 없습니다.")

    def generate_yearly_report(self):
        year = int(self.cb_year.get())
        # Filter transactions for the year
        mask = (self.transactions_df['Date'].dt.year == year) & (self.transactions_df['Type'] == 'OUT')
        yearly_trans = self.transactions_df[mask]
        
        report = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['MaterialID']
            
            # Calculate monthly usage for each month
            row_data = {
                '설비코드': mat.get('관리품번', ''),
                '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
                '분류': mat.get('품목군코드', ''),
                '규격': mat.get('규격', ''),
                '단위': mat.get('관리단위', ''),
                '제조사': mat.get('제조사', '')
            }
            
            # Add monthly columns (1월 ~ 12월)
            monthly_values = []
            for month in range(1, 13):
                month_mask = (yearly_trans['MaterialID'] == mat_id) & \
                            (yearly_trans['Date'].dt.month == month)
                month_usage = yearly_trans[month_mask]['Quantity'].sum()
                row_data[f'{month}월'] = month_usage
                monthly_values.append(month_usage)
            
            # Calculate totals
            total = sum(monthly_values)
            row_data['합계'] = total
            
            # Calculate cumulative total
            cumulative = 0
            for i, val in enumerate(monthly_values, 1):
                cumulative += val
                if i == 12:  # Only show final cumulative at the end
                    row_data['누계'] = cumulative
            
            report.append(row_data)
            
        report_df = pd.DataFrame(report)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 initialfile=f"Yearly_Usage_{year}.xlsx",
                                                 title="보고서 저장")
        if save_path:
            report_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", f"{year}년 보고서가 저장되었습니다.")

    def generate_monthly_report(self):
        year = int(self.cb_year.get())
        month = int(self.cb_month.get())
        
        mask = (self.transactions_df['Date'].dt.year == year) & \
               (self.transactions_df['Date'].dt.month == month) & \
               (self.transactions_df['Type'] == 'OUT')
        monthly_trans = self.transactions_df[mask]
        
        report = []
        for _, mat in self.materials_df.iterrows():
            mat_id = mat['MaterialID']
            total_usage = monthly_trans[monthly_trans['MaterialID'] == mat_id]['Quantity'].sum()
            report.append({
                '설비코드': mat.get('관리품번', ''),
                '자재명': mat.get('품목명', ''),  # Using 재고현황's 품목명 field
                '분류': mat.get('품목군코드', ''),
                '규격': mat.get('규격', ''),
                '단위': mat.get('관리단위', ''),
                '제조사': mat.get('제조사', ''),
                '월간 총 사용량': total_usage
            })
            
        report_df = pd.DataFrame(report)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 initialfile=f"Monthly_Usage_{year}_{month}.xlsx",
                                                 title="보고서 저장")
        if save_path:
            report_df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", f"{year}년 {month}월 보고서가 저장되었습니다.")

    def setup_monthly_usage_tab(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import setup_monthly_usage_tab_impl
        return setup_monthly_usage_tab_impl(self, *args, **kwargs)
    
    def update_monthly_usage_view(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import update_monthly_usage_view_impl
        return update_monthly_usage_view_impl(self, *args, **kwargs)

    def on_monthly_usage_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_monthly_usage_select_impl
        return on_monthly_usage_select_impl(self, *args, **kwargs)

    def _populate_monthly_summary_trees(self, df, has_note=None):
        """Helper to fill site and worker summary trees with given data subset"""
        # Clear current views
        for item in self.site_summary_tree.get_children():
            self.site_summary_tree.delete(item)
        for item in self.worker_summary_tree.get_children():
            self.worker_summary_tree.delete(item)
            
        # [NEW] Clear detached views
        detached = self.detached_windows.get('monthly')
        if detached:
            for item in detached['site_tree'].get_children(): detached['site_tree'].delete(item)
            for item in detached['worker_tree'].get_children(): detached['worker_tree'].delete(item)
            
        if df.empty:
            return

        # --- Populate Site Summary ---
        # Include MaterialID in grouping and add relevant NDT/Cost fields to aggregation
        site_agg_dict = {
            '검사비': 'sum',
            '출장비': 'sum',
            '제경비': 'sum',
            '기술료': 'sum',
            '환산물량': 'sum',
            '재료비': 'sum',
            '인건비': 'sum',
            'Usage': 'sum',
            '검사량': 'sum'
        }
        
        # Add RTK fields for RT매수 calculation
        rtk_fields = ['RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만', 'RTK_기타']
        for rf in rtk_fields:
            if rf in df.columns: 
                # [DEFENSIVE] Force numeric before aggregation
                df[rf] = pd.to_numeric(df[rf], errors='coerce').fillna(0.0)
                site_agg_dict[rf] = 'sum'
            
        # Add NDT fields
        ndt_map = {
            'NDT_형광자분': 'sum',
            'NDT_자분': 'sum',
            'NDT_흑색자분': 'sum',
            'NDT_페인트': 'sum',
            'NDT_백색페인트': 'sum',
            'NDT_침투제': 'sum',
            'NDT_세척제': 'sum',
            'NDT_현상제': 'sum',
            'NDT_형광': 'sum',
            'NDT_형광침투제': 'sum'
        }
        for nf in ndt_map:
            if nf in df.columns: 
                # [DEFENSIVE] Force numeric before aggregation
                df[nf] = pd.to_numeric(df[nf], errors='coerce').fillna(0.0)
                site_agg_dict[nf] = 'sum'

        # Separate film materials (for RT) while grouping equipment/empty together (for PAUT, MT, PT)
        def get_mat_group(r):
            m = str(r.get('검사방법', '')).strip().upper()
            if m == 'RT':
                mid = r.get('MaterialID', '')
                return '' if pd.isna(mid) else str(mid).strip()
            return ''
            
        df['_MatGroup'] = df.apply(get_mat_group, axis=1)
        
        # Ensure Site and Method are strings and filled
        if 'Site' in df.columns: df['Site'] = df['Site'].fillna('').astype(str)
        if '검사방법' in df.columns: df['검사방법'] = df['검사방법'].fillna('미지정').astype(str)
        
        if 'MaterialID' in site_agg_dict:
            del site_agg_dict['MaterialID']
            
        site_summary = df.groupby(['Site', '검사방법', '_MatGroup'], dropna=False).agg(site_agg_dict).reset_index()
        
        # [NEW] Track active columns for dynamic hiding
        site_cols = ('현장', '검사방법', '품목명', '수량', '단가', '금액', '형광자분', '흑색자분', '백색페인트', 
                     '침투제', '세척제', '현상제', '형광침투제', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계')
        active_cols = set(['현장', '검사방법', '품목명', '수량', '단가', '금액']) # Mandatory columns
        
        # Define a robust active check for numeric data
        def is_active(val):
            if val is None: return False
            s = str(val).strip().lower()
            if s in ('', '0', '0.0', '0.00', 'nan', 'none', '-'):
                return False
            try:
                # Remove common separators and units for numeric check
                v = float(s.replace(',', '').replace('원', '').replace('시간', '').strip())
                return abs(v) > 0.001
            except:
                return bool(s)

        tot_site_qty = 0.0
        tot_site_test_fee = 0.0
        tot_site_black_mag = 0.0
        tot_site_white_paint = 0.0
        tot_site_rtk = 0.0

        for _, row in site_summary.iterrows():
            # Get RT film usage from unified Usage column
            film_total = row.get('Usage', 0.0)
            
            # Calculate values
            black_mag = row.get('NDT_자분', 0.0) + row.get('NDT_흑색자분', 0.0)
            white_paint = row.get('NDT_페인트', 0.0) + row.get('NDT_백색페인트', 0.0)
            pt_pen = row.get('NDT_침투제', 0.0)
            pt_cln = row.get('NDT_세척제', 0.0)
            pt_dev = row.get('NDT_현상제', 0.0)
            pt_fluoro = row.get('NDT_형광', 0.0) + row.get('NDT_형광침투제', 0.0)
            
            r_center = row.get('RTK_센터미스', 0.0)
            r_density = row.get('RTK_농도', 0.0)
            r_marking = row.get('RTK_마킹미스', 0.0)
            r_film = row.get('RTK_필름마크', 0.0)
            r_careless = row.get('RTK_취급부주의', 0.0)
            r_customer = row.get('RTK_고객불만', 0.0)
            r_other = row.get('RTK_기타', 0.0)
            rtk_sum = r_center + r_density + r_marking + r_film + r_careless + r_customer + r_other
            
            qty_v = row.get('검사량', 0.0)
            test_fee_v = row.get('검사비', 0.0)
            unit_price_v = (test_fee_v / qty_v) if (qty_v > 0 and test_fee_v > 0) else 0.0
            
            tot_site_qty += qty_v
            tot_site_test_fee += test_fee_v
            tot_site_black_mag += black_mag
            tot_site_white_paint += white_paint
            tot_site_rtk += rtk_sum

            # Track which columns are active using robust threshold
            if is_active(unit_price_v): active_cols.add('단가')
            if is_active(test_fee_v): active_cols.add('금액')
            if is_active(row.get('NDT_형광자분', 0.0)): active_cols.add('형광자분')
            if is_active(black_mag): active_cols.add('흑색자분')
            if is_active(white_paint): active_cols.add('백색페인트')
            if is_active(pt_pen): active_cols.add('침투제')
            if is_active(pt_cln): active_cols.add('세척제')
            if is_active(pt_dev): active_cols.add('현상제')
            if is_active(pt_fluoro): active_cols.add('형광침투제')
            if is_active(r_center): active_cols.add('센터미스')
            if is_active(r_density): active_cols.add('농도')
            if is_active(r_marking): active_cols.add('마킹미스')
            if is_active(r_film): active_cols.add('필름마크')
            if is_active(r_careless): active_cols.add('취급부주의')
            if is_active(r_customer): active_cols.add('고객불만')
            if is_active(r_other): active_cols.add('기타')
            if is_active(rtk_sum): active_cols.add('RTK총계')
            
            # Get Material Name from _MatGroup
            curr_id = str(row.get('_MatGroup', '')).strip()
            mat_name = ""
            if curr_id and curr_id.lower() not in ('nan', 'none', '0', '0.0', ''):
                try: 
                    m_f = float(curr_id)
                    if m_f != 0.0:
                        matches = self.materials_df[pd.to_numeric(self.materials_df['MaterialID'], errors='coerce') == m_f]
                        if not matches.empty: mat_name = str(matches.iloc[0]['품목명'])
                except: pass
                if not mat_name:
                    matches = self.materials_df[self.materials_df['MaterialID'].astype(str) == str(curr_id)]
                    if not matches.empty: mat_name = str(matches.iloc[0]['품목명'])
                if not mat_name and curr_id not in ('0', '0.0'):
                    mat_name = f"ID:{curr_id}"
                    
            method_str = str(row.get('검사방법', '')).strip().upper()
            if not mat_name:
                if method_str == 'MT':
                    mat_name = "MT약품"
                elif method_str == 'PAUT':
                    mat_name = "-"
                elif method_str == 'PT':
                    mat_name = "PT약품"
                else:
                    mat_name = "-"

            values = (
                row['Site'],
                row['검사방법'],
                mat_name,
                f"{qty_v:.1f}" if is_active(qty_v) else '',
                f"{unit_price_v:,.0f}" if is_active(unit_price_v) else '',
                f"{test_fee_v:,.0f}" if is_active(test_fee_v) else '',
                f"{row.get('NDT_형광자분', 0.0):.1f}" if is_active(row.get('NDT_형광자분', 0.0)) else '',
                f"{black_mag:.1f}" if is_active(black_mag) else '',
                f"{white_paint:.1f}" if is_active(white_paint) else '',
                f"{pt_pen:.1f}" if is_active(pt_pen) else '',
                f"{pt_cln:.1f}" if is_active(pt_cln) else '',
                f"{pt_dev:.1f}" if is_active(pt_dev) else '',
                f"{pt_fluoro:.1f}" if is_active(pt_fluoro) else '',
                f"{r_center:.1f}" if is_active(r_center) else '',
                f"{r_density:.1f}" if is_active(r_density) else '',
                f"{r_marking:.1f}" if is_active(r_marking) else '',
                f"{r_film:.1f}" if is_active(r_film) else '',
                f"{r_careless:.1f}" if is_active(r_careless) else '',
                f"{r_customer:.1f}" if is_active(r_customer) else '',
                f"{r_other:.1f}" if is_active(r_other) else '',
                f"{rtk_sum:.1f}" if is_active(rtk_sum) else ''
            )
            self.site_summary_tree.insert('', tk.END, values=values)
            if detached:
                detached['site_tree'].insert('', tk.END, values=values)

        # Add overall summary row to Site Summary
        if not site_summary.empty:
            self.site_summary_tree.tag_configure('total', background='#E8F4F8', font=('Malgun Gothic', 10, 'bold'))
            overall_site_unit_price = (tot_site_test_fee / tot_site_qty) if (tot_site_qty > 0 and tot_site_test_fee > 0) else 0.0
            total_site_vals = (
                '--- 전체 누계 ---',
                '',
                '',
                f"{tot_site_qty:.1f}" if is_active(tot_site_qty) else '',
                f"{overall_site_unit_price:,.0f}" if is_active(overall_site_unit_price) else '',
                f"{tot_site_test_fee:,.0f}" if is_active(tot_site_test_fee) else '',
                '', # 형광자분
                f"{tot_site_black_mag:.1f}" if is_active(tot_site_black_mag) else '',
                f"{tot_site_white_paint:.1f}" if is_active(tot_site_white_paint) else '',
                '', '', '', '', # 침투제, 세척제, 현상제, 형광침투제
                '', '', '', '', '', '', '', # RTK 세부
                f"{tot_site_rtk:.1f}" if is_active(tot_site_rtk) else ''
            )
            self.site_summary_tree.insert('', tk.END, values=total_site_vals, tags=('total',))
            if detached:
                detached['site_tree'].tag_configure('total', background='#E8F4F8', font=('Malgun Gothic', 10, 'bold'))
                detached['site_tree'].insert('', tk.END, values=total_site_vals, tags=('total',))

        # Apply dynamic column hiding to Site Summary
        visible_sum_cols = [col for col in site_cols if col in active_cols]
        self.site_summary_tree['displaycolumns'] = visible_sum_cols
        if detached:
            detached['site_tree']['displaycolumns'] = visible_sum_cols

        # --- Populate Worker Summary ---
        worker_data = []
        
        # [ROBUST] Use the same precise column detection as update_monthly_usage_view
        def find_paired_cols_local(cols):
            pairs = []
            col_set = set(cols)
            for i in range(1, 11):
                u_n = 'User' if i == 1 else f'User{i}'
                w_n = 'WorkTime' if i == 1 else f'WorkTime{i}'
                o_n = 'OT' if i == 1 else f'OT{i}'
                if u_n in col_set:
                    pairs.append((u_n, w_n if w_n in col_set else None, o_n if o_n in col_set else None))
            return pairs

        df_norm = df.copy()
        df_norm.columns = [str(c).strip().replace(' ', '') for c in df_norm.columns]
        
        pairs = find_paired_cols_local(df_norm.columns)
        for uc, wc, oc in pairs:
            temp_df = pd.DataFrame()
            temp_df['WorkerName'] = df_norm[uc]
            temp_df['ShiftType'] = df_norm[wc] if wc else '주간'
            temp_df['OTValue'] = df_norm[oc] if oc else ''
            
            # Context for deduping
            temp_df['Date'] = df_norm['Date'] if 'Date' in df_norm.columns else None
            temp_df['Site'] = df_norm['Site'] if 'Site' in df_norm.columns else ''
            temp_df['MaterialID'] = df_norm['MaterialID'] if 'MaterialID' in df_norm.columns else ''
            temp_df['WorkTime'] = df_norm[wc] if wc else '' # This worker's own shift
            
            worker_data.append(temp_df)
        
        if worker_data:
            # [NEW] Track active columns for dynamic hiding in Worker Summary
            worker_sum_cols = ('작업자', '총공수', '연장(시간)', '야간(시간)', '휴일(시간)', '총OT(시간)', '연장(금액)', '야간(금액)', '휴일(금액)', '총OT(금액)')
            active_worker_cols = set(['작업자', '총공수'])
            
            worker_df = pd.concat(worker_data)
            worker_df['WorkerName'] = worker_df['WorkerName'].apply(self.clean_nan)
            worker_df = worker_df[worker_df['WorkerName'] != '']
            
            # [NEW] Worker-level Deduping to prevent double-counting in split records
            if not worker_df.empty:
                def make_worker_key(row):
                    d = self._safe_format_datetime(row.get('Date', ''), '%Y-%m-%d')
                    s = str(row.get('Site', '')).strip()
                    w = str(row.get('WorkTime', '')).strip()
                    m = str(row.get('MaterialID', '')).strip()
                    un = str(row.get('WorkerName', '')).strip()
                    return (d, s, w, un)
                
                worker_df['_dedupe_key'] = worker_df.apply(make_worker_key, axis=1)
                # Keep the first instance of a worker in any given activity unit
                worker_df = worker_df.drop_duplicates(subset=['_dedupe_key']).drop(columns=['_dedupe_key'])
            
            if not worker_df.empty:
                from site_apps.central.src.services.ot_calculator import ot_breakdown
                calculated = worker_df.apply(lambda row: ot_breakdown(row['ShiftType'], row['Date'], row['OTValue']), axis=1)
                for column, key in (('H_Day', 'day'), ('H_Night', 'night'), ('H_Holiday', 'holiday'),
                                    ('OT_H', 'hours'), ('A_Day', 'day_amount'), ('A_Night', 'night_amount'),
                                    ('A_Holiday', 'holiday_amount'), ('OT_A', 'amount')):
                    worker_df[column] = calculated.map(lambda result: result[key])
                # [FIXED] Parse fractional work counts from ShiftType (e.g., "0.5주간" -> 0.5)
                def parse_work_count(val):
                    s = str(val).strip()
                    if not s: return 1.0
                    try:
                        # Extract first float/int found in string
                        match = _re.search(r'([0-9\.]+)', s)
                        if match: return float(match.group(1))
                    except: pass
                    return 1.0
                
                worker_df['Count'] = worker_df['ShiftType'].apply(parse_work_count)
                # [REVERTED] Back to entry-based count
                worker_summary = worker_df.groupby('WorkerName').agg({
                    'Count': 'sum',
                    'H_Day': 'sum',
                    'H_Night': 'sum',
                    'H_Holiday': 'sum',
                    'OT_H': 'sum',
                    'A_Day': 'sum',
                    'A_Night': 'sum',
                    'A_Holiday': 'sum',
                    'OT_A': 'sum'
                }).reset_index()
                
                for _, row in worker_summary.iterrows():
                    # Track active columns using robust is_active helper
                    if is_active(row['H_Day']): active_worker_cols.add('연장(시간)')
                    if is_active(row['H_Night']): active_worker_cols.add('야간(시간)')
                    if is_active(row['H_Holiday']): active_worker_cols.add('휴일(시간)')
                    if is_active(row['OT_H']): active_worker_cols.add('총OT(시간)')
                    if is_active(row['A_Day']): active_worker_cols.add('연장(금액)')
                    if is_active(row['A_Night']): active_worker_cols.add('야간(금액)')
                    if is_active(row['A_Holiday']): active_worker_cols.add('휴일(금액)')
                    if is_active(row['OT_A']): active_worker_cols.add('총OT(금액)')
 
                    values = (
                        row['WorkerName'],
                        f"{row['Count']:.1f}" if is_active(row['Count']) else "",
                        f"{row['H_Day']:.1f}" if is_active(row['H_Day']) else "",
                        f"{row['H_Night']:.1f}" if is_active(row['H_Night']) else "",
                        f"{row['H_Holiday']:.1f}" if is_active(row['H_Holiday']) else "",
                        f"{row['OT_H']:.1f}" if is_active(row['OT_H']) else "",
                        f"{row['A_Day']:,.0f}" if is_active(row['A_Day']) else "",
                        f"{row['A_Night']:,.0f}" if is_active(row['A_Night']) else "",
                        f"{row['A_Holiday']:,.0f}" if is_active(row['A_Holiday']) else "",
                        f"{row['OT_A']:,.0f}" if is_active(row['OT_A']) else ""
                    )
                    self.worker_summary_tree.insert('', tk.END, values=values)
                    if detached:
                        detached['worker_tree'].insert('', tk.END, values=values)
                
                # Add overall summary row to Worker Summary
                if not worker_summary.empty:
                    self.worker_summary_tree.tag_configure('total', background='#E8F4F8', font=('Malgun Gothic', 10, 'bold'))
                    total_worker_vals = (
                        '--- 전체 누계 ---',
                        f"{worker_summary['Count'].sum():.1f}" if is_active(worker_summary['Count'].sum()) else "",
                        f"{worker_summary['H_Day'].sum():.1f}" if is_active(worker_summary['H_Day'].sum()) else "",
                        f"{worker_summary['H_Night'].sum():.1f}" if is_active(worker_summary['H_Night'].sum()) else "",
                        f"{worker_summary['H_Holiday'].sum():.1f}" if is_active(worker_summary['H_Holiday'].sum()) else "",
                        f"{worker_summary['OT_H'].sum():.1f}" if is_active(worker_summary['OT_H'].sum()) else "",
                        f"{worker_summary['A_Day'].sum():,.0f}" if is_active(worker_summary['A_Day'].sum()) else "",
                        f"{worker_summary['A_Night'].sum():,.0f}" if is_active(worker_summary['A_Night'].sum()) else "",
                        f"{worker_summary['A_Holiday'].sum():,.0f}" if is_active(worker_summary['A_Holiday'].sum()) else "",
                        f"{worker_summary['OT_A'].sum():,.0f}" if is_active(worker_summary['OT_A'].sum()) else ""
                    )
                    self.worker_summary_tree.insert('', tk.END, values=total_worker_vals, tags=('total',))
                    if detached:
                        detached['worker_tree'].tag_configure('total', background='#E8F4F8', font=('Malgun Gothic', 10, 'bold'))
                        detached['worker_tree'].insert('', tk.END, values=total_worker_vals, tags=('total',))

                # Apply dynamic column hiding to Worker Summary
                visible_worker_cols = [col for col in worker_sum_cols if col in active_worker_cols]
                self.worker_summary_tree['displaycolumns'] = visible_worker_cols
                if detached:
                    detached['worker_tree']['displaycolumns'] = visible_worker_cols

    def apply_worker_shift_hours_to_budget(self):
        """월별 탭의 작업자별 누계(주간/야간/휴일) 시간을 공사실행예산서 특별근무 투입시간에 적용한다."""
        if not hasattr(self, 'actual_labor_detail_widget'):
            messagebox.showwarning("공사탭 미초기화",
                                   "공사실행예산서 탭을 먼저 열어 초기화하세요.")
            return

        # worker_summary_tree에서 모든 행의 주간/야간/휴일 합산 (전체 누계 요약행 제외)
        total_h_day = 0.0
        total_h_night = 0.0
        total_h_holiday = 0.0
        worker_count_day = 0
        worker_count_night = 0
        worker_count_holiday = 0

        def _f(val):
            try: return float(str(val).replace(',', '') or 0)
            except: return 0.0

        for item in self.worker_summary_tree.get_children():
            tags = self.worker_summary_tree.item(item, 'tags')
            if 'total' in tags:
                continue
            vals = self.worker_summary_tree.item(item, 'values')
            if vals and '전체 누계' in str(vals[0]):
                continue
            # 컬럼 순서: 작업자, 총공수, 연장(시간), 야간(시간), 휴일(시간), 총OT(시간), 연장(금액), 야간(금액), 휴일(금액), 총OT(금액)
            h_day     = _f(vals[2]) if len(vals) > 2 else 0.0
            h_night   = _f(vals[3]) if len(vals) > 3 else 0.0
            h_holiday = _f(vals[4]) if len(vals) > 4 else 0.0
            if h_day > 0:
                total_h_day += h_day
                worker_count_day += 1
            if h_night > 0:
                total_h_night += h_night
                worker_count_night += 1
            if h_holiday > 0:
                total_h_holiday += h_holiday
                worker_count_holiday += 1

        if total_h_day == 0 and total_h_night == 0 and total_h_holiday == 0:
            messagebox.showinfo("데이터 없음",
                                "작업자별 누계에 주간/야간/휴일 시간 데이터가 없습니다.\n"
                                "월별 탭에서 조회를 먼저 실행하세요.")
            return

        ldw = self.actual_labor_detail_widget
        for label, count, hours, rate in (
                ('연장근무', worker_count_day, total_h_day, 4000),
                ('야간근무', worker_count_night, total_h_night, 5000),
                ('휴일근무', worker_count_holiday, total_h_holiday, 7500)):
            for field, value in (('personnel', count), ('period', hours / max(count, 1)), ('unit_price', rate)):
                widget = ldw.entries[label][field]
                widget.delete(0, 'end')
                widget.insert(0, f'{value:.15g}')
            ldw._on_input_change(label)

        # 계산 반영
        ldw.calculate_all()

        # 공사실행예산서 탭으로 이동
        try:
            tab_idx = [self.notebook.tab(i, 'text') for i in range(self.notebook.index('end'))].index('공사실행예산서')
            self.notebook.select(tab_idx)
        except:
            pass

        messagebox.showinfo("적용 완료",
                            f"특별근무 투입시간 적용 완료!\n"
                            f"  연장근무: {worker_count_day}명 / {total_h_day:.15g}h (평균 {total_h_day/max(worker_count_day,1):.15g}h/인)\n"
                            f"  야간근무: {worker_count_night}명 / {total_h_night:.15g}h (평균 {total_h_night/max(worker_count_night,1):.15g}h/인)\n"
                            f"  휴일근무: {worker_count_holiday}명 / {total_h_holiday:.15g}h (평균 {total_h_holiday/max(worker_count_holiday,1):.15g}h/인)")

    def export_monthly_usage_history(self):
        """Export monthly usage data, site summaries, and worker summaries to a multi-sheet Excel file"""
        try:
            # 1. Collect Main Monthly Usage Data
            columns = self.monthly_usage_tree['columns']
            monthly_data = []
            for item in self.monthly_usage_tree.get_children():
                monthly_data.append(self.monthly_usage_tree.item(item, 'values'))
            
            if not monthly_data:
                messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
                return

            # 2. Collect Site Summary Data
            site_columns = self.site_summary_tree['columns']
            site_data = []
            for item in self.site_summary_tree.get_children():
                site_data.append(self.site_summary_tree.item(item, 'values'))

            # 3. Collect Worker Summary Data
            worker_columns = self.worker_summary_tree['columns']
            worker_data = []
            for item in self.worker_summary_tree.get_children():
                worker_data.append(self.worker_summary_tree.item(item, 'values'))
            
            # Prepare filename
            today = datetime.datetime.now().strftime('%Y%m%d')
            # Using a descriptive filename indicating it's a comprehensive report (종합)
            filename = f"월별집계_종합_{today}.xlsx"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=filename,
                title="월별 집계(종합) 내역 저장",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if not save_path:
                return

            # --- Process Monthly Usage Tree (Sheet 1) ---
            columns_to_export = [c for c in columns if c != '(Full작업자)']
            col_to_idx = {col: i for i, col in enumerate(columns)}
            
            final_monthly_data = []
            for row in monthly_data:
                new_row = list(row)
                # Swap summarized worker with full list if (Full작업자) exists
                if '작업자' in col_to_idx and '(Full작업자)' in col_to_idx:
                    full_idx = col_to_idx['(Full작업자)']
                    if full_idx < len(new_row):
                        full_val = new_row[full_idx]
                        if full_val:
                            new_row[col_to_idx['작업자']] = full_val
                
                # Filter only display columns
                final_row = [new_row[i] for i, c in enumerate(columns) if c != '(Full작업자)' and i < len(new_row)]
                final_monthly_data.append(final_row)
            
            monthly_df = pd.DataFrame(final_monthly_data, columns=columns_to_export)
            monthly_df = self.clean_df_export(monthly_df)

            # --- Process Site Summary (Sheet 2) ---
            site_df = pd.DataFrame(site_data, columns=site_columns)
            site_df = self.clean_df_export(site_df)

            # --- Process Worker Summary (Sheet 3) ---
            worker_df = pd.DataFrame(worker_data, columns=worker_columns)
            worker_df = self.clean_df_export(worker_df)

            # --- Save to Excel with Multiple Sheets and AutoFit ---
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # Local helper for autofitting to avoid polluting global namespace
                def write_sheet(df, sheet_name):
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                    worksheet = writer.sheets[sheet_name]
                    
                    def get_display_width(s):
                        width = 0
                        for char in str(s):
                            if ord(char) > 127: width += 2 # Double width for Korean
                            else: width += 1
                        return width

                    for idx, col in enumerate(df.columns):
                        series = df[col].astype(str)
                        lengths = series.apply(get_display_width)
                        max_val_len = lengths.max() if not lengths.empty else 0
                        header_len = get_display_width(col)
                        # Cap at 40 to allow better horizontal density in summaries
                        final_width = min(max(max_val_len, header_len) + 2, 40)
                        
                        col_letter = worksheet.cell(row=1, column=idx+1).column_letter
                        worksheet.column_dimensions[col_letter].width = final_width

                write_sheet(monthly_df, "월별집계_내역")
                write_sheet(site_df, "현장별_누계")
                write_sheet(worker_df, "작업자별_누계")

            messagebox.showinfo("완료", "월별 집계 내역(종합)이 저장되었습니다.\n(내역, 현장별, 작업자별 3개 시트 포함)")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")
            traceback.print_exc()

    def create_draggable_container(self, parent, label_text, widget_class, config_key, manage_list_key=None, grid_info=None, **widget_kwargs):
        """Create a draggable container with a label and a widget, styled as a visible box"""
        box_border_color = (
            getattr(self, 'style', None).lookup('TLabelframe', 'bordercolor') if hasattr(self, 'style') else ''
        ) or (
            getattr(self, 'style', None).lookup('TFrame', 'bordercolor') if hasattr(self, 'style') else ''
        ) or '#c0c0c0'

        container = tk.Frame(
            parent,
            relief="solid",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=box_border_color,
            bg=self.theme_bg
        )
        
        # Header container for label and buttons
        hdr = ttk.Frame(container) 
        hdr.pack(side='top', fill='x', padx=0, pady=0)
        
        # Drag handle icon
        lbl_drag = ttk.Label(hdr, text="✥", font=('Arial', 9), cursor='fleur')
        lbl_drag.pack(side='left', padx=1)
        self.make_header_draggable(lbl_drag, container)
        
        # Label
        lbl = ttk.Label(hdr, text=label_text, font=('Malgun Gothic', 9, 'bold'))
        lbl.pack(side='left', padx=1)
        self.make_header_draggable(lbl, container)
        
        # Icons container on the right
        btn_box = ttk.Frame(hdr)
        btn_box.pack(side='right', padx=1)

        # Rename icon
        btn_rename = ttk.Label(btn_box, text="✏️", font=('Arial', 8), cursor='hand2')
        btn_rename.pack(side='left', padx=1)
        btn_rename.bind('<Button-1>', lambda e: self.rename_widget_label(config_key))
        
        # Clone icon
        btn_clone = ttk.Label(btn_box, text="📋", font=('Arial', 8), cursor='hand2')
        btn_clone.pack(side='left', padx=1)
        btn_clone.bind('<Button-1>', lambda e: self.clone_widget(config_key))

        # Delete icon (X)
        btn_del = ttk.Label(btn_box, text="❌", font=('Arial', 8), cursor='hand2')
        btn_del.pack(side='left', padx=1)
        btn_del.bind('<Button-1>', lambda e: self.remove_box(config_key))

        # Internal Content Area
        content_area = ttk.Frame(container, padding=(1, 0))
        content_area.pack(side='top', fill='both', expand=True)

        # Widget
        widget = widget_class(content_area, **widget_kwargs)
        widget.pack(side='left', fill='both', expand=True)

        # Manage List Icon (Gear) - if it's a list-based widget
        if manage_list_key:
            btn_manage = ttk.Label(btn_box, text="⚙️", font=('Arial', 8), cursor='hand2')
            btn_manage.pack(side='left', padx=1)
            # [FIX] Bind AFTER widget is defined
            btn_manage.bind('<Button-1>', lambda e, cb=widget: self.open_list_management_dialog(manage_list_key, target_cb=cb))
        
        # If the widget is a basic tk widget (Text, Canvas), set its background
        if hasattr(widget, 'config') and 'bg' in widget.keys():
            try: widget.config(bg=self.theme_bg)
            except: pass
        
        # NEW: Handle grid placement first if grid_info provided
        if grid_info:
            container.grid(**grid_info)
        
        # Track for layout reset/config
        container._config_key = config_key
        container._label_widget = lbl
        container._widget = widget
        container._widget_class = widget_class
        container._widget_kwargs = widget_kwargs
        container._manage_list_key = manage_list_key
        
        # Register and make draggable (this now captures the CORRECT grid info)
        self.draggable_items[config_key] = container
        self.make_draggable(container, config_key)
        
        return container, widget

    def bulk_register_materials_to_stock(self, material_list):
        """Register items from the preferred list to the main Material inventory if they don't exist or are inactive"""
        print(f"DEBUG: Starting bulk registration for {len(material_list)} items...")
        if not material_list:
            messagebox.showwarning("경고", "등록할 품목이 목록에 없습니다.")
            return
            
        added_count = 0
        reactivated_count = 0
        existing_indices = {} # Name.lower() -> index in materials_df
        max_id = 10000
        
        # Ensure materials_df is loaded
        if self.materials_df is None:
            self.materials_df = pd.DataFrame(columns=['MaterialID', '품목명', 'Active'])

        if not self.materials_df.empty:
            # Collect existing names and their positions
            # [STABILITY] Build index using (Name, Model) to allow parent items even if sub-items exist
            existing_keys = {} # (name_lower, model_lower) -> index
            for idx, mat in self.materials_df.iterrows():
                nm = str(mat.get('품목명', '')).strip().lower()
                md = str(mat.get('모델명', '')).strip().lower()
                if nm:
                    existing_keys[(nm, md)] = idx
            
            try:
                # Calculate next MaterialID
                if 'MaterialID' in self.materials_df.columns:
                    valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna()
                    if not valid_ids.empty:
                        max_id = int(valid_ids.max())
            except Exception as e:
                print(f"DEBUG: Error calculating max_id: {e}")
                max_id = 10000

        new_rows = []
        for mat in material_list:
            mat_clean = str(mat).strip()
            if not mat_clean: continue
            
            mat_lower = mat_clean.lower()
            # Check for parent item (empty model)
            if (mat_lower, '') in existing_keys:
                idx = existing_keys[(mat_lower, '')]
                # Check if it's inactive
                curr_active = self.materials_df.at[idx, 'Active']
                if str(curr_active) == '0' or curr_active == 0:
                    # Reactivate!
                    self.materials_df.at[idx, 'Active'] = 1
                    reactivated_count += 1
                    print(f"DEBUG: Reactivated item: {mat_clean}")
                else:
                    print(f"DEBUG: Skipping '{mat_clean}' (Already Active)")
            else:
                # Add as New
                max_id += 1
                new_row = {
                    'MaterialID': max_id,
                    '품목명': mat_clean,
                    'SN': '',
                    '모델명': '',
                    '관리단위': '매',
                    '수량': 0,
                    '재고하한': 0,
                    'Active': 1,
                    '구분': '소모품'
                }
                for col in self.materials_df.columns:
                    if col not in new_row:
                        new_row[col] = ''
                
                new_rows.append(new_row)
                added_count += 1
                print(f"DEBUG: Prepared new item: {mat_clean} (ID: {max_id})")
        
        if new_rows or reactivated_count > 0:
            try:
                if new_rows:
                    new_df = pd.DataFrame(new_rows)
                    new_df = new_df.reindex(columns=self.materials_df.columns, fill_value='')
                    self.materials_df = pd.concat([self.materials_df, new_df], ignore_index=True)
                
                # Save and Refresh
                self.save_data()
                self.update_stock_view()
                self.update_material_combo()
                
                msg = ""
                if added_count > 0: msg += f"{added_count}개의 품목 신규 등록 완료\n"
                if reactivated_count > 0: msg += f"{reactivated_count}개의 품목 재활성화(복구) 완료"
                
                messagebox.showinfo("등록 완료", msg)
            except Exception as e:
                print(f"DEBUG: Error during bulk registration: {e}")
                messagebox.showerror("오류", f"일괄 등록 중 오류가 발생했습니다:\n{e}")
        else:
            messagebox.showinfo("정보", "이미 모든 품목이 재고 현황에 활성화되어 있습니다.")

    def refresh_ui_for_list_change(self, config_key):
        """Update all related UI elements after a list (sites, users, etc) has changed"""
        # Dictionary mapping config keys to their current values
        list_map = {
            'sites': self.sites,
            'users': self.users,
            'warehouses': self.warehouses,
            'equipments': self.equipments,
            'worktimes': self.worktimes,
            'vehicles': self.vehicles,
            'companies': getattr(self, 'companies', []),
            'daily_units': getattr(self, 'daily_units', [])
        }
        
        if config_key not in list_map:
            return
            
        current_vals = list_map[config_key]
        # [NEW] Ensure uniqueness and strip whitespace before sorting
        unique_vals = sorted(list(set([str(v).strip() for v in current_vals if v])))
        sorted_vals = unique_vals
        
        # [FIX] Update the list in-place to ensure all UI dialogs holding references see the changes
        if isinstance(current_vals, list):
            current_vals[:] = sorted_vals
        
        # 1. Update standard widgets
        if config_key == 'sites':
            if hasattr(self, 'cb_daily_site'): self.cb_daily_site['values'] = sorted_vals
            if hasattr(self, 'cb_trans_site'): self.cb_trans_site['values'] = sorted_vals
            if hasattr(self, 'cb_daily_filter_site'):
                self.cb_daily_filter_site['values'] = ['전체'] + sorted_vals
            if hasattr(self, 'cb_budget_site'): self.cb_budget_site['values'] = sorted_vals
            if hasattr(self, 'cb_budget_view_site'): self.cb_budget_view_site['values'] = sorted_vals
        elif config_key == 'companies':
            # [NEW] Update all company comboboxes in daily usage
            if hasattr(self, 'cb_daily_company'): self.cb_daily_company['values'] = sorted_vals
            if hasattr(self, 'cb_daily_filter_company'):
                self.cb_daily_filter_company['values'] = ['전체'] + sorted_vals
        elif config_key == 'daily_units':
            print(f"DEBUG: Refreshing All Unit Comboboxes with: {sorted_vals}")
            # [FIX] daily_units is already updated in-place above via current_vals[:] = sorted_vals
            # 1. Re-merge to include DB units and update all UI automatically
            self.update_registration_combos()
            # update_registration_combos() handles cb_daily_unit and cb_unit automatically
        elif config_key == 'users':
            # Updated to match current attribute names if needed, 
            # but usually it's cb_daily_user for 1, and cb_daily_user{i} for 2-10
            for i in range(1, 11):
                attr = 'cb_daily_user' if i == 1 else f'cb_daily_user{i}'
                if hasattr(self, attr):
                    widget = getattr(self, attr)
                    if isinstance(widget, WorkerCompositeWidget):
                         widget.cb_name['values'] = [''] + sorted_vals
                    elif hasattr(widget, 'configure'):
                        try: widget['values'] = [''] + sorted_vals
                        except: pass
            
            if hasattr(self, 'ent_user'): self.ent_user['values'] = [''] + sorted_vals
        elif config_key == 'worktimes':
            # Update all worktime fields (1-10)
            for i in range(1, 11):
                attr = f'ent_worktime{i}'
                if hasattr(self, attr):
                    widget = getattr(self, attr)
                    if hasattr(widget, 'configure'):
                        try: 
                            widget['values'] = [''] + sorted_vals
                        except: pass
                grp_attr = f'worker_group{i}'
                if hasattr(self, grp_attr):
                    getattr(self, grp_attr).update_time_list(sorted_vals)
        elif config_key == 'equipments' or config_key == 'materials':
            # Both affect the unified material/equipment suggestion lists
            self.update_material_combo()
        elif config_key == 'test_items' or config_key == 'applied_codes':
            # These are handled by dynamic lambdas in _bind_combobox_word_suggest
            pass 
        elif config_key == 'warehouses':
            if hasattr(self, 'cb_warehouse'): self.cb_warehouse['values'] = sorted_vals
        elif config_key == 'vehicles':
            # 1. Update fixed panel
            if hasattr(self, 'fixed_vehicle_widget'):
                self.fixed_vehicle_widget.update_vehicle_list(sorted_vals)
            # 2. Update floating panels
            if hasattr(self, 'draggable_items'):
                for key, cont in self.draggable_items.items():
                    w = getattr(cont, '_widget', None)
                    if w and type(w).__name__ == 'VehicleInspectionWidget':
                        w.update_vehicle_list(sorted_vals)
            # (Legacy compatibility)
            if hasattr(self, 'vehicle_boxes'):
                for box in getattr(self, 'vehicle_boxes', []):
                    box.update_vehicle_list(sorted_vals)
            if hasattr(self, 'vehicle_inspections'):
                for key, widget_instance in getattr(self, 'vehicle_inspections', {}).items():
                    widget_instance.update_vehicle_list(sorted_vals)
        elif config_key == 'co_code' or config_key == '회사코드':
            # [NEW] Update company code combobox in daily usage
            if hasattr(self, 'cb_daily_co_code'): self.cb_daily_co_code['values'] = sorted_vals

        # Always save configuration after any list change
        self.save_tab_config()


        # 2. Update ALL draggable widgets (clones) that depend on this list
        for key, container in self.draggable_items.items():
            # Heuristic: if manage_list_key is missing but label suggests it's a worker/user/site
            m_key = getattr(container, '_manage_list_key', None)
            if not m_key:
                if hasattr(container, '_label_widget'):
                    lbl_text = container._label_widget.cget('text').lower()
                    if config_key == 'users' and any(x in lbl_text for x in ['작업자', '담당자', 'user', 'worker']):
                        m_key = 'users'
                        container._manage_list_key = 'users'
                    elif config_key == 'sites' and any(x in lbl_text for x in ['현장', 'site']):
                        m_key = 'sites'
                        container._manage_list_key = 'sites'
                    elif config_key == 'equipments' and any(x in lbl_text for x in ['장비', 'equip']):
                        m_key = 'equipments'
                        container._manage_list_key = 'equipments'
                    elif config_key == 'vehicles' and any(x in lbl_text for x in ['차량', 'vehicle']):
                        m_key = 'vehicles'
                        container._manage_list_key = 'vehicles'

            if m_key == config_key:
                if hasattr(container, '_widget') and hasattr(container._widget, 'config'):
                    try:
                        container._widget['values'] = sorted_vals
                    except:
                        pass

    def remove_box(self, key):
        """Intelligently remove a box: hide standard ones, destroy custom ones"""
        panel = getattr(self, 'vehicle_panel', None)
        editor = getattr(self, 'vehicle_inspections', {}).get(key)
        if panel is not None and editor in panel.editors:
            panel.remove(editor)
            return
        is_dynamic = key.startswith('memo_') or key.startswith('clone_') or key.startswith('checklist_') or key.startswith('vehicle_inspection_')
        
        # Allow deletion of dynamic widgets even if layout is locked
        if self.layout_locked and not is_dynamic: return
        
        if is_dynamic:
            # Permanent deletion for custom/dynamic items
            self.destroy_custom_widget(key)
        else:
            # Hiding for standard items (can be restored via Reset All)
            widget = self.draggable_items.get(key)
            if widget:
                # We reuse the existing hide_widget logic
                self.hide_widget(None, widget=widget)

    def destroy_custom_widget(self, key):
        """Destroy and remove from config any dynamic widget (clone or memo)"""
        # Allow destruction even if layout_locked (since it's a content management action)
        
        widget = self.draggable_items.get(key)
        if widget:
            widget.destroy()
            if key in self.draggable_items:
                del self.draggable_items[key]
            if key in self.memos:
                del self.memos[key]
            if key in self.checklists:
                del self.checklists[key]
            if key in self.vehicle_inspections:
                widget_to_remove = self.vehicle_inspections[key]
                if widget_to_remove in self.vehicle_boxes:
                    self.vehicle_boxes.remove(widget_to_remove)
                del self.vehicle_inspections[key]

            
            # Clean from in-memory config to prevent resurrection on next save_tab_config call
            if hasattr(self, 'tab_config') and 'draggable_geometries' in self.tab_config:
                if key in self.tab_config['draggable_geometries']:
                    del self.tab_config['draggable_geometries'][key]
            
            # Clean from config immediately to prevent resurrection on next load
            try:
                import json
                if os.path.exists(self.config_path):
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    
                    if 'draggable_geometries' in config and key in config['draggable_geometries']:
                        del config['draggable_geometries'][key]
                        
                    with open(self.config_path, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
            except:
                pass
            
            self.save_tab_config()

    def rename_widget_label(self, key):
        """Show a simple dialog to rename a widget's label"""
        if self.layout_locked: return
        
        widget = self.draggable_items.get(key)
        if not widget: return
        
        current_text = ""
        if hasattr(widget, '_label_widget'):
            current_text = widget._label_widget.cget('text')
        elif key in self.memos:
            current_text = self.memos[key]['title_entry'].get()
        elif key in self.checklists:
            current_text = self.checklists[key]['title_entry'].get()
            
        new_name = simpledialog.askstring("이름 변경", "새 이름을 입력하세요:", initialvalue=current_text)
        if new_name is not None:
            if hasattr(widget, '_label_widget'):
                widget._label_widget.config(text=new_name)
            elif key in self.memos:
                self.memos[key]['title_entry'].delete(0, 'end')
                self.memos[key]['title_entry'].insert(0, new_name)
            elif key in self.checklists:
                self.checklists[key]['title_entry'].delete(0, 'end')
                self.checklists[key]['title_entry'].insert(0, new_name)
            self.save_tab_config()

    def clone_widget(self, key):
        """Create a clone of an existing widget as a new custom box"""
        if self.layout_locked: return
        
        orig = self.draggable_items.get(key)
        if not orig: return
        
        import time
        new_key = f"clone_{int(time.time() * 1000)}"
        
        label_text = ""
        if hasattr(orig, '_label_widget'):
            label_text = orig._label_widget.cget('text')
        
        if hasattr(orig, '_widget_class'):
            # It's a container created via create_draggable_container
            cont, w = self.create_draggable_container(
                self.entry_inner_frame, 
                label_text, 
                orig._widget_class, 
                new_key, 
                manage_list_key=getattr(orig, '_manage_list_key', None), # Pass manage_list_key
                **orig._widget_kwargs
            )
            
            # Copy value from original widget
            if hasattr(orig, '_widget'):
                try:
                    if hasattr(orig._widget, 'get'):
                        current_val = str(orig._widget.get()) # Ensure string
                        
                        # Try generic Entry-like setting (works for Entry and Combobox text area)
                        if hasattr(w, 'delete') and hasattr(w, 'insert'):
                            try:
                                w.delete(0, 'end')
                                w.insert(0, current_val)
                            except:
                                # Readonly comboboxes might fail delete/insert
                                pass
                                
                        # Try specific set method (Combobox, Scale, etc)
                        if hasattr(w, 'set'):
                            w.set(current_val)
                    elif hasattr(orig._widget, 'get_data') and hasattr(w, 'set_data'):
                        # For complex widgets like VehicleInspectionWidget
                        w.set_data(orig._widget.get_data())
                except Exception as e:
                    print(f"Failed to copy value: {e}")

            
            cont.place(x=50, y=50) # Start position
            self.save_tab_config()

        elif key in self.memos:
            # It's a memo
            content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
            title = self.memos[key]['title_entry'].get()
            self.add_new_memo(initial_text=content, initial_title=title, key=new_key)
            self.save_tab_config()
        elif key in self.checklists:
            # It's a checklist
            self.duplicate_checklist(key)
            self.save_tab_config()

    def _bind_recursive(self, widget, target_container):
        """Recursively bind drag events to widget and its children"""
        from functools import partial
        
        # Bind events to this widget, targeting the container
        # Note: We use add=True to avoid overwriting existing bindings if possible, 
        # but for drag we usually want exclusive control or at least priority.
        # Here we just bind standard.
        
        # We need to capture the target_container in the callback
        widget.bind("<Button-3>", partial(self.on_drag_start, widget=target_container))
        widget.bind("<Shift-Button-3>", partial(self.on_resize_start, widget=target_container))
        widget.bind("<B3-Motion>", partial(self.on_mouse_motion, widget=target_container))
        widget.bind("<Double-Button-3>", partial(self.reset_widget_position, widget=target_container))
        widget.bind("<Control-Button-3>", partial(self.hide_widget, widget=target_container))
        widget.bind("<ButtonRelease-3>", partial(self.on_drag_stop, widget=target_container))
        
        # Recurse for children
        try:
            for child in widget.winfo_children():
                # Toplevel(DateEntry 달력 팝업 등) 내부는 절대 순회하지 않음
                # → winfo_children() 재귀만으로도 Toplevel이 화면에 나타날 수 있음
                if child.winfo_class() == 'Toplevel':
                    continue
                self._bind_recursive(child, target_container)
        except:
            pass

    def hide_widget(self, event, widget=None):
        """Hide a widget (Ctrl + Right Click)"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
            
        # [STABILITY FIX] Prevent hiding core widgets
        if hasattr(widget, '_config_key') and widget._config_key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
            messagebox.showwarning("보호됨", f"'{widget._config_key}' 항목은 화면의 필수 구성 요소이므로 숨길 수 없습니다.")
            return "break"
        
        # Remove from place layout
        if widget.winfo_manager() == 'place':
            widget.place_forget()
        elif widget.winfo_manager() == 'grid':
            widget.grid_forget()
            
        # Remove placeholder if exists
        self._remove_placeholder(widget)
        
        # Mark as hidden in config
        if hasattr(widget, '_config_key') and widget._config_key:
            try:
                import json
                if os.path.exists(self.config_path):
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                else:
                    config = {}
                
                if 'draggable_geometries' not in config:
                    config['draggable_geometries'] = {}
                
                if widget._config_key not in config['draggable_geometries']:
                    config['draggable_geometries'][widget._config_key] = {}
                
                config['draggable_geometries'][widget._config_key]['hidden'] = True
                
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"Failed to save hide status: {e}")
        return "break"

    def make_draggable(self, widget, config_key=None):
        """Make a widget draggable and resizable with right mouse button"""
        widget._config_key = config_key
        # Save original grid info for reset/placeholder
        # We need to do this immediately while it's still in the grid
        widget._original_grid_info = widget.grid_info()
        
        # Recursively bind to ensure clicking anywhere works
        self._bind_recursive(widget, widget)

    def make_header_draggable(self, widget, target_container):
        """Make a specific widget (header/label) draggable with Left Mouse Button targeting a container"""
        from functools import partial
        widget.bind("<Button-1>", partial(self.on_drag_start, widget=target_container))
        widget.bind("<B1-Motion>", partial(self.on_mouse_motion, widget=target_container))
        widget.bind("<ButtonRelease-1>", partial(self.on_drag_stop, widget=target_container))
        
    def reset_widget_position(self, event, widget=None):
        """Reset widget to original grid position"""
        if self.layout_locked:
            return "break"
            
        if widget is None:
            widget = event.widget
        
        # Remove from place layout
        widget.place_forget()
        
        # Remove placeholder if exists
        self._remove_placeholder(widget)
            
        # Restore to grid
        if hasattr(widget, '_original_grid_info'):
            widget.grid(**widget._original_grid_info)
        
        # Reset size variables if any
        if hasattr(widget, '_start_width'): del widget._start_width
        if hasattr(widget, '_start_height'): del widget._start_height
        
        # Remove from config
        if hasattr(widget, '_config_key') and widget._config_key:
            try:
                import json
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                if 'draggable_geometries' in config and widget._config_key in config['draggable_geometries']:
                    del config['draggable_geometries'][widget._config_key]
                
                # Backward compatibility: also check top level
                if widget._config_key in config:
                    del config[widget._config_key]
                    
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                    
            except Exception as e:
                print(f"Failed to reset config: {e}")

    def reset_all_widgets_layout(self):
        """Reset all widgets to their original grid slots, unhide them, and reset paned window sashes"""
        if messagebox.askyesno("초기화", "모든 항목의 위치와 크기를 초기화하고 숨겨진 항목을 다시 표시하시겠습니까?\n(창 분할 위치도 초기화됩니다.)"):
            # 1. Reset draggable widgets
            for key, widget in list(self.draggable_items.items()):
                # Explicitly unhide if it was hidden
                if widget.winfo_manager() == '':
                    # For grid, we just grid it back
                    if hasattr(widget, '_original_grid_info'):
                         widget.grid(**widget._original_grid_info)
                
                self.reset_widget_position(None, widget=widget)
            
            # 1.05 [NEW] Use custom defaults if they exist
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    
                    custom_defaults = config.get('custom_default_geometries', {})
                    if custom_defaults:
                        # Clear existing geometries to ensure fresh start
                        if 'draggable_geometries' not in config: config['draggable_geometries'] = {}
                        
                        for key, geo in custom_defaults.items():
                            if key in self.draggable_items:
                                w = self.draggable_items[key]
                                w.place(x=geo['x'], y=geo['y'], width=geo['width'], height=geo['height'])
                                # Also update config so it persists as the new current layout
                                config['draggable_geometries'][key] = geo
                        
                        # Save the updated config back to disk and memory
                        with open(self.config_path, 'w', encoding='utf-8') as f:
                            json.dump(config, f, ensure_ascii=False, indent=2)
                        self.tab_config = config
                except Exception as e:
                    print(f"Error applying custom defaults: {e}")
                
            # 1.1 Restore parent propagation safely and reset size
            if hasattr(self, 'entry_inner_frame'):
                self.entry_inner_frame.pack_propagate(True)
                self.entry_inner_frame.grid_propagate(True)
                
                # Perform a layout update pass
                self.root.update_idletasks()
                
                # Refresh scrollregion if in canvas
                if hasattr(self, 'entry_canvas'):
                    self.entry_canvas.configure(scrollregion=self.entry_canvas.bbox("all"))
            
            # 2. Reset sash positions (splitters)
            try:
                if hasattr(self, 'daily_usage_paned'):
                    # Give more space to the entry form by default (500px)
                    getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, 500) 
                if hasattr(self, 'daily_history_paned'):
                    total_w = self.daily_history_paned.winfo_width()
                    if total_w > 100:
                        self.daily_history_paned.sashpos(0, int(total_w * 0.7))
                    else:
                        self.daily_history_paned.sashpos(0, 800)
            except:
                pass

            self.save_tab_config()
            messagebox.showinfo("완료", "레이아웃과 창 분할 위치가 초기화되었습니다.")

    def save_current_layout_as_default(self):
        """Save the current layout as the custom 'reset' default"""
        try:
            current_geos = {}
            for key, widget in self.draggable_items.items():
                # [FIX] Capture all managed widgets (grid or place) 
                # This ensures the snapshot is complete even if some items weren't moved.
                manager = widget.winfo_manager()
                if manager in ['place', 'grid']:
                    current_geos[key] = {
                        'x': widget.winfo_x(),
                        'y': widget.winfo_y(),
                        'width': widget.winfo_width(),
                        'height': widget.winfo_height(),
                        'hidden': False
                    }
            
            if not current_geos:
                messagebox.showwarning("실패", "수동으로 배치된(드래그된) 항목이 없어 기본값으로 저장할 수 없습니다.")
                return

            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {}
                
            config['custom_default_geometries'] = current_geos
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            messagebox.showinfo("완료", "현재 레이아웃이 '기본 배치'로 저장되었습니다.\n앞으로 [초기화] 시 이 배치로 되돌아갑니다.")
        except Exception as e:
            messagebox.showerror("오류", f"기본 배치 저장 중 오류 발생: {e}")
            
    def toggle_layout_lock(self):
        """Toggle layout locking state"""
        self.layout_locked = not self.layout_locked
        if hasattr(self, 'btn_lock_layout'):
            if self.layout_locked:
                self.btn_lock_layout.config(text="🔒 배치 고정됨")
                self.style.configure("Lock.TButton", foreground="black")
            else:
                self.btn_lock_layout.config(text="🔓 배치 수정 중")
                self.style.configure("Lock.TButton", foreground="red")
        
        # Save layout lock state to config immediately
        if not hasattr(self, 'tab_config'):
            self.tab_config = {}
        self.tab_config['layout_locked'] = self.layout_locked
        
        # [FIX] Unlocking is no longer destructive. 
        # It simply enables movement without resetting positions.
        if self.layout_locked:
            print("Layout locked - current positions preserved")
        else:
            print("Layout unlocked - movement enabled")
        
        # Force save the state immediately to file
        self.save_tab_config(force=True)
        
        print(f"Layout lock {'enabled' if self.layout_locked else 'disabled'}")

    def on_drag_stop(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_drag_stop_impl
        return on_drag_stop_impl(self, *args, **kwargs)


    def add_new_memo(self, initial_text="", initial_title="메모", key=None):
        """Create a new movable/editable/copyable memo box with editable title"""
        import time
        if key is None:
            key = f"memo_{int(time.time() * 1000)}"
        
        # Create container

        memo_container = ttk.LabelFrame(self.entry_inner_frame)
        
        # Controls frame (top of memo)
        ctrl_frame = ttk.Frame(memo_container)
        ctrl_frame.pack(fill='x', side='top')
        
        # Editable Title Entry
        title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
        title_entry.pack(side='left', padx=2)
        title_entry.insert(0, initial_title)
        title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_memo(key))
        btn_copy.pack(side='right')
        
        btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
        btn_del.pack(side='right')
        
        # Text area
        text_area = tk.Text(memo_container, wrap='word', height=5, width=30, font=('Arial', 10), bg=self.theme_bg, highlightthickness=0)
        text_area.pack(fill='both', expand=True, padx=2, pady=2)
        text_area.insert('1.0', initial_text)
        
        # Bind text change to auto-save
        text_area.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        # Make draggable (Right Click)
        self.make_draggable(memo_container, key)
        
        # Make header draggable (Left Click)
        self.make_header_draggable(ctrl_frame, memo_container)
        self.draggable_items[key] = memo_container
        self.memos[key] = {
            'container': memo_container, 
            'text_widget': text_area,
            'title_entry': title_entry
        }
        
        # Initial placement if not loaded from config (will be handled by load_tab_config if exists)
        if key not in getattr(self, '_loading_memos', []):
            memo_container.place(x=1450, y=50) # Default start position
        
        return memo_container

    def duplicate_memo(self, key):
        """Duplicate an existing memo with its title and content"""
        if self.layout_locked: return
        if key in self.memos:
            content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
            title = self.memos[key]['title_entry'].get()
            self.add_new_memo(initial_text=content, initial_title=title)
            self.save_tab_config()

    def add_new_checklist(self, initial_data=None, initial_title="체크리스트", key=None):
        """Create a new movable/editable checklist box"""
        import time
        if key is None:
            key = f"checklist_{int(time.time() * 1000)}"
        
        # Create container
        check_container = ttk.LabelFrame(self.entry_inner_frame)
        
        # Controls frame (top of checklist)
        ctrl_frame = ttk.Frame(check_container)
        ctrl_frame.pack(fill='x', side='top')
        
        # Editable Title Entry
        title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
        title_entry.pack(side='left', padx=2)
        title_entry.insert(0, initial_title)
        title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_checklist(key))
        btn_copy.pack(side='right')
        
        btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
        btn_del.pack(side='right')
        
        # Add Item Area
        add_frame = ttk.Frame(check_container)
        add_frame.pack(fill='x', padx=2, pady=2)
        
        new_item_var = tk.StringVar()
        entry_new = ttk.Entry(add_frame, textvariable=new_item_var, width=20)
        entry_new.pack(side='left', fill='x', expand=True)
        
        def add_item(event=None):
             text = new_item_var.get().strip()
             if text:
                 self.add_checklist_item(item_frame, text, False, key)
                 new_item_var.set("")
                 self.save_tab_config()
        
        def _adjust_parent_height(self, parent, force=False):
            """Adjust parent frame height with performance check"""
            try:
                # Only update idletasks if forced or we're not in the middle of a high-speed interaction
                # This is the single biggest cause of UI stutter.
                if force:
                    parent.update_idletasks()
            except:
                pass
        
        entry_new.bind('<Return>', add_item)
        btn_add = ttk.Button(add_frame, text="➕", width=3, command=add_item)
        btn_add.pack(side='right')

        # Scrollable Frame for Items
        canvas_frame = ttk.Frame(check_container)
        canvas_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        canvas = tk.Canvas(canvas_frame, height=100, width=200, bg=self.theme_bg, highlightthickness=0) # Match theme
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        item_frame = ttk.Frame(canvas)
        
        item_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=item_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Scroll binding is handled globally in MaterialManager.__init__

        # Make draggable
        self.make_draggable(check_container, key)
        self.draggable_items[key] = check_container
        self.checklists[key] = {
            'container': check_container,
            'title_entry': title_entry,
            'item_frame': item_frame,
            'items': [] # List of item widgets/vars will be managed dynamically via children of item_frame
        }
        
        # Add initial items if provided
        if initial_data:
            for item in initial_data:
                self.add_checklist_item(item_frame, item.get('text', ''), item.get('checked', False), key)
        
        # Initial placement
        if key not in getattr(self, '_loading_memos', []): # Reuse loading flag or logic
            check_container.place(x=1450, y=50)
            
        return check_container

    def save_single_vehicle_data(self, widget):
        """Save ONLY the requested vehicle's data as a standalone entry in the daily log"""
        try:
            # 1. Collect minimal context
            usage_date = str(self.ent_daily_date.get_date())
            site = self.cb_daily_site.get().strip()
            if not site:
                messagebox.showwarning("입력 필요", "차량 점검을 기록할 '현장'을 선택해주세요.")
                return

            # 2. Collect vehicle data
            v_data = widget.get_data()
            v_no = v_data.get('vehicle_info', '').strip()
            if not v_no:
                messagebox.showwarning("입력 필요", "차량번호를 입력해주세요.")
                return

            # Confirm with user
            if not messagebox.askyesno("확인", f"[{v_no}] 차량 점검 정보만 단독으로 저장하시겠습니까?"):
                return

            # 3. Create record (Everything else is zero/empty)
            reserved = ['vehicle_info', 'mileage', 'remarks', '_raw_mileage']
            checks = "|".join([f"{k}:{v}" for k, v in v_data.items() if k not in reserved and v])

            from site_apps.central.src.aux_equipment import value as auxiliary_equipment_value
            record = {
                'Date': usage_date,
                '업체명': self.cb_daily_company.get().strip() or "현장기록",
                'Site': site,
                'User': "차량점검", # Visual tag in worker column
                'WorkTime': "",
                '장비명': self.cb_daily_equip.get().strip(),
                '보조장비': auxiliary_equipment_value(self),
                '검사방법': self.cb_daily_test_method.get().strip(),
                '회사코드': "",
                'Usage': 0.0, '단가': 0.0, '출장비': 0.0, '검사비': 0.0,
                '검사량': 0.0,
                'OT': 0.0, 'OT금액': 0.0,
                'MaterialID': "차량점검",
                '차량번호': v_no,
                '주행거리': v_data.get('_raw_mileage', '0'),
                '차량점검': checks,
                '차량비고': v_data.get('remarks', ''),
                'EntryTime': datetime.datetime.now(),
                '(Full작업자)': "차량점검"
            }
            # Initialize RTK/NDT columns as 0
            for rtk_cat in ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타", "RTK총계"]:
                record[f'RTK_{rtk_cat}'] = 0
            for ndt_cat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]:
                record[f'NDT_{ndt_cat}'] = 0

            # 4. Append and Save
            new_df = pd.DataFrame([record])
            previous_df = self.daily_usage_df
            self.daily_usage_df = pd.concat([previous_df, new_df], ignore_index=True)
            try:
                if self.save_data() is False:
                    self.daily_usage_df = previous_df
                    return
            except Exception:
                self.daily_usage_df = previous_df
                raise

            # Embedded tabs keep their saved values visible; legacy floating editors close.
            saved_key = next((key for key, value in self.vehicle_inspections.items()
                              if value is widget), None)
            panel = getattr(self, 'vehicle_panel', None)
            if panel is not None and widget in panel.editors:
                from site_apps.central.src.vehicle_records import fingerprint
                panel.sources[widget] = {'fingerprints': [fingerprint(self.daily_usage_df.iloc[-1])],
                                         'slot': 0, 'date': usage_date, 'site': site}
                panel.mark_saved(widget)
            elif saved_key is not None:
                self.remove_box(saved_key)
            elif hasattr(widget, 'reset_fields'):
                widget.reset_fields()
            status = getattr(self, 'daily_save_status', None)
            if status is not None:
                status.config(text=f'✓ {v_no} · 차량점검 저장 완료 ({usage_date})')
            
            # 5. UI Updates
            self.update_daily_usage_view()
            self.update_recent_entries_view()
            self.refresh_inquiry_filters()
            
            if status is None:
                messagebox.showinfo("성공", f"[{v_no}] 차량 정보가 기록에 추가되었습니다.")
            return True
            
        except Exception as e:
            messagebox.showerror("오류", f"차량 정보 저장 중 오류 발생: {e}")

    def add_vehicle_inspection_box(self, initial_data=None, key=None):
        """Add a vehicle tab; keep the legacy API for record/session loaders."""
        panel = getattr(self, 'vehicle_panel', None)
        if panel is not None:
            return panel.add(initial_data=initial_data, key=key)
        import datetime
        if key is None:
            key = f"vehicle_inspection_{int(datetime.datetime.now().timestamp() * 1000)}"
            
        # Create draggable container
        container, widget = self.create_draggable_container(
            self.entry_inner_frame, "차량 점검", VehicleInspectionWidget, key,
            manage_list_key='vehicles',
            theme_bg=getattr(self, 'theme_bg', '#f0f0f0'),
            vehicle_list=getattr(self, 'vehicles', []),
            on_save=self.save_single_vehicle_data
        )
        
        # Track specifically for data gathering
        self.vehicle_inspections[key] = widget
        if widget not in self.vehicle_boxes:
            self.vehicle_boxes.append(widget)
        
        # Load initial data if provided
        if initial_data:
            widget.set_data(initial_data)
            
        # Start in the visible left side, including when the input is scrolled.
        if key not in getattr(self, '_loading_memos', []):
            offset = 20 * min(len(self.vehicle_inspections) - 1, 4)
            canvas = self.entry_canvas
            container.place(x=int(canvas.canvasx(12)) + offset,
                            y=int(canvas.canvasy(12)) + offset,
                            width=600, height=300)
            container.lift()
            
        return widget

    def add_checklist_item(self, parent_frame, text, checked, checklist_key):
        """Add a single item row to the checklist"""
        row_frame = ttk.Frame(parent_frame)
        row_frame.pack(fill='x', pady=1)
        
        var = tk.BooleanVar(value=checked)
        cb = ttk.Checkbutton(row_frame, variable=var, command=lambda: self.save_tab_config())
        cb.pack(side='left')
        
        entry = ttk.Entry(row_frame)
        entry.insert(0, text)
        entry.pack(side='left', fill='x', expand=True)
        entry.bind('<FocusOut>', lambda e: self.save_tab_config())
        
        def delete_this_item():
            row_frame.destroy()
            self.save_tab_config()
            
        btn_del_item = ttk.Label(row_frame, text="❌", font=('Arial', 8), cursor='hand2', foreground='gray')
        btn_del_item.pack(side='right', padx=2)
        btn_del_item.bind('<Button-1>', lambda e: delete_this_item())
        
        # Store refs in widget for retrieval during save
        row_frame._checklist_data = {'var': var, 'entry': entry}

    def duplicate_checklist(self, key):
        """Duplicate an existing checklist"""
        if self.layout_locked: return
        if key in self.checklists:
            # get current data
            original_items = []
            item_frame = self.checklists[key]['item_frame']
            for child in item_frame.winfo_children():
                if hasattr(child, '_checklist_data'):
                    data = child._checklist_data
                    original_items.append({
                        'text': data['entry'].get(),
                        'checked': data['var'].get()
                    })
            
            title = self.checklists[key]['title_entry'].get()
            self.add_new_checklist(initial_data=original_items, initial_title=title)
            self.save_tab_config()


    def on_drag_start(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_drag_start_impl
        return on_drag_start_impl(self, *args, **kwargs)
        
    def on_resize_start(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_resize_start_impl
        return on_resize_start_impl(self, *args, **kwargs)
        
    def on_mouse_motion(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_mouse_motion_impl
        return on_mouse_motion_impl(self, *args, **kwargs)

    def _update_widget_position(self, event, widget):
        """Internal helper to calculate and set widget position/size during interaction"""
        dx = event.x_root - widget._drag_start_root_x
        dy = event.y_root - widget._drag_start_root_y
        parent = widget.master
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        
        if widget._interaction_mode == 'move':
            new_x = widget._drag_start_pos_x + dx
            new_y = widget._drag_start_pos_y + dy
            widget_w = widget.winfo_width()
            widget_h = widget.winfo_height()
            
            # Clamp
            if new_x < 0: new_x = 0
            elif new_x + widget_w > parent_w: new_x = max(0, parent_w - widget_w)
            
            # Loosen Y clamping: allow moving slightly below parent to trigger growth
            if new_y < 0: new_y = 0
            # Instead of strict clamp at parent_h, allow a small overflow to trigger resize logic
            elif new_y + widget_h > parent_h + 100: 
                new_y = parent_h + 100 - widget_h

            
            if widget.winfo_manager() != 'place':
                # Save original grid info and explicitly un-grid
                if not hasattr(widget, '_original_grid_info') and widget.grid_info():
                    widget._original_grid_info = widget.grid_info()
                self._ensure_placeholder(widget)
                widget.grid_forget() # Explicitly un-grid before placing
                widget.place(width=widget_w, height=widget_h)
                
            widget.place(x=new_x, y=new_y)
            widget.lift()
            
            # Auto-update scroll region if we're dragging near bounds
            self._ensure_canvas_scroll_region()
            
        elif widget._interaction_mode == 'resize':
            new_width = max(50, widget._start_width + dx)
            new_height = max(20, widget._start_height + dy)
            current_x = widget.winfo_x()
            current_y = widget.winfo_y()
            
            # Clamp
            if current_x + new_width > parent_w: new_width = max(50, parent_w - current_x)
            if current_y + new_height > parent_h: new_height = max(20, parent_h - current_y)
            
            if widget.winfo_manager() != 'place':
                # Save original grid info and explicitly un-grid
                if not hasattr(widget, '_original_grid_info') and widget.grid_info():
                    widget._original_grid_info = widget.grid_info()
                self._ensure_placeholder(widget)
                widget.grid_forget() # Explicitly un-grid before placing
                widget.place(x=current_x, y=current_y)
                
            widget.place(width=new_width, height=new_height)
            widget.lift()
        
        return "break"
        
    def _apply_push_down_logic(self, dragged_widget):
        """Recursively push widgets down if they overlap with the dragged widget"""
        try:
            # Current bounding box of the dragging widget
            x1 = dragged_widget.winfo_x()
            y1 = dragged_widget.winfo_y()
            w1 = dragged_widget.winfo_width()
            h1 = dragged_widget.winfo_height()
            
            padding = 10
            
            # Parent of the widgets
            parent = dragged_widget.master
            
            # OPTIMIZATION: Early filter candidates by parent to avoid massive iterations
            candidates = [w for w in self.draggable_items.values() if w.master == parent]
            
            for other in candidates:
                if other == dragged_widget:
                    continue
                
                # [STABILITY FIX] Skip if the 'other' widget's BOTTOM is ABOVE the dragged widget's TOP.
                # Strictly ignore anything physically above the current interaction zone.
                if (other.winfo_y() + other.winfo_height()) <= y1:
                    continue

                # Skip if not in the same parent
                if other.master != parent:
                    continue
                
                # Skip if hidden
                if not other.winfo_ismapped():
                    continue
                
                # Get current pos/size of the 'other' widget
                x2 = other.winfo_x()
                y2 = other.winfo_y()
                w2 = other.winfo_width()
                h2 = other.winfo_height()
                
                # 1. Horizontal Overlap Check
                if (x1 < x2 + w2) and (x1 + w1 > x2):
                    # 2. Vertical Collision Check
                    # If dragged_widget is hitting 'other' from the top
                    # We check if the bottom of dragged_widget is below the top of other
                    if y1 < y2 and (y1 + h1) > y2:
                        # New target Y for 'other'
                        new_y2 = y1 + h1 + padding
                        
                        # Only move if it's actually pushing it DOWN
                        if y2 < new_y2:
                            if other.winfo_manager() == 'grid':
                                # Convert grid to place
                                self._ensure_placeholder(other)
                                other.lift()
                                other.place(x=x2, y=new_y2, width=w2, height=h2)
                            else:
                                # Just update place
                                other.place(y=new_y2)
                            
                            # 3. Recursive Push
                            self._apply_push_down_logic(other)
        except Exception as e:
            # Silent fail to avoid interrupting drag
            print(f"Error in push-down logic: {e}")

    def _adjust_parent_height(self, parent, force=False):
        """Adjust parent frame height with performance check"""
        try:
            # Only update idletasks if forced or we're not in the middle of a high-speed interaction
            # This is the single biggest cause of UI stutter.
            if force:
                parent.update_idletasks()
            # 2. Start with the bounding box of all GRIDDED items
            try:
                # grid_bbox returns (x, y, width, height) of the grid
                bbox = parent.grid_bbox()
                required_h = bbox[1] + bbox[3] if bbox[3] > 0 else 0
                required_w = bbox[0] + bbox[2] if bbox[2] > 0 else 0
            except:
                required_h = 0
                required_w = 0

            # 3. Handle PLACE items as well (dragged/custom items)
            for child in parent.winfo_children():
                try:
                    manager = child.winfo_manager()
                    if manager == 'place':
                        info = child.place_info()
                        # Get relative or absolute y + height
                        y = child.winfo_y()
                        h = child.winfo_height()
                        x = child.winfo_x()
                        w = child.winfo_width()
                        required_h = max(required_h, y + h)
                        required_w = max(required_w, x + w)
                except:
                    continue
            
            # Add some padding
            new_height = required_h + 30
            
            # [STABILITY FIX] Guard against collapse but allow growth
            # Width is now dictated by canvas parent in responsive mode.
            if new_height < 500: new_height = 800
            
            # Only resize if the required dimensions are significantly different
            current_h = parent.winfo_height()
            if abs(current_h - new_height) > 10:
                 parent.config(height=new_height)
                 
                 # [RECURSIVE FIX] If this parent has a master that is part of the entry system,
                 # adjust its height too. 
                 if parent.master and parent.master != self.entry_inner_frame.master: # Don't go above canvas
                     self._adjust_parent_height(parent.master, force=force)
                     
        except Exception as e:
            print(f"Error adjusting parent height: {e}")


    def _ensure_placeholder(self, widget, width=None, height=None):
        """Ensure a placeholder exists in the grid where the widget used to be"""
        # Map widget to a placeholder attribute name dynamically
        # We can use the widget id or a dictionary, but simpler to attach to widget
        if not hasattr(widget, '_placeholder'):
            # Use provided dims or current widget dims
            w = width if width is not None else widget.winfo_width()
            h = height if height is not None else widget.winfo_height()
            
            # Create a frame to hold the space
            widget._placeholder = ttk.Frame(widget.master, width=w, height=h)
            
            # Grid it at the original position
            if hasattr(widget, '_original_grid_info'):
                widget._placeholder.grid(**widget._original_grid_info)
                # Ensure the placeholder doesn't shrink and holds its size
                widget._placeholder.grid_propagate(False)
                
    def _remove_placeholder(self, widget):
        """Remove placeholder for a widget"""
        if hasattr(widget, '_placeholder'):
            widget._placeholder.destroy()
            del widget._placeholder

        

    def setup_daily_usage_tab(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_view import setup_daily_usage_tab_impl
        return setup_daily_usage_tab_impl(self, *args, **kwargs)

    def on_recent_record_click(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_recent_record_click_impl
        return on_recent_record_click_impl(self, *args, **kwargs)

    def delete_recent_entry(self, event=None):
        """Delete only selected daily records, retaining unrelated stock transactions."""
        selection = self.tv_recent.selection()
        if not selection:
            return
        try:
            indices = set()
            for item in selection:
                tags = self.tv_recent.item(item, 'tags')
                if not tags or int(tags[0]) not in self.daily_usage_df.index:
                    raise ValueError('기록이 변경되었습니다. 새로고침 후 다시 선택해주세요.')
                indices.add(int(tags[0]))
        except (ValueError, TypeError):
            messagebox.showwarning('삭제 오류', '선택한 기록을 확인할 수 없습니다. 새로고침해주세요.', parent=self.root)
            return
        if not messagebox.askyesno('삭제 확인',
                f'선택한 일일 작업 기록 {len(indices)}건을 삭제할까요?\n'
                '검사량·작업자·메모가 함께 삭제되고 재고 및 집계에 반영됩니다.\n'
                '별도 입출고 기록은 유지됩니다.', parent=self.root):
            return
        old_daily = self.daily_usage_df.copy(deep=True)
        old_transactions = self.transactions_df.copy(deep=True)
        cloud_deletes = []
        for idx in indices:
            row = old_daily.loc[idx]
            etime = str(row.get('EntryTime', '')).split('.')[0].strip()
            if etime and etime.lower() not in ('', 'nan', 'nat', 'none'):
                cloud_deletes.append(etime)
        try:
            # Auto-deduction transactions are already excluded by stock calculation.
            # Date/site matches cannot identify which daily job created a legacy transaction.
            self.daily_usage_df = self.daily_usage_df.drop(list(indices))
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.daily_usage_df = old_daily
            self.transactions_df = old_transactions
            messagebox.showerror('삭제 저장 실패', f'{error}\n원본 기록을 복원했습니다.', parent=self.root)
            return

        # [CLOUD SYNC] Propagate deletion to Google Apps Script and update local state immediately
        if cloud_deletes:
            try:
                import threading
                from site_apps.central.src.sync_google_sheet import push_batch_to_web_app, DEFAULT_INVENTORY_PATH, _load_sync_state, _save_sync_state
                try:
                    state = _load_sync_state(DEFAULT_INVENTORY_PATH)
                    cur_times = set(state.get("cloud_entry_times", [])) - set(cloud_deletes)
                    _save_sync_state(DEFAULT_INVENTORY_PATH, cur_times)
                except Exception as ex:
                    print(f"DEBUG: Failed to update local sync state on delete: {ex}")

                def _async_cloud_delete(deletes_list):
                    try:
                        push_batch_to_web_app(deletes=deletes_list)
                    except Exception as err:
                        print(f"DEBUG: Async cloud delete failed: {err}")
                threading.Thread(target=_async_cloud_delete, args=(cloud_deletes,), daemon=True).start()
            except Exception as e:
                print(f"DEBUG: Cloud delete dispatch error: {e}")
        self.update_daily_usage_view()
        self.update_recent_entries_view()
        self.update_stock_view()
        self.update_transaction_view()
        self.refresh_inquiry_filters()
        self.update_monthly_usage_view()
        if hasattr(self, 'budget_view_tree'):
            self.update_budget_site_view()

    def update_recent_entries_view(self):
        """최신 작업일자와 입력시간 순으로 최근 기록 30건을 표시한다."""
        if not hasattr(self, 'tv_recent') or getattr(self, 'daily_usage_df', None) is None:
            return
            
        for item in self.tv_recent.get_children():
            self.tv_recent.delete(item)
            
        try:
            # 선택/수정에서 원본 행 번호를 사용하므로 DataFrame 인덱스는 유지한다.
            recent_df = self.daily_usage_df.copy()
            selected_date = getattr(self, '_daily_recent_date_filter', None)
            if selected_date is not None:
                if not recent_df.empty:
                    dates = pd.to_datetime(recent_df['Date'], errors='coerce').dt.date
                    recent_df = recent_df[dates == selected_date]
                    site = self.cb_daily_site.get().strip()
                    if site and 'Site' in recent_df:
                        recent_df = recent_df[recent_df['Site'].fillna('').astype(str).str.strip() == site]
                if hasattr(self, 'recent_frame'):
                    self.recent_frame.config(text=f'{selected_date:%Y-%m-%d} 저장 내역 · {len(recent_df)}건')
            elif hasattr(self, 'recent_frame'):
                self.recent_frame.config(text=(f'전체 기록 (최신 날짜순 · {len(recent_df)}건)'
                    if getattr(self, '_daily_recent_show_all', False)
                    else '최근 기록 (최신 날짜순 · 최대 30건)'))
            if hasattr(self, 'recent_scope_status'):
                if selected_date is not None:
                    site_label = self.cb_daily_site.get().strip() or '전체 현장'
                    scope_text = f'{selected_date:%Y-%m-%d} · {site_label} · {len(recent_df):,}건'
                elif getattr(self, '_daily_recent_show_all', False):
                    scope_text = f'전체 · {len(recent_df):,}건'
                else:
                    scope_text = f'최근 · {min(30, len(recent_df)):,}건 / 전체 {len(recent_df):,}건'
                self.recent_scope_status.config(text=scope_text)
            if recent_df.empty:
                return

            recent_df['_recent_date'] = pd.to_datetime(recent_df['Date'], errors='coerce')
            if 'EntryTime' in recent_df.columns:
                recent_df['_recent_entry_time'] = pd.to_datetime(
                    recent_df['EntryTime'], errors='coerce'
                )
                recent_df = recent_df.sort_values(
                    by=['_recent_date', '_recent_entry_time'],
                    ascending=[False, False],
                    na_position='last',
                )
            else:
                recent_df = recent_df.sort_values(
                    by='_recent_date', ascending=False, na_position='last'
                )
            if selected_date is None and not getattr(self, '_daily_recent_show_all', False):
                recent_df = recent_df.head(30)
                
            for idx, row in recent_df.iterrows():
                # Extract first worker
                first_worker = str(row.get('User', '')).strip()
                if first_worker.lower() == 'nan':
                    first_worker = ''
                elif first_worker:
                    import re
                    # Remove shift markers like (주간), (야간) if present
                    match = re.match(r"\((주간|야간|휴일|주야간)\)\s*(.*)", first_worker)
                    if match:
                        first_worker = match.group(2).strip()
                        
                date_str = str(row.get('Date', '')).split(' ')[0]
                mat_id = row.get('MaterialID', '')
                mat_name = self.get_material_display_name(mat_id) if hasattr(self, 'get_material_display_name') else mat_id
                
                # [FIX] If Material is empty but Equipment is present, show Equipment in the TreeView
                if not mat_name or str(mat_name).strip() == '':
                    equip_name = str(row.get('장비명', '')).strip()
                    if equip_name and equip_name.lower() != 'nan':
                        mat_name = f"[장비] {equip_name}"
                
                loc_type = str(row.get('구분', '')).strip()
                if not loc_type or loc_type.lower() == 'nan':
                    site = str(row.get('Site', '')).strip()
                    item = str(row.get('검사품명', '')).strip().upper()
                    if '관리소' in site or '관리소' in item or 'STATION' in item or 'V/S' in item or 'B/V' in item:
                        loc_type = '플랜트(관리소)'
                    else:
                        loc_type = '열배관'
                
                values = (
                    idx,
                    date_str,
                    row.get('Site', ''),
                    loc_type,
                    row.get('검사방법', ''),
                    row.get('검사품명', ''),
                    mat_name,
                    row.get('Usage', ''),
                    first_worker,
                    row.get('검사구분', ''),
                    row.get('조인트수', ''),
                    row.get('불량수', ''),
                    row.get('관경(Inch)', '')
                )
                self.tv_recent.insert('', 'end', values=values, tags=(str(idx),))
                
            # Keep the newest row visible after refreshing.
            if self.tv_recent.get_children():
                self.tv_recent.yview_moveto(0)
        except Exception as e:
            print(f"DEBUG: Error updating recent entries view: {e}")

    def setup_daily_usage_query_tab(self):
        """Setup the daily usage query tab"""
        display_frame = ttk.Frame(self.tab_daily_usage_query)
        display_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # KPI Summary Frame
        self.kpi_frame = ttk.LabelFrame(display_frame, text="조회 기간 요약 (KPI)")
        self.kpi_frame.pack(fill='x', padx=5, pady=(0, 5))
        self.lbl_kpi_summary = ttk.Label(self.kpi_frame, text="데이터를 불러오는 중...", font=('Malgun Gothic', 10, 'bold'), foreground='#00529B')
        self.lbl_kpi_summary.pack(side='left', padx=10, pady=5)
        
        # Filter controls
        filter_frame = ttk.Frame(display_frame)
        filter_frame.pack(fill='x', padx=5, pady=0)
        
        # --- Row 1: Date Filters ---
        date_row = ttk.Frame(filter_frame)
        date_row.pack(fill='x', pady=2)
        
        ttk.Label(date_row, text="시작일:").pack(side='left', padx=5)
        self.ent_daily_start_date = DateEntry(date_row, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
        self.ent_daily_start_date.pack(side='left', padx=5)
        # Default to showing all history (starting from 2024)
        start_date = datetime.datetime(2024, 1, 1)
        self.ent_daily_start_date.set_date(start_date)
        
        ttk.Label(date_row, text="종료일:").pack(side='left', padx=5)
        self.ent_daily_end_date = DateEntry(date_row, width=12, date_pattern='yyyy-mm-dd', locale='ko_KR', state='readonly', showweeknumbers=True)
        self.ent_daily_end_date.pack(side='left', padx=5)
        self.ent_daily_end_date.set_date(datetime.datetime.now())
        
        # --- Row 2: Search Filters ---
        filter_row = ttk.Frame(filter_frame)
        filter_row.pack(fill='x', pady=2)
        
        ttk.Label(filter_row, text="업체명:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_company = ttk.Combobox(filter_row, width=12)
        self.cb_daily_filter_company.pack(side='left', padx=2)
        self.cb_daily_filter_company.set('전체')

        def _on_filter_company_changed(event=None):
            comp = self.cb_daily_filter_company.get().strip()
            if not comp or comp == '전체':
                sites = ['전체'] + sorted(self.sites)
            elif hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
                co_col = '업체명' if '업체명' in self.daily_usage_df.columns else ('Company' if 'Company' in self.daily_usage_df.columns else None)
                site_col = '현장' if '현장' in self.daily_usage_df.columns else ('Site' if 'Site' in self.daily_usage_df.columns else None)
                if co_col and site_col:
                    matched = sorted(set(self.daily_usage_df[self.daily_usage_df[co_col] == comp][site_col].dropna().astype(str).str.strip().unique()))
                    matched = [s for s in matched if s and s != 'nan']
                    sites = ['전체'] + (matched or sorted(self.sites))
                else:
                    sites = ['전체'] + sorted(self.sites)
            else:
                sites = ['전체'] + sorted(self.sites)
            self.cb_daily_filter_site['values'] = sites
            site_get = getattr(self.cb_daily_filter_site, 'get', None)
            if site_get and site_get() not in sites:
                self.cb_daily_filter_site.set('전체')
            self.update_daily_usage_view()

        self.cb_daily_filter_company.bind("<<ComboboxSelected>>", _on_filter_company_changed)
        tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                  command=lambda: self.open_list_management_dialog('companies', target_cb=self.cb_daily_filter_company)).pack(side='left', padx=(0, 5))
        
        ttk.Label(filter_row, text="현장:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_site = ttk.Combobox(filter_row, width=12)
        self.cb_daily_filter_site.pack(side='left', padx=2)
        self.cb_daily_filter_site.set('전체')
        tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                  command=lambda: self.open_list_management_dialog('sites', target_cb=self.cb_daily_filter_site)).pack(side='left', padx=(0, 5))
        
        ttk.Label(filter_row, text="자재명:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_material = ttk.Combobox(filter_row, width=15)
        self.cb_daily_filter_material.pack(side='left', padx=2)
        self.cb_daily_filter_material.set('전체')
        tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                  command=lambda: self.open_list_management_dialog('materials', target_cb=self.cb_daily_filter_material)).pack(side='left', padx=(0, 5))
        
        ttk.Label(filter_row, text="장비명:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_equipment = ttk.Combobox(filter_row, width=15)
        self.cb_daily_filter_equipment.pack(side='left', padx=2)
        self.cb_daily_filter_equipment.set('전체')
        tk.Button(filter_row, text="⚙️", font=('Malgun Gothic', 8), bd=0, bg=self.theme_bg, fg='blue', cursor='hand2',
                  command=lambda: self.open_list_management_dialog('equipments', target_cb=self.cb_daily_filter_equipment)).pack(side='left', padx=(0, 5))
        
        ttk.Label(filter_row, text="작업자:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_worker = ttk.Combobox(filter_row, width=10)
        self.cb_daily_filter_worker.pack(side='left', padx=2)
        self.cb_daily_filter_worker.set('전체')
        
        ttk.Label(filter_row, text="분류:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_shift = ttk.Combobox(filter_row, width=8, state="readonly", values=["전체", "주간", "야간", "주야간", "휴일"])
        self.cb_daily_filter_shift.pack(side='left', padx=2)
        self.cb_daily_filter_shift.set('전체')

        ttk.Label(filter_row, text="검사방법:").pack(side='left', padx=(5, 2))
        self.cb_daily_filter_method = ttk.Combobox(filter_row, width=8, state="readonly", values=["전체", "PAUT", "RT", "MT", "PT"])
        self.cb_daily_filter_method.pack(side='left', padx=2)
        self.cb_daily_filter_method.set('전체')
        
        # --- Row 3: Action Buttons (1줄 통합 툴바) ---
        btn_row = ttk.Frame(filter_frame)
        btn_row.pack(fill='x', pady=(4, 2))
        
        # [그룹 1: 조회 제어]
        btn_filter = ttk.Button(btn_row, text="🔍 조회", style='Primary.TButton', command=self.update_daily_usage_view)
        btn_filter.pack(side='left', padx=(5, 3))
        
        btn_filter_reset = ttk.Button(btn_row, text="♻️ 필터 초기화", style='Warning.TButton', command=self.reset_daily_usage_filters)
        btn_filter_reset.pack(side='left', padx=3)

        btn_col_manage = ttk.Button(btn_row, text="⚙️ 컬럼 관리", style='Secondary.TButton', command=self.show_column_visibility_dialog)
        btn_col_manage.pack(side='left', padx=(3, 8))
        
        # [그룹 2: 레코드 편집 및 저장]
        ttk.Button(btn_row, text="📥 입력폼에 불러오기", style='Primary.TButton',
                   command=lambda: self._on_daily_usage_select(None)).pack(side='left', padx=3)
        btn_edit = ttk.Button(btn_row, text="✏️ 선택 수정", style='Info.TButton', command=self.open_edit_daily_usage_dialog)
        btn_edit.pack(side='left', padx=3)

        btn_delete = ttk.Button(btn_row, text="🗑️ 선택 삭제", style='Danger.TButton', command=self.delete_daily_usage_entry)
        btn_delete.pack(side='left', padx=3)

        self.btn_daily_save_list = ttk.Button(btn_row, text="💾 변경사항 저장", command=self.save_all_daily_usage_changes, style='Action.TButton')
        self.btn_daily_save_list.pack(side='left', padx=(3, 8))

        # [그룹 3: 데이터 내보내기 및 백업]
        btn_export = ttk.Button(btn_row, text="📊 엑셀 내보내기", style='Action.TButton', command=self.export_daily_usage_history)
        btn_export.pack(side='left', padx=3)

        btn_export_all = ttk.Button(btn_row, text="📦 전체 백업", style='Secondary.TButton', command=self.export_all_daily_usage)
        btn_export_all.pack(side='left', padx=3)

        # Bindings
        filter_widgets = [
            self.cb_daily_filter_site, self.cb_daily_filter_company, self.cb_daily_filter_material, 
            self.cb_daily_filter_equipment, self.cb_daily_filter_worker, 
            self.cb_daily_filter_shift, self.cb_daily_filter_method
        ]
        for widget in filter_widgets:
            widget.bind("<Return>", lambda e: self.update_daily_usage_view())
            widget.bind("<<ComboboxSelected>>", lambda e: self.update_daily_usage_view())

        for date_widget in [self.ent_daily_start_date, self.ent_daily_end_date]:
            date_widget.bind("<<DateEntrySelected>>", lambda e: self.update_daily_usage_view())
            try:
                date_widget.bind("<Return>", lambda e: self.update_daily_usage_view())
                for child in date_widget.winfo_children():
                    if isinstance(child, (tk.Entry, ttk.Entry)):
                        child.bind("<Return>", lambda e: self.update_daily_usage_view())
            except: pass
        
        # Treeview for daily usage records
        tree_container = ttk.Frame(display_frame)
        tree_container.pack(expand=True, fill='both', padx=5, pady=5)
        
        list_frame = ttk.Frame(tree_container)
        list_frame.pack(fill='both', expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(list_frame, orient="vertical")
        hsb = ttk.Scrollbar(list_frame, orient="horizontal")
        
        # Treeview with RTK categories and NDT materials
        # Note: Workers 1-10 columns are kept in the 'columns' tuple for data storage,
        # but we will only show a consolidated '작업자' in 'displaycolumns'.
        # Added '(Full작업자)' for Excel export backup.
        columns = ('날짜', '업체명', '적용코드', '현장', '구분', '검사품명', '성적서번호', '작업자', '작업시간', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', '장비명', '검사방법', '회사코드', '수량', '단위', '단가', '출장비', '일식', '검사비', 'OT시간', 'OT금액', '품목명', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '입력시간', '차량번호', '주행거리', '차량점검', '차량비고', '(Full작업자)')
        columns += ('메모',)
        self.daily_usage_tree = ttk.Treeview(list_frame, columns=columns, show='headings',
                                              yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Define display columns - 목적에 최적화된 추천 기본 컬럼 (원가 세부항목 대신 단가·검사금액으로 통합)
        visible_defaults = ['날짜', '메모', '구분', '검사방법', '품목명', '수량', '단위', '단가', '검사비', '작업자', '작업시간', 'OT금액', '장비명', '업체명', '비고']
        self.daily_usage_tree['displaycolumns'] = visible_defaults
        
        # [NEW NDT COLUMNS] Add dynamically
        new_cols = ['작업형태', '조건1', '조건2', '보정계수', '제경비', '기술료', '환산물량', '재료비', '인건비', '검사구분', '조인트수', '불량수', '관경(Inch)', '핫태핑차수', '핫태핑개소', '핫태핑관경']
        new_cols += ['보조장비', '원본매수', '재검사매수', '참여인원', '참여시간', '참여OT']
        c_list = list(columns)
        for nc in new_cols:
            if nc not in c_list:
                c_list.append(nc)
        columns = tuple(c_list)
        
        v_list = list(visible_defaults)
        
        self.daily_usage_tree.configure(columns=columns)
        self.daily_usage_tree['displaycolumns'] = v_list
        
        for nc in new_cols:
            self.daily_usage_tree.heading(nc, text=("자재명" if nc == "품목명" else nc), command=lambda c=nc: self.treeview_sort_column(self.daily_usage_tree, c, False))
            self.daily_usage_tree.column(nc, width=100, minwidth=20, stretch=False, anchor='center')

        
        vsb.config(command=self.daily_usage_tree.yview)
        hsb.config(command=self.daily_usage_tree.xview)
        
        # Column configuration
        col_widths = {
            '날짜': 105, '메모': 55, '구분': 85, '검사방법': 75, '수량': 70, '단위': 45,
            '단가': 85, '검사비': 105,
            '작업자': 150, '작업시간': 130, 'OT금액': 90, '장비명': 130, '업체명': 120, '비고': 180,
            '현장': 90, '검사품명': 100, '적용코드': 90, '성적서번호': 110,
            'OT시간': 75, '출장비': 85, '일식': 75,
            '차량번호': 110, '주행거리': 85, '차량점검': 150, '차량비고': 130,
            '품목명': 160, '입력시간': 150,
            'OT1': 90, 'OT2': 90, 'OT3': 90, 'OT4': 90, 'OT5': 90,
            'OT6': 90, 'OT7': 90, 'OT8': 90, 'OT9': 90, 'OT10': 90,
        }
        
        for col in columns:
            disp_title = "자재명" if col == "품목명" else ("검사금액" if col == "검사비" else col)
            self.daily_usage_tree.heading(col, text=disp_title, command=lambda c=col: self.treeview_sort_column(self.daily_usage_tree, c, False))
            width = col_widths.get(col, 100)
            anchor = 'e' if col in ('수량', '단가', '검사비', 'OT금액', '출장비', '일식') else 'center'
            self.daily_usage_tree.column(col, width=width, minwidth=20, stretch=False, anchor=anchor)
        
        # Grid layout for list_frame
        self.daily_usage_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Auto-save column widths when user resizes columns
        # Worker details remain available on a single click; double-click opens notes.
        self.daily_usage_tree.bind("<Button-1>", lambda e: self.show_worker_popup(e, self.daily_usage_tree), add="+")
        self.daily_usage_tree.bind("<Double-1>", self._show_daily_usage_note)
        
        # [NEW] Enable column reordering via drag & drop and context menu
        def _show_daily_usage_tree_context_menu(event):
            region = self.daily_usage_tree.identify_region(event.x, event.y)
            if region == "heading":
                self._show_heading_context_menu(event, self.daily_usage_tree)
                return
            item = self.daily_usage_tree.identify_row(event.y)
            if item:
                self.daily_usage_tree.selection_set(item)
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label="📥 입력 화면에 불러오기", command=lambda: self._on_daily_usage_select(None))
                menu.add_command(label="✏️ 선택 항목 수정...", command=self.open_edit_daily_usage_dialog)
                menu.add_command(label="🗑️ 선택 항목 삭제", command=self.delete_daily_usage_entry)
                menu.add_separator()
                menu.add_command(label="🔍 조회 (새로고침)", command=self.update_daily_usage_view)
                menu.add_command(label="📊 조회결과 엑셀 내보내기", command=self.export_daily_usage_history)
                menu.add_separator()
                menu.add_command(label="⚙️ 컬럼 관리(숨기기/보이기)...", command=self.show_column_visibility_dialog)
                try:
                    menu.tk_popup(event.x_root, event.y_root)
                finally:
                    menu.grab_release()

        self.enable_tree_column_drag(self.daily_usage_tree, context_menu_handler=_show_daily_usage_tree_context_menu)
        
        def save_column_widths(event=None):
            self.save_tab_config()
            
        self.daily_usage_tree.bind('<ButtonRelease-1>', save_column_widths)

        # Selection and context menus must not replace unfinished entry data.
        # Loading is explicit; keep double-click available for the note viewer.
        self.daily_usage_tree.bind('<Return>', self._on_daily_usage_select)
        
        # Optimize treeview scroll region to prevent unnecessary scrollbars
        def optimize_treeview_scroll_region():
            try:
                if hasattr(self, 'daily_usage_tree'):
                    self.daily_usage_tree.update_idletasks()
                    
                    # Get treeview content dimensions
                    total_items = len(self.daily_usage_tree.get_children())
                    if total_items > 0:
                        # Calculate approximate content height
                        item_height = 25  # Approximate height per row
                        content_height = total_items * item_height + 50  # Add some padding
                        
                        # Get treeview dimensions
                        tree_width = self.daily_usage_tree.winfo_width()
                        tree_height = self.daily_usage_tree.winfo_height()
                        
                        # Set scroll region to content size if smaller than treeview
                        if content_height < tree_height:
                            self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, content_height))
                            print(f"Treeview scroll region optimized: (0, 0, {tree_width}, {content_height})")
                        else:
                            # Use full treeview dimensions
                            self.daily_usage_tree.configure(yscrollcommand=(0, 0, tree_width, tree_height))
                            print(f"Treeview scroll region set to full size: (0, 0, {tree_width}, {tree_height})")
            except Exception as e:
                print(f"Error optimizing treeview scroll region: {e}")
        
        # Apply optimization after a short delay
        self.root.after(500, optimize_treeview_scroll_region)


    def _show_daily_usage_note(self, event):
        """Display the clicked record's full note without opening the editor."""
        tree = self.daily_usage_tree
        if tree.identify_region(event.x, event.y) not in ('cell', 'tree'):
            return
        item = tree.identify_row(event.y)
        if not item:
            return
        tags = tree.item(item, 'tags')
        if not tags or 'total' in tags:
            return 'break'
        try:
            entry = self.daily_usage_df.loc[int(tags[0])]
        except (ValueError, KeyError, IndexError):
            return 'break'
        from site_apps.central.src.views.memo_details import daily_memo_lookup, record_notes
        note, daily_memo = record_notes(entry, daily_memo_lookup(self.daily_usage_df))
        details = f"[기록 비고]\n{note or '등록된 비고가 없습니다.'}\n\n[날짜별 메모]\n{daily_memo or '등록된 메모가 없습니다.'}"
        popup = tk.Toplevel(self.root)
        popup.title('비고 · 날짜별 메모')
        popup.geometry('650x320')
        popup.transient(self.root)
        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill='both', expand=True)
        text = tk.Text(frame, wrap='word', font=('Malgun Gothic', 11))
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        text.pack(fill='both', expand=True)
        text.insert('1.0', details)
        text.configure(state='disabled')
        ttk.Button(popup, text='닫기', command=popup.destroy).pack(pady=(0, 10))
        popup.bind('<Escape>', lambda e: popup.destroy())
        text.focus_set()
        return 'break'

    def _on_daily_usage_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_daily_usage_select_impl
        return _on_daily_usage_select_impl(self, *args, **kwargs)


    

    def export_budget_estimation(self):
        """Export the estimated budget details (Planned & Actuals) to Excel"""
        site = self.cb_budget_site.get()
        if not site:
            messagebox.showwarning("현장 미선택", "실행예산서를 내보낼 현장을 선택하거나 입력해주세요.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"공사실행예산정산서_{site}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not file_path:
            return
            
        # Collect Summary Data (Plan, Actual, Diff)
        def get_val(attr):
            return getattr(self, attr).get().strip() if hasattr(self, attr) else ""

        summary_data = {
            'site': site,
            'period': get_val('ent_budget_period'),
            'actual_period': get_val('ent_budget_actual_period'),
            'diff_period': get_val('ent_budget_diff_period'),
            
            'revenue': get_val('ent_budget_revenue'),
            'actual_revenue': get_val('ent_budget_actual_revenue'),
            'diff_revenue': get_val('ent_budget_diff_revenue'),
            
            'unitprice': get_val('ent_budget_unit_price'),
            'actual_unitprice': get_val('ent_budget_actual_unit_price'),
            'diff_unitprice': get_val('ent_budget_diff_unit_price'),
            
            'labor': get_val('ent_budget_labor'),
            'actual_labor': get_val('ent_budget_actual_labor'),
            'diff_labor': get_val('ent_budget_diff_labor'),
            
            'material': get_val('ent_budget_material'),
            'actual_material': get_val('ent_budget_actual_material'),
            'diff_material': get_val('ent_budget_diff_material'),
            
            'expense': get_val('ent_budget_expense'),
            'actual_expense': get_val('ent_budget_actual_expense'),
            'diff_expense': get_val('ent_budget_diff_expense'),
            
            'outsource': get_val('ent_budget_outsource'),
            'actual_outsource': get_val('ent_budget_actual_outsource'),
            'diff_outsource': get_val('ent_budget_diff_outsource'),
            
            'profit': get_val('ent_budget_profit'),
            'actual_profit': get_val('ent_budget_actual_profit'),
            'diff_profit': get_val('ent_budget_diff_profit'),
            
            'margin': get_val('ent_budget_margin'),
            'actual_margin': get_val('ent_budget_actual_margin'),
            'diff_margin': get_val('ent_budget_diff_margin')
        }
        
        # Details - Planned
        labor_data = self.labor_detail_widget.get_data() if hasattr(self, 'labor_detail_widget') else {}
        
        raw_mat = self.material_detail_widget.get_data() if hasattr(self, 'material_detail_widget') else []
        material_data = []
        defaults = self.material_detail_widget.default_items if hasattr(self, 'material_detail_widget') else []
        for i, row in enumerate(raw_mat):
            if i < len(defaults):
                name, spec, unit, _ = defaults[i]
                row['name'] = name
                row['spec'] = spec
                row['unit'] = unit
                material_data.append(row)
                
        expense_data = self.expense_detail_widget.get_data() if hasattr(self, 'expense_detail_widget') else {}

        # Details - Actuals
        actual_labor_data = self.actual_labor_detail_widget.get_data() if hasattr(self, 'actual_labor_detail_widget') else {}
        
        raw_act_mat = self.actual_material_detail_widget.get_data() if hasattr(self, 'actual_material_detail_widget') else []
        actual_material_data = []
        act_defaults = getattr(self.actual_material_detail_widget, 'default_items', defaults)
        for i, row in enumerate(raw_act_mat):
            if i < len(act_defaults):
                name, spec, unit, _ = act_defaults[i]
                row['name'] = name
                row['spec'] = spec
                row['unit'] = unit
                actual_material_data.append(row)
                
        actual_expense_data = self.actual_expense_detail_widget.get_data() if hasattr(self, 'actual_expense_detail_widget') else {}

        # Base salaries for reference
        base_salaries = self.get_base_salaries() if hasattr(self, 'get_base_salaries') else getattr(self.labor_detail_widget, 'base_salaries', {})
        
        try:
            from site_apps.central.src.utils.export_helper import export_budget_estimation_to_excel
            if export_budget_estimation_to_excel(
                file_path,
                summary_data,
                labor_data,
                material_data,
                expense_data,
                actual_labor_data=actual_labor_data,
                actual_material_data=actual_material_data,
                actual_expense_data=actual_expense_data,
                base_salaries=base_salaries
            ):
                messagebox.showinfo("내보내기 완료", f"공사실행예산 및 정산서가 성공적으로 저장되었습니다.\n{file_path}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다.\n{e}")

    def export_budget_sales_status(self):
        """Export the currently filtered budget sales status to Excel"""
        if not hasattr(self, 'budget_view_tree') or not self.budget_view_tree.get_children():
            messagebox.showwarning("데이터 없음", "내보낼 데이터가 없습니다. 먼저 조회를 해주세요.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"매출현황_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not file_path:
            return

        # Collect data from treeview
        data = []
        columns = [self.budget_view_tree.heading(col)['text'] for col in self.budget_view_tree['columns']]
        for item in self.budget_view_tree.get_children():
            data.append(self.budget_view_tree.item(item)['values'])
            
        df = pd.DataFrame(data, columns=columns)
        
        # Add summary row at the bottom
        summary_row = {col: '' for col in columns}
        summary_row['날짜'] = '합계'
        if hasattr(self, 'lbl_bv_actual_income'):
            summary_row['검사단가'] = self.lbl_bv_actual_income.cget('text').replace(' 원', '').replace(',', '').strip()
        
        # Compute OT/Total directly from columns since Budget KPIs might show differently
        ot_idx = columns.index('OT합계') if 'OT합계' in columns else -1
        total_idx = columns.index('합계') if '합계' in columns else -1
        
        sum_ot = 0
        sum_tot = 0
        for _d in data:
            if ot_idx >= 0:
                try: sum_ot += float(str(_d[ot_idx]).replace(',', ''))
                except: pass
            if total_idx >= 0:
                try: sum_tot += float(str(_d[total_idx]).replace(',', ''))
                except: pass

        summary_row['OT합계'] = f"{sum_ot:g}"
        summary_row['합계'] = f"{sum_tot:g}"
        
        # Add the summary row to the data list for the DataFrame
        data.append([summary_row.get(col, '') for col in columns])
        
        df = pd.DataFrame(data, columns=columns)
        
        try:
            from openpyxl.utils import get_column_letter
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='SalesStatus', index=False)
                # Auto-adjust columns
                worksheet = writer.sheets['SalesStatus']
                for i, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    col_letter = get_column_letter(i + 1)
                    worksheet.column_dimensions[col_letter].width = max_len
            
            messagebox.showinfo("내보내기 완료", "매출현황이 엑셀로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {e}")

    def setup_ndt_billing_tab(self):
        # [NEW] Popout Control Frame
        ctrl_frame = ttk.Frame(self.tab_ndt_billing)
        ctrl_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(ctrl_frame, text="💡 이 탭을 별도의 창으로 분리하여 듀얼 모니터에서 작업할 수 있습니다.", foreground='#00529B').pack(side='left')
        ttk.Button(ctrl_frame, text="🔍 팝업창으로 열기", command=self.open_detached_ndt_billing_view).pack(side='right')
        
        # Embed the NDT Calculator as a frame inside this tab
        self.ndt_calculator = NDTCalculatorTab(self.tab_ndt_billing, main_app=self)
        self.ndt_calculator.pack(fill='both', expand=True)

    def open_detached_ndt_billing_view(self):
        """기성 정산(NDT) 탭을 별도의 팝업창으로 엽니다."""
        if hasattr(self, 'detached_windows') and 'ndt_billing' in self.detached_windows and self.detached_windows['ndt_billing']['window'].winfo_exists():
            self.detached_windows['ndt_billing']['window'].lift()
            return
            
        if not hasattr(self, 'detached_windows'):
            self.detached_windows = {}
            
        popup = tk.Toplevel(self.root)
        popup.title("💰 기성 정산 (NDT) (팝업)")
        popup.geometry("1600x900")
        
        self.detached_windows['ndt_billing'] = {'window': popup}
        
        popup_calc = NDTCalculatorTab(popup, main_app=self)
        popup_calc.pack(fill='both', expand=True, padx=5, pady=5)

    def setup_budget_tab(self):
        from site_apps.central.src.views.budget_view import setup_budget_tab_impl
        return setup_budget_tab_impl(self)

    def update_budget_view(self):
        """Refresh the budget inquiry view to reflect updated budget data."""
        if hasattr(self, 'cb_budget_view_site'):
            # Only refresh if a site is actually being viewed
            if self.cb_budget_view_site.get():
                self.update_budget_site_view()

    def _refresh_budget_view_tree_columns(self, reload_data=False):
        """Rebuild budget site performance tree column structure."""
        if not hasattr(self, 'budget_view_tree'):
            return

        custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
        # [FIX] built-in 컬럼과 이름/키가 중복되는 사용자 컬럼 정리
        builtin_keys = set(getattr(self, 'budget_view_builtin_cols', []) or [])
        builtin_headings = set((getattr(self, 'budget_view_builtin_head_map', {}) or {}).values())
        sanitized_custom_cols = []
        seen_custom_keys = set()
        seen_custom_names = set()
        for c in custom_cols:
            key = str(c.get('key', '')).strip()
            name = str(c.get('name', key)).strip()
            if not key:
                continue
            # built-in 과 키 충돌/이름 충돌이면 제거
            if key in builtin_keys or name in builtin_headings:
                continue
            # custom 내부 중복도 제거
            if key in seen_custom_keys or name in seen_custom_names:
                continue
            seen_custom_keys.add(key)
            seen_custom_names.add(name)
            sanitized_custom_cols.append(c)
        if len(sanitized_custom_cols) != len(custom_cols):
            self.budget_view_custom_columns = sanitized_custom_cols
            custom_cols = sanitized_custom_cols

        custom_keys = [c.get('key') for c in custom_cols if c.get('key')]
        self.budget_view_cols = tuple(list(self.budget_view_builtin_cols) + custom_keys)

        head_map = dict(getattr(self, 'budget_view_builtin_head_map', {}))
        width_map = dict(getattr(self, 'budget_view_builtin_width_map', {}))
        for col in custom_cols:
            key = col.get('key')
            if not key:
                continue
            head_map[key] = col.get('name', key)
            width_map[key] = int(col.get('width', 120) or 120)

        alias_map = getattr(self, 'budget_view_heading_aliases', {}) or {}
        self.budget_view_tree.configure(columns=self.budget_view_cols)
        for col in self.budget_view_cols:
            heading_text = alias_map.get(col, head_map.get(col, col))
            self.budget_view_tree.heading(col, text=heading_text,
                                          command=lambda c=col: self.treeview_sort_column(self.budget_view_tree, c, False))
            self.budget_view_tree.column(col, width=width_map.get(col, 100), minwidth=40, anchor='center', stretch=False)

        visible_source = getattr(self, 'budget_view_visible_cols', []) or list(self.budget_view_default_cols)
        visible = [c for c in visible_source if c in self.budget_view_cols]
        self.budget_view_tree['displaycolumns'] = visible if visible else self.budget_view_default_cols
        self.budget_view_visible_cols = list(self.budget_view_tree['displaycolumns'])

        if reload_data:
            self.update_budget_site_view()

    def add_budget_view_custom_column(self):
        """Add a custom column to budget site performance view."""
        name = simpledialog.askstring("컬럼 추가", "새 컬럼 이름을 입력하세요:", parent=self.root)
        if not name or not str(name).strip():
            return
        name = str(name).strip()

        # [FIX] built-in/기존 사용자 컬럼과 이름 중복 방지
        builtin_names = set((getattr(self, 'budget_view_builtin_head_map', {}) or {}).values())
        existing_custom_names = {str(c.get('name', '')).strip() for c in (getattr(self, 'budget_view_custom_columns', []) or [])}
        if name in builtin_names or name in existing_custom_names:
            messagebox.showwarning("중복 컬럼", f"'{name}' 컬럼은 이미 존재합니다.")
            return

        default_value = simpledialog.askstring("컬럼 내용", f"'{name}' 컬럼의 기본 내용을 입력하세요(선택):", parent=self.root)

        existing = set(getattr(self, 'budget_view_cols', []))
        key_base = re.sub(r'[^0-9A-Za-z가-힣_]+', '_', name).strip('_') or '사용자컬럼'
        key = key_base
        idx = 1
        while key in existing:
            idx += 1
            key = f"{key_base}_{idx}"

        custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
        custom_cols.append({'key': key, 'name': name, 'default': default_value or '', 'width': 120})
        self.budget_view_custom_columns = custom_cols

        visible = list(getattr(self, 'budget_view_visible_cols', []) or list(self.budget_view_default_cols))
        if key not in visible:
            visible.append(key)
        self.budget_view_visible_cols = visible

        self._refresh_budget_view_tree_columns(reload_data=True)
        self.save_tab_config()

    def rename_budget_view_column(self, col_key):
        """Rename a budget site performance column heading."""
        if not col_key:
            return
        current_name = self.budget_view_tree.heading(col_key)['text'] if hasattr(self, 'budget_view_tree') else col_key
        new_name = simpledialog.askstring("컬럼 이름 변경", "새 컬럼 이름을 입력하세요:", initialvalue=current_name, parent=self.root)
        if not new_name or not str(new_name).strip():
            return
        new_name = str(new_name).strip()

        # Custom column: rename the stored name itself
        updated = False
        for col in getattr(self, 'budget_view_custom_columns', []) or []:
            if col.get('key') == col_key:
                col['name'] = new_name
                updated = True
                break
        if not updated:
            aliases = getattr(self, 'budget_view_heading_aliases', {}) or {}
            default_name = getattr(self, 'budget_view_builtin_head_map', {}).get(col_key, col_key)
            if new_name == default_name:
                aliases.pop(col_key, None)
            else:
                aliases[col_key] = new_name
            self.budget_view_heading_aliases = aliases

        self._refresh_budget_view_tree_columns(reload_data=False)
        self.save_tab_config()

    def set_budget_view_custom_column_content(self, col_key):
        """Change default content for a custom column."""
        custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
        target = next((c for c in custom_cols if c.get('key') == col_key), None)
        if not target:
            return
        new_value = simpledialog.askstring("컬럼 내용 설정", f"'{target.get('name', col_key)}' 컬럼의 기본 내용을 입력하세요:",
                                           initialvalue=target.get('default', ''), parent=self.root)
        if new_value is None:
            return
        target['default'] = new_value
        self._refresh_budget_view_tree_columns(reload_data=True)
        self.save_tab_config()

    def delete_budget_view_custom_column(self, col_key):
        """Delete a custom budget column."""
        custom_cols = getattr(self, 'budget_view_custom_columns', []) or []
        target = next((c for c in custom_cols if c.get('key') == col_key), None)
        if not target:
            return
        if not messagebox.askyesno("컬럼 삭제", f"'{target.get('name', col_key)}' 컬럼을 삭제하시겠습니까?", parent=self.root):
            return

        self.budget_view_custom_columns = [c for c in custom_cols if c.get('key') != col_key]
        self.budget_view_visible_cols = [c for c in getattr(self, 'budget_view_visible_cols', []) if c != col_key]
        aliases = getattr(self, 'budget_view_heading_aliases', {}) or {}
        aliases.pop(col_key, None)
        self.budget_view_heading_aliases = aliases
        self._refresh_budget_view_tree_columns(reload_data=True)
        self.save_tab_config()

    def show_budget_view_column_dialog(self):
        """Open dialog to show/hide columns in budget site performance view"""
        if not hasattr(self, 'budget_view_tree'):
            return

        all_cols = list(self.budget_view_tree['columns'])
        active_cols = self.budget_view_tree['displaycolumns']
        if not active_cols or active_cols == ('#all'):
            active_cols = all_cols

        dialog_cols = [(col, self.budget_view_tree.heading(col)['text']) for col in all_cols]
        dialog = ColumnSelectionDialog(self.root, dialog_cols, title="현장별 실적 표시 컬럼 관리")
        for col, var in dialog.vars.items():
            var.set(col in active_cols)

        dialog.wait_window()

        if dialog.result is not None:
            # 필수 컬럼 보정: 날짜/현장은 항상 유지
            final_selection = list(dialog.result)
            for mandatory in ('Date', 'Site'):
                if mandatory in all_cols and mandatory not in final_selection:
                    final_selection.append(mandatory)

            sorted_selection = [c for c in all_cols if c in final_selection]
            self.budget_view_tree['displaycolumns'] = sorted_selection if sorted_selection else all_cols
            self.budget_view_visible_cols = list(self.budget_view_tree['displaycolumns'])
            self.save_tab_config()

    def _show_budget_view_heading_context_menu(self, event):
        """공사탭 현장별 실적 Treeview 헤더 우클릭 메뉴"""
        if not hasattr(self, 'budget_view_tree'):
            return
        tree = self.budget_view_tree
        try:
            if tree.identify_region(event.x, event.y) != 'heading':
                return
        except Exception:
            return

        column_id = tree.identify_column(event.x)
        col_key = self._get_column_name_from_id(tree, column_id)
        if not col_key:
            return
        col_text = tree.heading(col_key)['text']
        custom_keys = {c.get('key') for c in (getattr(self, 'budget_view_custom_columns', []) or [])}

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"⬅️ '{col_text}' 왼쪽으로 이동", command=lambda: self._move_column_visual(tree, column_id, -1))
        menu.add_command(label=f"➡️ '{col_text}' 오른쪽으로 이동", command=lambda: self._move_column_visual(tree, column_id, 1))
        menu.add_separator()
        menu.add_command(label=f"✏️ '{col_text}' 이름 변경...", command=lambda: self.rename_budget_view_column(col_key))
        menu.add_command(label="➕ 사용자 컬럼 추가...", command=self.add_budget_view_custom_column)
        if col_key in custom_keys:
            menu.add_command(label=f"📝 '{col_text}' 내용 설정...", command=lambda: self.set_budget_view_custom_column_content(col_key))
            menu.add_command(label=f"🗑️ '{col_text}' 사용자 컬럼 삭제", command=lambda: self.delete_budget_view_custom_column(col_key))
        menu.add_separator()
        menu.add_command(label="⚙️ 컬럼 관리(추가/삭제)...", command=self.show_budget_view_column_dialog)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def update_budget_site_view(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import update_budget_site_view_impl
        return update_budget_site_view_impl(self, *args, **kwargs)


    def _update_budget_kpis(self):
        """Update the top KPI summary labels based on current budget form values"""
        try:
            def _get_val(widget_name):
                try: 
                    w = getattr(self, widget_name, None)
                    if w:
                        v = str(w.get()).replace(',', '').split(' ')[0]
                        return float(v or 0)
                    return 0.0
                except: return 0.0
                
            def _set_val(widget_name, val, is_percent=False):
                w = getattr(self, widget_name, None)
                if w:
                    st = w.cget('state')
                    if st == 'readonly': w.config(state='normal')
                    w.delete(0, 'end')
                    w.insert(0, f"{val:,.1f}" if is_percent else f"{val:,.0f}")
                    if st == 'readonly': w.config(state='readonly')

            def _calc_profit(prefix):
                rev = _get_val(f'{prefix}revenue')
                if rev == 0: rev = _get_val(f'{prefix}unit_price')
                lab = _get_val(f'{prefix}labor')
                mat = _get_val(f'{prefix}material')
                exp = _get_val(f'{prefix}expense')
                out = _get_val(f'{prefix}outsource')
                
                if prefix == 'ent_budget_' and hasattr(self, 'expense_detail_widget') and hasattr(self.expense_detail_widget, 'lbl_grand_total_cost'):
                    try:
                        raw_t = self.expense_detail_widget.lbl_grand_total_cost.cget('text')
                        tc = float("".join(c for c in raw_t if c.isdigit() or c == '.') or 0)
                    except: tc = lab + mat + exp + out
                else:
                    tc = lab + mat + exp + out
                    
                prof = rev - tc
                mar = (prof / rev * 100) if rev > 0 else 0.0
                _set_val(f'{prefix}profit', prof)
                _set_val(f'{prefix}margin', mar, True)
                return rev, lab, mat, exp, out, prof, mar
                
            p_r, p_l, p_m, p_e, p_o, p_p, p_mg = _calc_profit('ent_budget_')
            a_r, a_l, a_m, a_e, a_o, a_p, a_mg = _calc_profit('ent_budget_actual_')
            
            _set_val('ent_budget_diff_revenue', p_r - a_r)
            _set_val('ent_budget_diff_unit_price', _get_val('ent_budget_unit_price') - _get_val('ent_budget_actual_unit_price'))
            _set_val('ent_budget_diff_labor', p_l - a_l)
            _set_val('ent_budget_diff_material', p_m - a_m)
            _set_val('ent_budget_diff_expense', p_e - a_e)
            _set_val('ent_budget_diff_outsource', p_o - a_o)
            _set_val('ent_budget_diff_profit', p_p - a_p)
            _set_val('ent_budget_diff_margin', p_mg - a_mg, True)
            
            if hasattr(self, 'lbl_kpi_rev'):
                self.lbl_kpi_rev.config(text=f"{p_r:,.0f}원")
                self.lbl_kpi_cost.config(text=f"{p_l+p_m+p_e+p_o:,.0f}원")
                self.lbl_kpi_profit.config(text=f"{p_p:,.0f}원", foreground="#ef4444" if p_p < 0 else "#10b981")
                self.lbl_kpi_margin.config(text=f"{p_mg:.1f}%", foreground="#ef4444" if p_mg < 0 else "#10b981")
            self.root.update_idletasks()
        except Exception as e:
            print(f"DEBUG: Budget KPI sync error: {e}")

    def _load_budget_to_form(self, site, silent=False):
        """선택한 현장의 예산서를 상단 입력 폼에 로드한다."""
        previous_site = getattr(self, '_budget_loaded_site', '')
        if not site:
            if not silent: messagebox.showwarning("현장 미선택", "현장을 먼저 선택해주세요.")
            self.cb_budget_site.set(previous_site)
            return
        if not hasattr(self, 'budget_df') or self.budget_df.empty:
            if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
            self.cb_budget_site.set(previous_site)
            return
        match = self.budget_df[self.budget_df['Site'] == site]
        if match.empty:
            if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
            self.cb_budget_site.set(previous_site)
            return

        current = self._capture_budget_input()
        baseline = getattr(self, '_budget_baseline', None)
        dirty = current != baseline if baseline is not None else any(
            str(value).strip() not in ('', '0', '0.0') for value in current['fields'].values())
        if dirty and not messagebox.askyesno('예산 불러오기',
                '저장하지 않은 작성 내용을 버리고 선택한 현장의 예산을 불러올까요?', parent=self.root):
            self.cb_budget_site.set(previous_site)
            return
        self.clear_budget_form()
        row = match.iloc[0]
        self.cb_budget_site.set(site)
        
        mappings = [
            ('ent_budget_period', 'Period'), ('ent_budget_actual_period', 'Actual_Period'),
            ('ent_budget_revenue', 'Revenue'), ('ent_budget_unit_price','UnitPrice'),
            ('ent_budget_labor', 'LaborCost'), ('ent_budget_material', 'MaterialCost'),
            ('ent_budget_expense', 'Expense'), ('ent_budget_outsource', 'OutsourceCost'),
            ('ent_budget_profit', 'Profit'), ('ent_budget_note', 'Note'),
            ('ent_budget_actual_revenue', 'Actual_Revenue'), ('ent_budget_actual_unit_price','Actual_UnitPrice'),
            ('ent_budget_actual_labor', 'Actual_LaborCost'), ('ent_budget_actual_material', 'Actual_MaterialCost'),
            ('ent_budget_actual_expense', 'Actual_Expense'), ('ent_budget_actual_outsource', 'Actual_OutsourceCost'),
            ('ent_budget_actual_profit', 'Actual_Profit'), ('ent_budget_actual_note', 'Actual_Note')
        ]
        
        for attr, col in mappings:
            w = getattr(self, attr, None)
            if w and col in match.columns:
                st = w.cget('state')
                if st == 'readonly': w.config(state='normal')
                w.delete(0, 'end')
                val = row[col]
                if not pd.isna(val) and str(val).lower() != 'nan':
                    if 'note' not in attr.lower():
                        try:
                            # Try to convert to float for formatting
                            f_val = float(str(val).replace(',', ''))
                            w.insert(0, f"{f_val:,.0f}")
                        except ValueError:
                            w.insert(0, str(val))
                    else:
                        w.insert(0, str(val))
                if st == 'readonly': w.config(state='readonly')

        import json as _json
        for json_col, widget_attr in [('LaborDetail',   'labor_detail_widget'),
                                       ('MaterialDetail','material_detail_widget'),
                                       ('ExpenseDetail', 'expense_detail_widget'),
                                       ('Actual_LaborDetail',   'actual_labor_detail_widget'),
                                       ('Actual_MaterialDetail','actual_material_detail_widget'),
                                       ('Actual_ExpenseDetail', 'actual_expense_detail_widget')]:
            jdata = row.get(json_col, '')
            widget = getattr(self, widget_attr, None)
            if widget:
                # [FIX] 사후원가(실적) 위젯인데 DB에 저장된 내역이 없는 경우, 사전예산 내역의 텍스트 구조를 복사 (인원/수량은 비움)
                if json_col.startswith('Actual_') and (not jdata or str(jdata).strip() in ['', '{}', '[]']):
                    planned_col = json_col.replace('Actual_', '')
                    p_jdata = row.get(planned_col, '')
                    if p_jdata and str(p_jdata).strip() not in ['', '{}', '[]']:
                        try:
                            p_data = _json.loads(p_jdata)
                            # 인원 및 수량 초기화
                            if 'Labor' in json_col and isinstance(p_data, dict):
                                for k, v in p_data.items():
                                    if isinstance(v, dict):
                                        v['personnel'] = ''
                                        v['period'] = ''
                            elif 'Material' in json_col and isinstance(p_data, list):
                                for v in p_data:
                                    if isinstance(v, dict):
                                        v['qty'] = ''
                            elif 'Expense' in json_col and isinstance(p_data, dict):
                                for v_list in p_data.values():
                                    if isinstance(v_list, list):
                                        for v in v_list:
                                            if isinstance(v, dict):
                                                if 'qty' in v: v['qty'] = ''
                                                if 'days' in v: v['days'] = ''
                                                # [NEW] Convert units and prices to daily format for Actual Expense
                                                item_name = str(v.get('cat', v.get('item', ''))).strip()
                                                if item_name == '차량유지비':
                                                    v['unit'] = '일'; v['price'] = 6667
                                                elif item_name == '소모품비':
                                                    v['unit'] = '일'; v['price'] = 3333
                                                elif item_name == '복리후생비':
                                                    v['unit'] = '일'; v['price'] = 1667
                                                elif item_name == 'Se-175':
                                                    v['unit'] = '일'; v['price'] = 35714
                            widget.set_data(p_data)
                            continue
                        except Exception as e: 
                            print(f"DEBUG: Failed to copy planned structure to actual widget: {e}")
                
                # 정상 로드 로직
                if jdata and str(jdata).strip() not in ['', '{}', '[]']:
                    try:
                        loaded_data = _json.loads(jdata)
                        if json_col == 'Actual_ExpenseDetail' and isinstance(loaded_data, dict):
                            # Force update existing loaded data to daily units as well
                            for v_list in loaded_data.values():
                                if isinstance(v_list, list):
                                    for v in v_list:
                                        if isinstance(v, dict):
                                            item_name = str(v.get('cat', v.get('item', ''))).strip()
                                            if item_name == '차량유지비' and str(v.get('unit')) == '개월':
                                                v['unit'] = '일'; v['price'] = 6667; v['qty'] = ''
                                            elif item_name == '소모품비' and str(v.get('unit')) == '개월':
                                                v['unit'] = '일'; v['price'] = 3333; v['qty'] = ''
                                            elif item_name == '복리후생비' and str(v.get('unit')) == '개월':
                                                v['unit'] = '일'; v['price'] = 1667; v['qty'] = ''
                                            elif item_name == 'Se-175' and str(v.get('unit')) == '식':
                                                v['unit'] = '일'; v['price'] = 35714; v['qty'] = ''
                        widget.set_data(loaded_data)
                    except:
                        widget.reset()
                else:
                    widget.reset()

        if not silent: messagebox.showinfo("로드 완료", f"'{site}' 현장의 예산서를 불러왔습니다.")
        # [STABILITY] Update KPI summary after load
        self._update_budget_kpis()

        self._budget_loaded_site = site
        self._budget_baseline = self._capture_budget_input()

    def fill_budget_from_actuals(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import fill_budget_from_actuals_impl
        return fill_budget_from_actuals_impl(self, *args, **kwargs)

    def _capture_budget_input(self):
        fields = {key: widget.get() for key, widget in getattr(self, 'budget_widgets', {}).items()
                  if key != 'cb_budget_site' and hasattr(widget, 'get')
                  and (not hasattr(widget, 'cget') or str(widget.cget('state')) != 'readonly')}
        details = {name: getattr(self, name).get_data() for name in (
            'labor_detail_widget', 'material_detail_widget', 'expense_detail_widget',
            'actual_labor_detail_widget', 'actual_material_detail_widget', 'actual_expense_detail_widget')
            if hasattr(self, name)}
        return {'fields': fields, 'details': details}

    def save_budget_entry(self):
        """Save or update budget entry"""
        site = self.cb_budget_site.get().strip()
        if not site:
            messagebox.showwarning("입력 오류", "현장명을 입력해주세요.")
            return

        try:
            def _get(attr):
                widget = getattr(self, attr, None)
                if widget is None:
                    return 0.0
                text = str(widget.get()).replace(',', '').strip()
                try:
                    value = float(text or 0)
                except (ValueError, TypeError):
                    raise ValueError('금액에는 숫자를 입력해주세요.')
                if not math.isfinite(value):
                    raise ValueError('금액에는 유한한 숫자를 입력해주세요.')
                return value

            rev = _get('ent_budget_revenue')
            unit_price = _get('ent_budget_unit_price')
            lab = _get('ent_budget_labor')
            mat = _get('ent_budget_material')
            exp = _get('ent_budget_expense')
            out = _get('ent_budget_outsource')
            profit = rev - (lab + mat + exp + out)
            note = self.ent_budget_note.get().strip() if hasattr(self, 'ent_budget_note') else ""
            period = getattr(self, 'ent_budget_period').get().strip() if hasattr(self, 'ent_budget_period') else ""

            a_rev = _get('ent_budget_actual_revenue')
            a_unit = _get('ent_budget_actual_unit_price')
            a_lab = _get('ent_budget_actual_labor')
            a_mat = _get('ent_budget_actual_material')
            a_exp = _get('ent_budget_actual_expense')
            a_out = _get('ent_budget_actual_outsource')
            a_prof = a_rev - (a_lab + a_mat + a_exp + a_out)
            a_note = getattr(self, 'ent_budget_actual_note').get().strip() if hasattr(self, 'ent_budget_actual_note') else ""
            a_period = getattr(self, 'ent_budget_actual_period').get().strip() if hasattr(self, 'ent_budget_actual_period') else ""
        except ValueError:
            messagebox.showerror("입력 오류", "금액은 숫자여야 합니다.")
            return

        new_data = {
            'Site': site,
            'Period': period,
            'Revenue': rev,
            'UnitPrice': unit_price,
            'LaborCost': lab,
            'MaterialCost': mat,
            'Expense': exp,
            'OutsourceCost': out,
            'Profit': profit,
            'Note': note,
            'Actual_Period': a_period,
            'Actual_Revenue': a_rev, 'Actual_UnitPrice': a_unit, 'Actual_LaborCost': a_lab,
            'Actual_MaterialCost': a_mat, 'Actual_Expense': a_exp, 'Actual_OutsourceCost': a_out,
            'Actual_Profit': a_prof, 'Actual_Note': a_note,
            'LaborDetail': json.dumps(self.labor_detail_widget.get_data()) if hasattr(self, 'labor_detail_widget') else "{}",
            'MaterialDetail': json.dumps(self.material_detail_widget.get_data()) if hasattr(self, 'material_detail_widget') else "{}",
            'ExpenseDetail': json.dumps(self.expense_detail_widget.get_data()) if hasattr(self, 'expense_detail_widget') else "{}",
            'Actual_LaborDetail': json.dumps(self.actual_labor_detail_widget.get_data()) if hasattr(self, 'actual_labor_detail_widget') else "{}",
            'Actual_MaterialDetail': json.dumps(self.actual_material_detail_widget.get_data()) if hasattr(self, 'actual_material_detail_widget') else "{}",
            'Actual_ExpenseDetail': json.dumps(self.actual_expense_detail_widget.get_data()) if hasattr(self, 'actual_expense_detail_widget') else "{}"
        }

        # [NEW] Robust Sanitization: Ensure NO NaN or empty strings ever reach the float64 columns
        # This fixes TypeError: Invalid value '' for dtype 'float64' permanently
        numeric_keys = {
            'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit',
            'Actual_Revenue', 'Actual_UnitPrice', 'Actual_LaborCost', 'Actual_MaterialCost',
            'Actual_Expense', 'Actual_OutsourceCost', 'Actual_Profit'
        }
        new_data_sanitized = {}
        for k, v in new_data.items():
            if k in numeric_keys:
                try:
                    # Coerce everything to float for numeric columns
                    if pd.isna(v) or str(v).strip().lower() in ('nan', 'none', ''):
                        new_data_sanitized[k] = 0.0
                    else:
                        new_data_sanitized[k] = float(v)
                except:
                    new_data_sanitized[k] = 0.0
            else:
                # String columns (Note, Detail JSONs, Site)
                if pd.isna(v) or str(v).lower() == 'nan':
                    new_data_sanitized[k] = ""
                else:
                    new_data_sanitized[k] = str(v)


        original = self.budget_df.copy(deep=True)
        try:
            if not self.budget_df.empty and site in self.budget_df['Site'].values:
                # [FIX] Use explicit .loc[index, key] to ensure correct row update
                idx = self.budget_df[self.budget_df['Site'] == site].index[0]
                for key, val in new_data_sanitized.items():
                    if isinstance(val, str) and key in self.budget_df.columns and self.budget_df[key].dtype != 'object':
                        self.budget_df[key] = self.budget_df[key].astype(object)
                    self.budget_df.loc[idx, key] = val
            else:
                # Create a properly typed DataFrame for concatenation to avoid dtype clashes
                row_df = pd.DataFrame([new_data_sanitized])
                self.budget_df = pd.concat([self.budget_df, row_df], ignore_index=True)

            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.budget_df = original
            messagebox.showerror('예산 저장 실패', f'{error}\n기존 예산 기록을 복원했습니다.', parent=self.root)
            return
        self._budget_loaded_site = site
        self._budget_baseline = self._capture_budget_input()
        messagebox.showinfo("성공", f"'{site}' 현장의 예산 정보가 저장/수정되었습니다.")

        # [NEW] Refresh UI to show the updated data in comparison rows immediately
        self.update_budget_view()

        if 'budget' in self.detached_windows: self._refresh_detached_budget()
        # [NEW] 저장 후 현장 목록 갱신 (새로 기입한 현장명 반영)
        self.refresh_inquiry_filters()
        # self.clear_budget_form() # [UX] Don't clear immediately so user can verify

    def _toggle_hidden_site(self, site, hide=True):
        """현장명을 드롭다운 목록에서 숨기거나 다시 표시한다."""
        if not hasattr(self, 'hidden_sites'):
            self.hidden_sites = []
        if hide:
            if site not in self.hidden_sites:
                self.hidden_sites.append(site)
            messagebox.showinfo("숨기기 완료", f"'{site}' 현장이 목록에서 숨겨졌습니다.\n다시 표시하려면 현장명 입력칸을 우클릭하세요.")
        else:
            if site in self.hidden_sites:
                self.hidden_sites.remove(site)
            messagebox.showinfo("표시 완료", f"'{site}' 현장이 목록에 다시 표시됩니다.")
        # config 저장 및 목록 갱신
        self.save_tab_config()
        self.refresh_inquiry_filters()

    def _show_all_hidden_sites(self):
        """숨긴 현장명을 모두 다시 표시한다."""
        if not getattr(self, 'hidden_sites', []):
            messagebox.showinfo("알림", "숨겨진 현장이 없습니다.")
            return
        count = len(self.hidden_sites)
        names = '\n'.join(f'  • {s}' for s in self.hidden_sites)
        confirmed = messagebox.askyesno(
            "모두 표시",
            f"현재 숨겨진 현장 {count}개:\n{names}\n\n모두 다시 표시하시겠습니까?"
        )
        if confirmed:
            self.hidden_sites.clear()
            self.save_tab_config()
            self.refresh_inquiry_filters()
            messagebox.showinfo("완료", "모든 현장이 다시 표시됩니다.")

    def delete_budget_entry(self):
        """Delete selected budget entry"""
        site = self.cb_budget_site.get().strip()
        if not site:
            messagebox.showwarning("삭제 오류", "삭제할 현장명을 입력해주세요.")
            return
            
        if messagebox.askyesno("삭제 확인", f"'{site}' 현장의 예산 정보를 삭제하시겠습니까?"):
            original = self.budget_df.copy(deep=True)
            try:
                self.budget_df = self.budget_df[self.budget_df['Site'] != site]
                if not self.save_data(require_primary=True):
                    raise RuntimeError('원본 파일에 저장하지 못했습니다.')
            except Exception as error:
                self.budget_df = original
                messagebox.showerror('예산 삭제 실패', f'{error}\n기존 예산 기록을 복원했습니다.', parent=self.root)
                return
            messagebox.showinfo("성공", "삭제되었습니다.")
            self.update_budget_view()
            if 'budget' in self.detached_windows: self._refresh_detached_budget()
            # [NEW] 삭제 후 현장 목록 갱신
            self.refresh_inquiry_filters()
            self.clear_budget_form()


    def clear_budget_form(self):
        """Reset budget form fields while maintaining site selection context"""
        # self.cb_budget_site.set('')  <- Removed to preserve 'appearance' as per user request
        if hasattr(self, 'budget_widgets'):
            for k, w in self.budget_widgets.items():
                if k == "cb_budget_site": continue
                if hasattr(w, 'cget'):
                    st = w.cget('state')
                    if st == 'readonly': w.config(state='normal')
                    w.delete(0, 'end')
                    if st == 'readonly': w.config(state='readonly')
                else:
                    try: w.delete(0, 'end')
                    except: pass
        # Reset both planned and actual detail tables, including values serialized on save.
        for name in ('labor_detail_widget', 'material_detail_widget', 'expense_detail_widget',
                     'actual_labor_detail_widget', 'actual_material_detail_widget', 'actual_expense_detail_widget'):
            widget = getattr(self, name, None)
            if widget is not None:
                widget.reset()

        # [STABILITY] Reset KPI summary
        self._update_budget_kpis()

    def on_budget_tree_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_budget_tree_select_impl
        return on_budget_tree_select_impl(self, *args, **kwargs)

    def _ensure_sash_visible(self):
        """No-op as details panel is removed"""
        pass
    
    def _ensure_daily_usage_sash_visibility(self):
        """Ensure sash position is visible and properly sized using ratio"""
        try:
            if not hasattr(self, 'daily_usage_paned'):
                return
                
            # [STABILITY] Skip if we are still in the initial loading/restoration phase
            # load_tab_config will handle the initial placement.
            if not getattr(self, 'is_ready', False):
                return

            # Force update to get correct dimensions
            self.daily_usage_paned.update_idletasks()
            
            # Get current paned window dimensions
            total_h = self.daily_usage_paned.winfo_height()
            
            if total_h <= 0:
                print("PanedWindow height is 0, retrying...")
                self.root.after(100, self._ensure_daily_usage_sash_visibility)
                return
            
            # If locked, prioritize absolute position if valid
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                if hasattr(self, 'tab_config') and 'daily_usage_sash_pos' in self.tab_config:
                    target_pos = int(self.tab_config['daily_usage_sash_pos'])
                    # Ensure within bounds
                    if 50 < target_pos < total_h - 50:
                        getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, target_pos)
                        print(f"Restored locked absolute position: {target_pos}")
                        return

            # Try to restore from saved ratio first
            # [UX FIX] 작업자 박스 하단이 가려지지 않도록 상단 영역 기본 비율 상향
            target_ratio = 0.45  # Default ratio
            if hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
                target_ratio = self.tab_config['daily_usage_sash_ratio']
            
            # Calculate target position
            target_pos = int(total_h * target_ratio)
            
            # Apply relaxed bounds + practical minimum top area (Reduced from 460 to 100 to allow user custom sizing)
            min_pos = max(int(total_h * 0.1), 100)
            max_pos = int(total_h * 0.95)
            
            if target_pos < min_pos:
                target_pos = min_pos
            elif target_pos > max_pos:
                target_pos = max_pos
            
            # Set sash position
            getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, target_pos)
            
            # Verify position was set correctly
            actual_pos = 500
            actual_ratio = actual_pos / total_h if total_h > 0 else target_ratio
            
            # Update config
            if hasattr(self, 'is_ready') and self.is_ready:
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                self.tab_config['daily_usage_sash_ratio'] = actual_ratio
                self.tab_config['daily_usage_sash_pos'] = actual_pos
            
            # Update canvas scroll region
            self.root.after(100, self._ensure_canvas_scroll_region)
            
        except Exception as e:
            print(f"Error ensuring sash visibility: {e}")
            self.root.after(500, self._ensure_daily_usage_sash_visibility)
    
    def _ensure_monthly_sash_visibility(self):
        """월별 집계 3개 테이블 높이를 저장된 설정 또는 사진 비율(28.2%, 59.0%)로 복원"""
        try:
            if not hasattr(self, 'monthly_paned') or not self.monthly_paned.winfo_exists():
                return
            
            # 탭이 선택되어 화면에 렌더링될 때 정확한 높이 계산
            self.monthly_paned.update_idletasks()
            total_h = self.monthly_paned.winfo_height()
            if total_h <= 50:
                return
            
            # 사진 기반 기본 비율: 상단(월별사용량) 28.2%, 중간(현장별누계) 30.8%, 하단(작업자별) 41.0%
            r0 = 0.282
            r1 = 0.590
            
            if hasattr(self, 'tab_config'):
                saved_ratios = self.tab_config.get('monthly_sash_ratios')
                if saved_ratios and len(saved_ratios) >= 2:
                    try:
                        r0 = float(saved_ratios[0])
                        r1 = float(saved_ratios[1])
                    except: pass
            
            target_pos0 = int(total_h * r0)
            target_pos1 = int(total_h * r1)
            
            # 경계 한계값 (각 테이블이 최소 50px 이상 보이도록 제한)
            min_pane_h = 50
            target_pos0 = max(min_pane_h, min(target_pos0, total_h - 2 * min_pane_h))
            target_pos1 = max(target_pos0 + min_pane_h, min(target_pos1, total_h - min_pane_h))
            
            try:
                self.monthly_paned.sashpos(0, target_pos0)
                self.monthly_paned.sashpos(1, target_pos1)
            except Exception as e:
                print(f"Error setting monthly sash pos: {e}")
        except Exception as e:
            print(f"Error in _ensure_monthly_sash_visibility: {e}")

    def _on_monthly_sash_changed(self, event=None):
        """사용자가 월별 집계 테이블 구분선을 조절했을 때 높이 및 비율 저장"""
        if not getattr(self, 'is_ready', False):
            return
        try:
            if not hasattr(self, 'monthly_paned') or not self.monthly_paned.winfo_exists():
                return
            self.monthly_paned.update_idletasks()
            total_h = self.monthly_paned.winfo_height()
            if total_h <= 50:
                return
            s0 = self.monthly_paned.sashpos(0)
            s1 = self.monthly_paned.sashpos(1)
            if s0 <= 0 or s1 <= 0 or s0 >= s1:
                return
            if not hasattr(self, 'tab_config'):
                self.tab_config = {}
            r0 = round(s0 / total_h, 4)
            r1 = round(s1 / total_h, 4)
            self.tab_config['monthly_sash_positions'] = [s0, s1]
            self.tab_config['monthly_sash_ratios'] = [r0, r1]
            self.save_tab_config()
            print(f"Saved monthly sash positions: [{s0}, {s1}], ratios: [{r0}, {r1}]")
        except Exception as e:
            print(f"Error saving monthly sash positions: {e}")
    
    
    
    def update_resolution_display(self):
        """Update the resolution display"""
        try:
            if hasattr(self, 'resolution_label'):
                # Get window dimensions
                window_width = self.root.winfo_width()
                window_height = self.root.winfo_height()
                
                # Get screen dimensions
                screen_width = self.root.winfo_screenwidth()
                screen_height = self.root.winfo_screenheight()
                
                # Get daily usage paned window dimensions
                if hasattr(self, 'daily_usage_paned'):
                    self.daily_usage_paned.update_idletasks()
                    pane_height = self.daily_usage_paned.winfo_height()
                    pane_width = self.daily_usage_paned.winfo_width()
                    
                    # Get current sash position and ratio
                    try:
                        sash_pos = 500
                        ratio = (sash_pos / pane_height * 100) if pane_height > 0 else 0
                        resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height} | 패널: {pane_width}x{pane_height} | 경계: {ratio:.1f}%"
                    except:
                        resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height} | 패널: {pane_width}x{pane_height}"
                else:
                    resolution_text = f"창: {window_width}x{window_height} | 화면: {screen_width}x{screen_height}"
                
                self.resolution_label.config(text=resolution_text)
                
                # Schedule next update
                self.root.after(500, self.update_resolution_display)
        except Exception as e:
            print(f"Error updating resolution display: {e}")
            # Still schedule next update even if there's an error
            if hasattr(self, 'resolution_label'):
                self.root.after(1000, self.update_resolution_display)
    
    def _restore_locked_position(self):
        """Restore sash to locked position"""
        try:
            if not hasattr(self, 'daily_usage_paned'): return
            
            if hasattr(self, 'tab_config'):
                total_h = self.daily_usage_paned.winfo_height()
                if total_h <= 0: return
                
                # Prioritize absolute position when locked
                if 'daily_usage_sash_pos' in self.tab_config:
                    pos = int(self.tab_config['daily_usage_sash_pos'])
                    # Only apply if it doesn't hide the bottom area completely
                    if 50 < pos < total_h - 50:
                        getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, pos)
                        return
                
                # Fallback to ratio
                if 'daily_usage_sash_ratio' in self.tab_config:
                    ratio = self.tab_config['daily_usage_sash_ratio']
                    locked_pos = int(total_h * ratio)
                    getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, locked_pos)
        except Exception as e:
            print(f"Error restoring locked position: {e}")
    
    def _on_daily_usage_resize(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_daily_usage_resize_impl
        return _on_daily_usage_resize_impl(self, *args, **kwargs)
            # Fallback if something went wrong during resize

    def toggle_sash_lock(self):
        """Toggle sash lock state with auto-fit logic"""
        try:
            self.daily_usage_sash_locked = not self.daily_usage_sash_locked
            
            if self.daily_usage_sash_locked:
                self.root.update_idletasks()
                
                total_height = self.daily_usage_paned.winfo_height()
                sash_pos = 500
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                
                # Calculate and save ratio and absolute position
                ratio = sash_pos / total_height if total_height > 0 else 0.5
                self.tab_config['daily_usage_sash_ratio'] = ratio
                self.tab_config['daily_usage_sash_pos'] = sash_pos
                self.tab_config['daily_usage_sash_locked'] = True
                
                # Save configuration immediately
                self.save_tab_config(force=True)
                
                if hasattr(self, 'btn_sash_lock'):
                    self.btn_sash_lock.config(text="🔒 경계 고정됨")
                    self.btn_sash_lock.configure(style="SashLock.TButton")
                    self.style.configure("SashLock.TButton", foreground="red")
                
                self._start_sash_monitor()
                print(f"Daily usage sash position LOCKED at height: {sash_pos}")
            else:
                # Unlock
                if hasattr(self, 'btn_sash_lock'):
                    self.btn_sash_lock.config(text="🔓 경계 자유")
                    self.btn_sash_lock.configure(style="TButton")
                
                self._stop_sash_monitor()
                if hasattr(self, 'tab_config'):
                    self.tab_config['daily_usage_sash_locked'] = False
                    self.save_tab_config(force=True)
                print("Daily usage sash position UNLOCKED")
            
            # [NEW] Refresh scroll region to ensure it stops correctly at the new boundary
            self._ensure_canvas_scroll_region()
        except Exception as e:
            print(f"Error toggling sash lock: {e}")
    
    def _start_sash_monitor(self):
        """Start periodic monitoring of sash position"""
        if hasattr(self, '_sash_monitor_job'):
            self.root.after_cancel(self._sash_monitor_job)
        
        def check_sash():
            if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
                self._restore_locked_position()
                # Schedule next check
                self._sash_monitor_job = self.root.after(200, check_sash)
        
        # Start monitoring
        self._sash_monitor_job = self.root.after(200, check_sash)
        print("Started sash position monitoring")
    
    def _stop_sash_monitor(self):
        """Stop periodic monitoring of sash position"""
        if hasattr(self, '_sash_monitor_job'):
            self.root.after_cancel(self._sash_monitor_job)
            delattr(self, '_sash_monitor_job')
            print("Stopped sash position monitoring")
    
    def _on_main_window_resize(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_main_window_resize_impl
        return _on_main_window_resize_impl(self, *args, **kwargs)
    
    
    def show_error_dialog(self, title, message):
        """Show a custom error dialog with draggable text"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x400")
        dialog.resizable(True, True)
        
        # Make dialog modal
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"600x400+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title label
        title_label = ttk.Label(main_frame, text=title, font=('Malgun Gothic', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Text widget with scrollbar for error message
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill='both', expand=True)
        
        text_widget = tk.Text(text_frame, wrap='word', font=('Malgun Gothic', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Insert error message
        text_widget.insert('1.0', message)
        text_widget.configure(state='normal')  # Allow selection and copying
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        # Copy button
        def copy_text():
            dialog.clipboard_clear()
            selected_text = text_widget.get('sel.first', 'sel.last') if text_widget.tag_ranges('sel') else text_widget.get('1.0', 'end')
            dialog.clipboard_append(selected_text)
        
        copy_btn = ttk.Button(button_frame, text="복사하기", command=copy_text)
        copy_btn.pack(side='left', padx=5)
        
        # Close button
        close_btn = ttk.Button(button_frame, text="닫기", command=dialog.destroy)
        close_btn.pack(side='right', padx=5)
        
        # Focus on text widget
        text_widget.focus_set()
        
        # Bind Escape key to close
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        
        # Wait for dialog to close
        dialog.wait_window()

    def on_daily_usage_tree_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_daily_usage_tree_select_impl
        return on_daily_usage_tree_select_impl(self, *args, **kwargs)

    def delete_selected_site(self):
        """Remove the selected site from the suggestion list"""
        site = self.cb_daily_site.get().strip()
        if not site:
            messagebox.showwarning("선택 오류", "삭제할 현장명을 선택해주세요.")
            return
            
        if site in self.sites:
            if messagebox.askyesno("삭제 확인", f"'{site}' 현장명을 목록에서 삭제하시겠습니까?\n(기존 기록은 삭제되지 않습니다.)"):
                self.sites.remove(site)
                self.cb_daily_site['values'] = self.sites
                self.cb_daily_site.set('')
                self.save_tab_config()
                messagebox.showinfo("완료", "현장명이 삭제되었습니다.")
        else:
            messagebox.showinfo("알림", "목록에 없는 현장명입니다.")

    def calculate_rtk_total(self):
        """Calculate total RTK usage"""
        try:
            total = 0
            rtk_categories = ["센터미스", "농도", "마킹미스", "필름마크", "취급부주의", "고객불만", "기타"]
            for category in rtk_categories:
                value = self.rtk_entries[category].get()
                if value:
                    total += float(value)
            self.rtk_entries["총계"].config(state='normal')
            self.rtk_entries["총계"].delete(0, tk.END)
            # Format as integer if possible, else float
            if total == int(total):
                self.rtk_entries["총계"].insert(0, str(int(total)))
            else:
                self.rtk_entries["총계"].insert(0, f"{total:.1f}")
            self.rtk_entries["총계"].config(state='readonly')
        except ValueError:
            pass  # Ignore invalid input during typing
    
    def format_entry_with_commas(self, *args, **kwargs):
        from site_apps.central.src.utils.helpers import format_entry_with_commas_impl
        return format_entry_with_commas_impl(self, *args, **kwargs)

    def update_daily_test_fee_calc(self, event=None):
        """Auto-calculate Inspection Fee = (Amount * Unit Price) + Travel Expense"""
        try:
            def get_f(entry):
                try:
                    val = entry.get().strip().replace(',', '')
                    return float(val) if val else 0.0
                except: return 0.0

            amount = get_f(self.ent_daily_test_amount)
            price = get_f(self.ent_daily_unit_price)
            travel = get_f(self.ent_daily_travel_cost)
            
            calc_fee = (amount * price) + travel
            
            # Update the fee field with commas
            self.ent_daily_test_fee.delete(0, tk.END)
            self.ent_daily_test_fee.insert(0, f"{calc_fee:,.0f}")
        except:
            pass

    def _add_single_usage_record_logic(self, mat_id, date_val, site, auto_save=True, prepare_only=False):
        """현장 사용량 기록 1건을 처리하는 핵심 로직 (단일/일괄 공용)"""
        def to_f(ent):
            try:
                # Handle both Entry widgets and raw values
                val = ent.get().replace(',', '') if hasattr(ent, 'get') else str(ent).replace(',', '')
                return float(val) if val else 0.0
            except: return 0.0

        # 1. 작업자 목록 및 시간 데이터 구성
        if self.cb_daily_test_method.get().strip().upper() == 'RT':
            from site_apps.central.src.rt_quantities import rt_counts
            rt_counts({'RT_OR': self.ent_daily_test_amount.get(), 'RT_RE': self.ndt_rej_joint_var.get()})
        from site_apps.central.src.worker_selection import worker_is_selected
        workers_list = []
        worker_data_map = {}
        for i in range(1, 11):
            group = getattr(self, f'worker_group{i}', None)
            u_key = 'User' if i == 1 else f'User{i}'
            wt_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
            ot_key = 'OT' if i == 1 else f'OT{i}'
            m_key = 'Meal' if i == 1 else f'Meal{i}'
            
            if group and worker_is_selected(group):
                name = group.get_worker().strip()
                wt = group.get_time().strip()
                ot = group.get_ot().strip()
                meal = group.get_meal().strip()
                worker_data_map[u_key] = name
                worker_data_map[wt_key] = wt
                worker_data_map[ot_key] = ot
                worker_data_map[m_key] = meal
                if name: workers_list.append(name)
            else:
                worker_data_map[u_key] = ""
                worker_data_map[wt_key] = ""
                worker_data_map[ot_key] = ""
                worker_data_map[m_key] = ""
        
        all_workers = ", ".join(workers_list)

        # 2. RTK 데이터 구성
        rtk_data = {}
        for cat, ent in self.rtk_entries.items():
            rtk_data[f'RTK_{cat}'] = to_f(ent)

        # 3. 차량 데이터 수집 (하단 고정 패널 + 추가 차량 탭)
        panel = getattr(self, 'vehicle_panel', None)
        if panel is not None:
            panel.check_pending_edits()
        all_v_widgets = []
        if hasattr(self, 'fixed_vehicle_widget') and self.fixed_vehicle_widget.winfo_exists():
            all_v_widgets.append(self.fixed_vehicle_widget)
        if hasattr(self, 'vehicle_boxes'):
            for b in self.vehicle_boxes:
                if b not in all_v_widgets:
                    all_v_widgets.append(b)
            
        living_boxes = []
        for b in all_v_widgets:
            if panel is not None and b in panel.sources:
                continue  # Existing vehicle records are updated through 수정 저장.
            if hasattr(b, 'winfo_exists') and b.winfo_exists():
                living_boxes.append(b)
                
        # 4. 업체(Company) 데이터 수집
        company_data_list = []
        client_code = self.cb_daily_client.get().strip() if hasattr(self, 'cb_daily_client') else '한국지역난방공사'
        if not client_code:
            client_code = '한국지역난방공사'
        if hasattr(self, 'ndt_company_entries'):
            for company_entries in self.ndt_company_entries:
                is_first = (company_entries == self.ndt_company_entries[0])
                # [FIX] 첫 번째(주요) 레코드의 회사코드는 항상 발주처(client_code)를 유지하여 자재 협력사 코드로 오염 방지
                company_code = client_code if is_first else (company_entries['_company'].get().strip() or client_code)
                ndt_data = {name: to_f(company_entries.get(name)) for name in self.ndt_materials_all}
                
                # [REFINEMENT] Create record if there is NDT data OR it is the first entry (guaranteed row)
                if any(v > 0 for v in ndt_data.values()) or is_first:
                    company_data_list.append({
                        '회사코드': company_code,
                        'ndt_data': ndt_data,
                        'is_primary': is_first
                    })
        
        if not company_data_list: # [FALLBACK] Ensure at least one row exists
             company_data_list.append({'회사코드': client_code or "한국지역난방공사", 'ndt_data': {}, 'is_primary': True})

        # 5. 공통 레코드 데이터 구성 (기본값)
        from site_apps.central.src.aux_equipment import value as auxiliary_equipment_value
        common_data = {
            'Date': date_val,
            'Site': site,
            'MaterialID': mat_id,
            '장비명': self.cb_daily_equip.get().strip(),
            '보조장비': auxiliary_equipment_value(self),
            '검사방법': self.cb_daily_test_method.get().strip(),
            '검사품명': self.ent_daily_inspection_item.get().strip(),
            '적용코드': self.ent_daily_applied_code.get().strip(),
            '성적서번호': self.ent_daily_report_no.get().strip(),
            '검사자': "", # Derived from workers list in output
            'Usage': to_f(self.ent_daily_test_amount),
            '검사량': to_f(self.ent_daily_test_amount),
            '단가': to_f(self.ent_daily_unit_price),
            '출장비': to_f(self.ent_daily_travel_cost),
            '업체명': self.cb_daily_company.get().strip(),
            'Unit': self.cb_daily_unit.get().strip(),
            '작업형태': self.ndt_work_time_var.get() if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else "",
            '조건1': self.ndt_source_var.get() if self.cb_daily_test_method.get().strip() == 'RT' else (self.ndt_pipe_var.get() if self.cb_daily_test_method.get().strip() in ['UT','PT','PAUT','MT'] else ""),
            '조건2': self.ndt_thickness_var.get() if self.cb_daily_test_method.get().strip() in ['RT','UT'] else "",
            '제경비': getattr(self, '_last_ndt_overhead', 0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 0,
            '기술료': getattr(self, '_last_ndt_tech', 0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 0,
            '보정계수': getattr(self, '_last_ndt_factor', 1.0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 1.0,
            '환산물량': getattr(self, '_last_ndt_adj_qty', 0.0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 0.0,
            '재료비': getattr(self, '_last_ndt_mat_cost', 0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 0,
            '인건비': getattr(self, '_last_ndt_lab_cost', 0) if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else 0,
            '핫태핑차수': self.hot_tap_stage_var.get() if self.ndt_loc_type_var.get() == '핫태핑' else '',
            '핫태핑개소': self.hot_tap_location_var.get().strip() if self.ndt_loc_type_var.get() == '핫태핑' else '',
            '핫태핑관경': self.hot_tap_size_var.get() if self.ndt_loc_type_var.get() == '핫태핑' else '',
            '검사구분': "ORI",
            '구분': self.ndt_loc_type_var.get().strip() if self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT'] else "",
            '조인트수': "",
            '불량수': "",
            '관경(Inch)': getattr(self, 'ndt_report_pipe_var', tk.StringVar(value="")).get().strip(),
        }

        total_meal_worker = sum(to_f(worker_data_map.get(f'Meal{i}' if i > 1 else 'Meal', 0)) for i in range(1, 11))
        global_meal = to_f(self.ent_daily_meal_cost)
        
        # [NEW] Append REP defect trace to Note if applicable (Option A)
        base_note = self._get_merged_memo_and_note()
        rep_count = getattr(self, 'ndt_rep_joint_count_var', tk.StringVar(value="")).get().strip()
        insp_type = getattr(self, 'ndt_inspection_type_var', tk.StringVar(value="ORI")).get().strip()
        if insp_type == "ORI" and rep_count and rep_count != "0":
            base_note = (base_note + f" [불량(REP) {rep_count}개 발생]").strip()
        
        # [NEW] Collect DailyMemo
        daily_memo_text = ""
        if hasattr(self, 'main_memo_text'):
            try: daily_memo_text = self.main_memo_text.get('1.0', 'end-1c').strip()
            except: pass

        common_data.update({
            '일식': total_meal_worker if total_meal_worker > 0 else global_meal,
            '검사비': to_f(self.ent_daily_test_fee),
            'FilmCount': 0.0,
            'EntryTime': datetime.datetime.now(),
            '회사코드': "",
            'DailyMemo': daily_memo_text,
            **worker_data_map
        })

        is_ndt = self.cb_daily_test_method.get().strip() in ['RT','UT','PT','PAUT','MT']
        record_types = []
        if is_ndt:
            ori_j = getattr(self, 'ndt_ori_joint_var', tk.StringVar(value="")).get().strip()
            rep_j = getattr(self, 'ndt_rep_joint_var', tk.StringVar(value="")).get().strip()
            try:
                has_ori = float(ori_j.replace(',', '') or 0) > 0
            except (ValueError, TypeError):
                has_ori = bool(ori_j and ori_j != '0')
            try:
                has_rep = float(rep_j.replace(',', '') or 0) > 0
            except (ValueError, TypeError):
                has_rep = bool(rep_j and rep_j != '0')
            if has_ori or not has_rep:
                record_types.append('ORI')
            if has_rep:
                record_types.append('REP')
        else:
            record_types.append('DEFAULT')

        # --- 차량 데이터 병합 (단일 행 저장용) ---
        merged_v_no = []
        merged_v_mileage = []
        merged_v_check = []
        merged_v_remarks = []

        for v_widget in living_boxes:
            v_data = v_widget.get_data()
            v_no = str(v_data.get('vehicle_info', '')).strip()
            v_mileage = str(v_data.get('_raw_mileage', '')).strip()  # Use raw mileage for safety
            v_remarks = str(v_data.get('remarks', '')).strip()
            
            reserved = ['vehicle_info', 'mileage', 'remarks', '_raw_mileage']
            checks_list = []
            for k, v in v_data.items():
                if k not in reserved and v:
                    checks_list.append(f"{k}:{v}")
            v_check_str = "|".join(checks_list)
            
            # 값이 하나라도 있으면 추가
            if v_no or v_mileage or v_check_str or v_remarks:
                merged_v_no.append(v_no)
                merged_v_mileage.append(v_mileage)
                merged_v_check.append(v_check_str)
                merged_v_remarks.append(v_remarks)
        
        final_v_no = " || ".join(merged_v_no)
        final_v_mileage = " || ".join(merged_v_mileage)
        final_v_check = " || ".join(merged_v_check)
        final_v_remarks = " || ".join(merged_v_remarks)

        records_to_save = []
        # 차량 개수로 행이 늘어나지 않게 max_rows 고정
        max_rows = max(len(company_data_list), 1)
        
        for r_type in record_types:
            r_common_data = common_data.copy()
            
            if is_ndt:
                r_common_data['검사구분'] = r_type
                if r_type == 'ORI':
                    r_common_data['조인트수'] = getattr(self, 'ndt_ori_joint_var', tk.StringVar(value="")).get().strip()
                    qty_str = getattr(self, 'ndt_ori_qty_var', tk.StringVar(value="")).get().strip()
                    if qty_str:
                        r_qty = float(qty_str)
                        r_common_data['Usage'] = r_qty
                        r_common_data['검사량'] = r_qty
                    
                    base_note = self._get_merged_memo_and_note()
                    rej_count = getattr(self, 'ndt_rej_joint_var', tk.StringVar(value="")).get().strip()
                    r_common_data['불량수'] = rej_count if rej_count else "0"
                    if common_data['검사방법'] == 'RT':
                        r_common_data['원본매수'] = to_f(self.ndt_ori_qty_var)
                        r_common_data['Usage'] = r_common_data['검사량'] = common_data['검사량']
                    if rej_count and rej_count != "0":
                        unit = '매' if common_data['검사방법'] == 'RT' else '개'
                        base_note = (base_note + f" [당일 불량(REJ) {rej_count}{unit} 발생]").strip()
                    r_common_data['Note'] = base_note
                elif r_type == 'REP':
                    r_common_data['조인트수'] = getattr(self, 'ndt_rep_joint_var', tk.StringVar(value="")).get().strip()
                    qty_str = getattr(self, 'ndt_rep_qty_var', tk.StringVar(value="")).get().strip()
                    if qty_str:
                        r_qty = float(qty_str)
                        r_common_data['Usage'] = r_qty
                        r_common_data['검사량'] = r_qty
                    
                    r_common_data['Note'] = "(재검사 기록) " + self._get_merged_memo_and_note()
                    if common_data['검사방법'] == 'RT':
                        r_common_data['재검사매수'] = to_f(self.ndt_rep_qty_var)
                        r_common_data['Usage'] = r_common_data['검사량'] = 0.0
                        for key in ('재료비', '인건비', '제경비', '기술료', '환산물량'):
                            r_common_data[key] = 0.0
                        r_common_data['Note'] += f" [REP {r_common_data['재검사매수']:g}매 별도 기록 · 총 물량 제외]"
                    r_common_data['출장비'] = 0.0
                    r_common_data['일식'] = 0.0
                    r_common_data['검사비'] = 0.0 
            else:
                r_common_data['검사구분'] = ""
                r_common_data['Note'] = self._get_merged_memo_and_note()

            if common_data.get('핫태핑차수'):
                tag = f"[핫태핑 {common_data['핫태핑개소']} {common_data['핫태핑관경']} {common_data['핫태핑차수']}]"
                r_common_data['Note'] = (str(r_common_data.get('Note', '')) + ' ' + tag).strip()

            for i in range(max_rows):
                row_record = r_common_data.copy()
                
                # --- 업체/수량 데이터 배분 ---
                if i < len(company_data_list):
                    cd = company_data_list[i]
                    row_record['회사코드'] = cd['회사코드']
                    
                    if not cd['is_primary'] or r_type == 'REP':
                        for j in range(1, 11):
                            wt_k = 'WorkTime' if j == 1 else f'WorkTime{j}'
                            ot_k = 'OT' if j == 1 else f'OT{j}'
                            row_record[wt_k] = ""
                            row_record[ot_k] = ""
                        for k in rtk_data: row_record[k] = 0.0
                    else:
                        for k, v in rtk_data.items(): row_record[k] = v

                    if not cd['is_primary']:
                        for k in ['Usage', '검사량', '단가', '출장비', '검사비', '일식', 'FilmCount']:
                            row_record[k] = 0.0

                    for k, v in cd['ndt_data'].items():
                        row_record[f'NDT_{k}'] = v if r_type != 'REP' else 0.0
                else:
                    for k in ['Usage', '검사량', '단가', '출장비', '검사비', '일식', 'FilmCount']: row_record[k] = 0.0
                    for k in rtk_data: row_record[k] = 0.0
                    for j in range(1, 11):
                        row_record['WorkTime' if j==1 else f'WorkTime{j}'] = ""
                        row_record['OT' if j==1 else f'OT{j}'] = ""
                    for name in self.ndt_materials_all:
                        row_record[f'NDT_{name}'] = 0.0

                # --- 차량 데이터 배분 ---
                if i == 0 and r_type != 'REP':
                    row_record['차량번호'] = final_v_no
                    row_record['주행거리'] = final_v_mileage
                    row_record['차량점검'] = final_v_check
                    row_record['차량비고'] = final_v_remarks
                else:
                    row_record['차량번호'] = ""
                    row_record['주행거리'] = ""
                    row_record['차량점검'] = ""
                    row_record['차량비고'] = ""
                
                records_to_save.append(row_record)

        plan = dict(common_data=common_data, company_data_list=company_data_list,
                    all_workers=all_workers, workers_list=workers_list, records_to_save=records_to_save)
        if prepare_only:
            return plan
        return self._commit_usage_plan(plan, auto_save=auto_save)

    def _commit_usage_plan(self, plan, auto_save=True):
        """Commit a prepared inspection without reading the current entry form."""
        common_data = plan['common_data']
        company_data_list = plan['company_data_list']
        all_workers = plan['all_workers']
        workers_list = plan['workers_list']
        records_to_save = plan['records_to_save']
        date_val, site, mat_id = (common_data[k] for k in ('Date', 'Site', 'MaterialID'))
        # [NEW] Auto-save new unit, equipment, and site (with improved duplicate check)
        new_unit = common_data['Unit'].strip()
        if new_unit and not any(u.strip() == new_unit for u in self.daily_units):
            self.daily_units.append(new_unit)
            self.refresh_ui_for_list_change('daily_units')

        new_equip = common_data['장비명'].strip()
        if new_equip and not any(e.strip() == new_equip for e in self.equipments):
            self.equipments.append(new_equip)
            self.refresh_ui_for_list_change('equipments')

        new_site = common_data['Site'].strip()
        if new_site and not any(s.strip() == new_site for s in self.sites):
            self.sites.append(new_site)
            self.refresh_ui_for_list_change('sites')

        # 7. 재고 트랜잭션 처리 (원본 로직 유지 - 출고는 실제 물량만큼 한 번만 발생)
        # [NEW] PAUT 및 장비류는 재고 차감에서 제외 (사용자 요청)
        method_raw = common_data.get('검사방법', '')
        method = str(method_raw).strip().upper()
        is_excluded = (method == 'PAUT')
        
        if not is_excluded:
            # 7-1. NDT 약품 차감 (세척제, 현상제 등)
            for cd in company_data_list:
                 if cd.get('ndt_data') and any(v > 0 for v in cd['ndt_data'].values()):
                      self._auto_reconcile_and_register_ndt(date_val, site, cd['ndt_data'], all_workers, cd['회사코드'])
            
            # 7-2. 메인 자재 차감 (RT 필름 등)
            if mat_id and common_data['검사량'] > 0:
                mat_display = self.get_material_display_name(mat_id)
                # [FIX] 소모품인 경우에만 재고 차감 (장비류 제외)
                if self._is_consumable_material(mat_display, method):
                    self._create_manual_stock_transaction(date_val, mat_id, 'OUT', common_data['검사량'], site, all_workers, f"{site} 현장 사용 (자동 차감)")
                else:
                    print(f"[DEBUG] Skipping stock deduction for equipment/excluded: {mat_display}")

        # 8. 데이터프레임 업데이트
        if records_to_save:
            self.daily_usage_df = pd.concat([self.daily_usage_df, pd.DataFrame(records_to_save)], ignore_index=True)
            
            # 리스트 명단 자동 추가 (Master List 관리)
            any_list_changed = False
            if site and site not in self.sites:
                self.sites.append(site); self.sites.sort(); any_list_changed = True
            for w in workers_list:
                if w and w not in self.users:
                    self.users.append(w); self.users.sort(); any_list_changed = True
            
            # 모든 생성된 레코드에서 차량번호 수집 및 추가
            current_bases = {v.split('(')[0].strip(): v for v in self.vehicles}
            for r in records_to_save:
                v_no_raw = r.get('차량번호', '').strip()
                if v_no_raw and v_no_raw.lower() not in ['nan', 'none']:
                    v_list = [v.strip() for v in v_no_raw.replace('||', ',').split(',') if v.strip()]
                    for v_no in v_list:
                        base_no = v_no.split('(')[0].strip()
                        if base_no not in current_bases:
                            if any(t in base_no for t in ['81두1580', '89보4028', '90너4889']):
                                new_v = f"{base_no} (탑차)"
                            else:
                                new_v = f"{base_no} (스타렉스)"
                            self.vehicles.append(new_v)
                            current_bases[base_no] = new_v
                            any_list_changed = True
            
            if any_list_changed:
                self.vehicles.sort()
                self.refresh_ui_for_list_change('vehicles') # Update UI comboboxes immediately
                if auto_save:
                    self.save_tab_config()

        if auto_save:
            res = self.save_data()
            return len(records_to_save) if res else 0
        return len(records_to_save)

    def add_ndt_company_section(self):
        """회사별 NDT 입력 섹션 추가"""
        if not hasattr(self, 'ndt_company_entries'):
            self.ndt_company_entries = []
        
        company_idx = len(self.ndt_company_entries)
        ndt_materials = self.ndt_materials_all


        
        # Company frame (Standard frame to save space, no title border)
        company_frame = ttk.Frame(self.ndt_company_container)
        company_frame.pack(fill='x', padx=2, pady=2)
        
        # Single Grid frame for everything to save vertical space
        grid_frame = ttk.Frame(company_frame)
        for c in range(8): grid_frame.columnconfigure(c, weight=1, uniform="ndt_rtk")
        grid_frame.pack(fill='x', padx=2, pady=2)
        
        # Company code selector (Index 0)
        ttk.Label(grid_frame, text="회사코드:", font=('Arial', 8, 'bold')).grid(row=0, column=0, padx=1, pady=1, sticky='e')
        cb_co = ttk.Combobox(grid_frame, width=8, values=getattr(self, 'co_code_list', []))
        cb_co.grid(row=0, column=1, padx=1, pady=1, sticky='ew')
        cb_co.set('')  # Default empty
        
        entries = {'_company': cb_co}  # Store company combobox
        
        # NDT entries grid (Index 1 to 7)
        for i, mat in enumerate(ndt_materials):
            idx = i + 1
            r = idx // 4
            c = (idx % 4) * 2
            ttk.Label(grid_frame, text=f"{mat}:", font=('Arial', 8)).grid(row=r, column=c, padx=1, pady=1, sticky='e')
            e = ttk.Entry(grid_frame, width=6)
            e.grid(row=r, column=c+1, padx=1, pady=1, sticky='ew')
            entries[mat] = e
        
        self.ndt_company_entries.append(entries)
    
    def remove_last_ndt_company(self):
        """마지막 회사 섹션 삭제"""
        if hasattr(self, 'ndt_company_entries') and len(self.ndt_company_entries) > 1:
            # Remove last entry dict
            self.ndt_company_entries.pop()
            # Destroy last widget in container
            widgets = self.ndt_company_container.winfo_children()
            if widgets:
                widgets[-1].destroy()

    def sync_worker_times(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import sync_worker_times_impl
        return sync_worker_times_impl(self, *args, **kwargs)
    def _load_ndt_product_map(self):
        """config 에서 NDT 약품 → 실제 DB 품목명 매핑 불러오기"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                return cfg.get('ndt_product_map', {})
        except Exception:
            pass
        return {}

    def _save_ndt_product_map(self, map_data):
        """NDT 약품 -> 실제 DB 품목명 매핑 저장"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
            else:
                cfg = {}
            cfg['ndt_product_map'] = map_data
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            return True
        except Exception:
            return False

    def open_ndt_product_map_dialog(self):
        """NDT 약품 -> 실제 DB 품목명 매핑 설정 다이얼로그"""
        dlg = tk.Toplevel(self.root)
        dlg.title("NDT 약품-품목 매핑 설정")
        dlg.geometry("500x600")
        dlg.transient(self.root)
        dlg.grab_set()

        main_frame = ttk.Frame(dlg, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="현장 입력 약품명", font=('Arial', 10, 'bold')).grid(row=0, column=0, pady=10)
        ttk.Label(main_frame, text="창고 재고 품목 (매핑)", font=('Arial', 10, 'bold')).grid(row=0, column=1, pady=10)

        ndt_materials = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        current_map = self._load_ndt_product_map()
        
        material_options = []
        for _, row in self.materials_df.iterrows():
            disp = self.get_material_display_name(row['MaterialID'])
            material_options.append(disp)
        material_options.sort()

        combos = {}
        for i, mat in enumerate(ndt_materials):
            ttk.Label(main_frame, text=mat).grid(row=i+1, column=0, padx=5, pady=5, sticky='w')
            cb = ttk.Combobox(main_frame, values=material_options, width=40)
            cb.grid(row=i+1, column=1, padx=5, pady=5, sticky='ew')
            current_id = current_map.get(mat, "")
            if current_id:
                cb.set(self.get_material_display_name(current_id))
            combos[mat] = cb

        def _save():
            new_map = {}
            for mat, cb in combos.items():
                disp = cb.get().strip()
                if disp:
                    for _, row in self.materials_df.iterrows():
                        if self.get_material_display_name(row['MaterialID']) == disp:
                            new_map[mat] = row['MaterialID']
                            break
            if self._save_ndt_product_map(new_map):
                messagebox.showinfo("성공", "매핑 설정이 저장되었습니다.")
                dlg.destroy()
            else:
                messagebox.showerror("오류", "설정을 저장하지 못했습니다.")

        def _clear():
            for cb in combos.values():
                cb.set('')

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=len(ndt_materials)+1, column=0, columnspan=2, pady=20)

        ttk.Button(btn_frame, text="저장", command=_save, width=10).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="전체 초기화", command=_clear, width=10).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="닫기", command=dlg.destroy, width=10).pack(side='left', padx=8)

    def open_equipment_search_dialog(self):
        """Select multiple central equipment units without losing serial numbers."""
        from site_apps.central.src.aux_equipment import select_main
        return select_main(self.root, self, self.cb_daily_equip)

    def open_preferred_item_search_dialog(self):
        """Deprecated in favor of unified open_material_search_dialog"""
        self.open_material_search_dialog(target_form='registration')

    def _get_merged_memo_and_note(self):
        """[NEW] Collects user-entered note and all open memo contents for archiving"""
        main_note = ""
        if hasattr(self, 'ent_daily_note'):
            main_note = self.ent_daily_note.get().strip()
            
        memo_archive = []
        if hasattr(self, 'memos'):
            for key, m in self.memos.items():
                try:
                    title = m['title_entry'].get().strip()
                    content = m['text_widget'].get('1.0', 'end-1c').strip()
                    if content:
                        memo_archive.append(content)
                except: continue
        
        if not memo_archive:
            return main_note
            
        archive_text = " | ".join(memo_archive)
        return f"{main_note} | {archive_text}" if main_note else archive_text

    def load_daily_usage_to_form(self, record):
        """[NEW] Load a database record back into the input form for editing or report generation"""
        from site_apps.central.src.inspection_batch import participation_record
        record = participation_record(record)
        try:
            def to_f_local(v):
                try: return float(str(v).replace(',', '')) if v else 0.0
                except: return 0.0

            # 1. Basic Info
            if 'Date' in record:
                try: self.ent_daily_date.set_date(pd.to_datetime(record['Date']))
                except: pass
            self._last_daily_date = self.ent_daily_date.get_date()
            if hasattr(self, 'cb_daily_client'):
                client_val = self.clean_nan(record.get('회사코드', '')).strip()
                valid_clients = ['한국지역난방공사', '한국지역난방공사 중앙지사']
                if hasattr(self.cb_daily_client, 'cget'):
                    try:
                        vals = list(self.cb_daily_client.cget('values'))
                        if vals:
                            valid_clients = vals
                    except: pass
                if client_val in valid_clients:
                    self.cb_daily_client.set(client_val)
                elif '한국지역난방' in client_val:
                    self.cb_daily_client.set(client_val)
                else:
                    self.cb_daily_client.set('한국지역난방공사')
            if '업체명' in record:
                comp_val = self.clean_nan(record['업체명']).split(' [')[0]
                self.cb_daily_company.set(comp_val)
            if 'Site' in record:
                self.cb_daily_site.set(self.clean_nan(record['Site']))
            if '구분' in record:
                self.ndt_loc_type_var.set(self.clean_nan(record['구분']))
            if '적용코드' in record:
                self.ent_daily_applied_code.delete(0, tk.END)
                self.ent_daily_applied_code.insert(0, self.clean_nan(record['적용코드']))
            from site_apps.central.src.aux_equipment import restore
            restore(self, record.get('보조장비', ''))
            if '장비명' in record:
                self.cb_daily_equip.delete(0, tk.END)
                self.cb_daily_equip.insert(0, self.clean_nan(record['장비명']))
            if '성적서번호' in record:
                self.ent_daily_report_no.delete(0, tk.END)
                self.ent_daily_report_no.insert(0, self.clean_nan(record['성적서번호']))
            if '검사품명' in record:
                self.ent_daily_inspection_item.delete(0, tk.END)
                self.ent_daily_inspection_item.insert(0, self.clean_nan(record['검사품명']))
            if '검사방법' in record:
                method = self.clean_nan(record.get('검사방법', ''))
                self.cb_daily_test_method.set(method)
                
                # [FIX] Explicitly show/hide NDT related frames to ensure they appear
                if method in ["MT", "PT"]:
                    if hasattr(self, 'ndt_frame'):
                        try: self.ndt_frame.grid()
                        except: pass
                else:
                    if hasattr(self, 'ndt_frame'):
                        try: self.ndt_frame.grid_remove()
                        except: pass
                
                if method in ["RT", "UT", "PT", "PAUT"]:
                    if hasattr(self, 'ndt_calc_frame'):
                        try:
                            self.ndt_calc_frame.grid()
                            self.ndt_calc_frame.lift()
                        except: pass
                    if method == "RT":
                        if hasattr(self, 'rtk_grid'):
                            try: self.rtk_grid.grid()
                            except: pass
                    else:
                        if hasattr(self, 'rtk_grid'):
                            try: self.rtk_grid.grid_remove()
                            except: pass
                else:
                    if hasattr(self, 'ndt_calc_frame'):
                        try: self.ndt_calc_frame.grid_remove()
                        except: pass
                    if hasattr(self, 'rtk_grid'):
                        try: self.rtk_grid.grid_remove()
                        except: pass
                        
                self.cb_daily_test_method.event_generate('<<ComboboxSelected>>')
            if 'Unit' in record:
                self.cb_daily_unit.set(self.clean_nan(record['Unit']))
            elif '단위' in record:
                self.cb_daily_unit.set(self.clean_nan(record['단위']))
            
            self.hot_tap_stage_var.set(self.clean_nan(record.get('핫태핑차수', '')) or '1차')
            self.hot_tap_size_var.set(self.clean_nan(record.get('핫태핑관경', '')) or '600A')
            self.hot_tap_location_var.set(self.clean_nan(record.get('핫태핑개소', '')))
            # 2. Quantities & Costs
            def set_val(ent, key):
                val = record.get(key, 0)
                ent.delete(0, tk.END)
                if isinstance(val, (int, float)):
                    if val != 0: ent.insert(0, f"{val:,.0f}" if any(x in key for x in ['비', '가', '일식']) else str(val))
                else: 
                    ent.insert(0, str(val))

            set_val(self.ent_daily_test_amount, '검사량')
            set_val(self.ent_daily_unit_price, '단가')
            set_val(self.ent_daily_travel_cost, '출장비')
            set_val(self.ent_daily_test_fee, '검사비')
            set_val(self.ent_daily_meal_cost, '일식')
            if '작업형태' in record and record['작업형태']: self.ndt_work_time_var.set(record['작업형태'])
            if '조건1' in record and record['조건1']: 
                if record.get('검사방법') == 'RT': self.ndt_source_var.set(record['조건1'])
                else: self.ndt_pipe_var.set(record['조건1'])
            if '조건2' in record and record['조건2']: self.ndt_thickness_var.set(record['조건2'])
            def get_valid_rate(key, default_val):
                if key in record:
                    val = record[key]
                    if pd.notna(val) and str(val).strip() and str(val).lower() != 'nan':
                        try: return float(val)
                        except: pass
                return default_val
            
            self.ndt_overhead_var.set(get_valid_rate('제경비율', 110.0))
            self.ndt_tech_var.set(get_valid_rate('기술료율', 20.0))

            # NDT 상세 조건 필드 로딩
            if '관경(Inch)' in record and pd.notna(record['관경(Inch)']):
                self.ndt_report_pipe_var.set(str(record['관경(Inch)']).replace('.0', '') if str(record['관경(Inch)']).endswith('.0') else str(record['관경(Inch)']))
            else:
                self.ndt_report_pipe_var.set("")
                
            insp_type = self.clean_nan(record.get('검사구분', 'ORI')).upper()
            if not insp_type: insp_type = 'ORI'
            if insp_type == 'REP':
                self.ndt_rep_joint_var.set(self.clean_nan(record.get('조인트수', '')))
                self.ndt_rep_qty_var.set(self.clean_nan(record.get('재검사매수', record.get('Usage', ''))))
            else:
                self.ndt_ori_joint_var.set(self.clean_nan(record.get('조인트수', '')))
                self.ndt_ori_qty_var.set(self.clean_nan(record.get('원본매수', record.get('Usage', ''))))
                
            if '불량수' in record and pd.notna(record['불량수']):
                self.ndt_rej_joint_var.set(str(record['불량수']))
            else:
                self.ndt_rej_joint_var.set("")
            
            # 3. Material
            if 'MaterialID' in record:
                mat_id_val = record['MaterialID']
                disp_name = self.get_material_display_name(mat_id_val)
                # [FIX] If display name is empty but MaterialID is a non-NaN string (e.g. 'JIREH Scanner'),
                # show the raw MaterialID string as a fallback
                if not disp_name:
                    try:
                        if not pd.isna(mat_id_val) and str(mat_id_val).strip().lower() not in ('nan', ''):
                            disp_name = str(mat_id_val).strip()
                    except: pass
                print(f"[DEBUG] load_to_form: MaterialID={repr(mat_id_val)}, disp={repr(disp_name)}")
                if isinstance(self.cb_daily_material, ttk.Combobox):
                    self.cb_daily_material.set(disp_name)
                else:
                    self.cb_daily_material.delete(0, tk.END)
                    self.cb_daily_material.insert(0, disp_name)
                from site_apps.central.src.material_selection import remember_selection
                remember_selection(self, mat_id_val)

            # 4. Workers
            # First, clear existing workers if needed or just overwrite
            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if not group: continue
                u_key = 'User' if i == 1 else f'User{i}'
                wt_key = 'WorkTime' if i == 1 else f'WorkTime{i}'
                ot_key = 'OT' if i == 1 else f'OT{i}'
                m_key = 'Meal' if i == 1 else f'Meal{i}'
                
                group.cb_name.set(self.clean_nan(record.get(u_key, '')))
                # [FIX] Use set_time to properly parse (Shift) Time string and populate both shift and time widgets
                group.set_time(self.clean_nan(record.get(wt_key, '')))
                
                group.ent_ot.delete(0, tk.END)
                group.ent_ot.insert(0, self.clean_nan(record.get(ot_key, '')))
                
                group.ent_meal.delete(0, tk.END)
                group.ent_meal.insert(0, self.clean_nan(record.get(m_key, '')))

            # 5. RTK

            for cat, ent in self.rtk_entries.items():
                if cat == "총계": continue
                val = record.get(f'RTK_{cat}', 0)
                ent.delete(0, tk.END)
                if abs(to_f_local(val)) > 0.001: ent.insert(0, str(val))
            self.calculate_rtk_total()


            # 6. Vehicle
            v_no_raw = self.clean_nan(record.get('차량번호', ''))
            v_insp_raw = self.clean_nan(record.get('차량점검', ''))
            v_mileage_raw = self.clean_nan(record.get('주행거리', ''))
            v_remarks_raw = self.clean_nan(record.get('차량비고', ''))

            v_no_list = [x.strip() for x in str(v_no_raw).split('||')] if v_no_raw else []
            v_insp_list = [x.strip() for x in str(v_insp_raw).split('||')] if v_insp_raw else []
            v_mileage_list = [x.strip() for x in str(v_mileage_raw).split('||')] if v_mileage_raw else []
            v_remarks_list = [x.strip() for x in str(v_remarks_raw).split('||')] if v_remarks_raw else []
            
            max_v_count = max(len(v_no_list), len(v_insp_list), len(v_mileage_list), len(v_remarks_list))
            
            # 기존 플로팅 창 닫기 및 고정 패널 비우기
            if hasattr(self, 'vehicle_panel'):
                self.vehicle_panel.reset()
            if hasattr(self, 'draggable_items'):
                for key, cont in list(self.draggable_items.items()):
                    w = getattr(cont, '_widget', None)
                    if not w: continue
                    if type(w).__name__ == 'VehicleInspectionWidget':
                        self.remove_box(key)
            if hasattr(self, 'fixed_vehicle_widget'):
                self.fixed_vehicle_widget.reset_fields()

            for i in range(max_v_count):
                cur_no = v_no_list[i] if i < len(v_no_list) else ""
                cur_insp = v_insp_list[i] if i < len(v_insp_list) else ""
                cur_mileage = v_mileage_list[i] if i < len(v_mileage_list) else ""
                cur_remarks = v_remarks_list[i] if i < len(v_remarks_list) else ""
                
                if not (cur_no or cur_insp or cur_mileage or cur_remarks):
                    continue

                v_parsed = {'vehicle_info': cur_no, 'mileage': cur_mileage, 'remarks': cur_remarks}
                if ':' in cur_insp:
                    for pair in cur_insp.split('|'):
                        if ':' in pair:
                            k, v_val = pair.split(':', 1)
                            v_parsed[k.strip()] = v_val.strip()
                elif cur_insp and cur_insp != 'nan':
                    for k in cur_insp.split(','):
                        k_clean = k.strip()
                        if k_clean:
                            if 'locking' in k_clean: v_parsed[k_clean] = '잠금'
                            elif 'cleaning' in k_clean: v_parsed[k_clean] = '함'
                            else: v_parsed[k_clean] = '양호'
                
                if i == 0 and hasattr(self, 'fixed_vehicle_widget'):
                    v_widget = self.fixed_vehicle_widget
                    v_widget.set_data(v_parsed)
                else:
                    if hasattr(self, 'add_vehicle_inspection_box'):
                        v_widget = self.add_vehicle_inspection_box()
                        if hasattr(v_widget, 'set_data'):
                            v_widget.set_data(v_parsed)
            
            # 7. NDT Chemicals
            if hasattr(self, 'ndt_company_entries') and self.ndt_company_entries:
                # [STABILITY] Clear first if multiple companies, but for now focus on the primary/first entry
                first = self.ndt_company_entries[0]
                ndt_names = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
                for name in ndt_names:
                    col_name = f"NDT_{name}"
                    if col_name in record:
                        val = record[col_name]
                        if name in first:
                            first[name].delete(0, tk.END)
                            if abs(to_f_local(val)) > 0.001:
                                first[name].insert(0, str(val))
                
                # Also load company if available
                if '회사코드' in record:
                    raw_co = self.clean_nan(record.get('회사코드', '')).strip()
                    if raw_co and raw_co not in ('한국지역난방공사', '한국지역난방공사 중앙지사'):
                        first['_company'].set(raw_co)
                    else:
                        first['_company'].set('')

            # 8. Note
            note_val = self.clean_nan(record.get('Note', ''))
            if 'Note' in record:
                self.ent_daily_note.delete(0, tk.END)
                # [REMOVED] 사용자 요청으로 과거 비고(Note) 내용을 불러오지 않음
                # self.ent_daily_note.insert(0, note_val)
                
            # 9. DailyMemo
            # 사용자가 저장한 메모를 불러오도록 복구 (과거 기록 호환성 포함)
            memo_val = self.clean_nan(record.get('DailyMemo', ''))
            if not memo_val and note_val:
                if ' | ' in str(note_val):
                    parts = str(note_val).split(' | ', 1)
                    if len(parts) > 1:
                        memo_val = parts[1]
                else:
                    memo_val = note_val
                    
            if hasattr(self, 'main_memo_text') and self.main_memo_text:
                try:
                    self.main_memo_text.delete('1.0', tk.END)
                    if memo_val:
                        self.main_memo_text.insert(tk.END, memo_val)
                except: pass

        except Exception as e:
            print(f"Error loading record to form: {e}")

    def clear_daily_usage_form_all(self, keep_date=False, keep_workers=False):
        """현장별 일일 사용량 입력 폼의 모든 필드를 초기화합니다."""
        # 1. 콤보박스 선택 해제
        self._daily_edit_source = None
        self._daily_material_selection = None
        for cb in [self.cb_daily_company, self.cb_daily_site, 
                   self.cb_daily_test_method, self.cb_daily_unit]:
            cb.set('')
        
        # 2. 날짜 초기화 (연속 기입 시 편의를 위해 keep_date 옵션 제공)
        if not keep_date:
            try:
                self.ent_daily_date.set_date(datetime.date.today())
            except: pass
        
        self._last_daily_date = self.ent_daily_date.get_date()
        from site_apps.central.src.aux_equipment import restore
        restore(self)
        # 3. 모든 Entry 필드 비우기
        for ent in [self.ent_daily_inspection_item, self.ent_daily_applied_code, 
                   self.ent_daily_equip, self.cb_daily_material, # Both are now Entries
                   self.ent_daily_test_amount, self.ent_daily_unit_price, 
                   self.ent_daily_report_no, self.ent_daily_travel_cost, 
                   self.ent_daily_meal_cost, self.ent_daily_test_fee, self.ent_daily_note]:
            ent.delete(0, tk.END)
            
        # 기본값 복구
        self.cb_daily_unit.set('매')
        self.ndt_work_time_var.set("일반")
        self.ndt_source_var.set("Se-75 (1.0)")
        self.ndt_thickness_var.set("조건없음 (1.0)")
        self.ndt_pipe_var.set("250mm 초과 [10인치 이상] (1.0)")
        self.ndt_overhead_var.set(110.0)
        self.ndt_tech_var.set(20.0)
        self.ndt_ori_joint_var.set("")
        self.ndt_ori_qty_var.set("")
        self.ndt_rep_joint_var.set("")
        self.ndt_rep_qty_var.set("")
        self.ndt_rej_joint_var.set("")
        self.ndt_report_pipe_var.set("")
        self.ndt_loc_type_var.set('열수송관')
        self.hot_tap_stage_var.set('1차')
        self.hot_tap_size_var.set('600A')
        self.hot_tap_location_var.set('')
        self.ndt_calc_frame.grid_remove()
        self.rtk_grid.grid_remove() # [NEW] Hide RTK on clear
        self.ndt_frame.grid_remove() # [NEW] Hide NDT frame on clear
        
        # 4. NDT 자재 및 RTK 필드 초기화
        if hasattr(self, 'ndt_company_entries'):
            for entries in self.ndt_company_entries:
                for mat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]:
                    if mat in entries: entries[mat].delete(0, tk.END)
        
        for ent in self.rtk_entries.values():
            if ent.winfo_exists():
                ent.config(state='normal')
                ent.delete(0, tk.END)
                if ent == self.rtk_entries.get("총계"): 
                    ent.config(state='readonly')
        
        # 5. 차량 점검 섹션 초기화 (중복 방지 및 고정 패널 사용)
        for key in list(getattr(self, 'vehicle_inspections', {})):
            self.remove_box(key)
        if hasattr(self, 'draggable_items'):
            for key, cont in list(self.draggable_items.items()):
                w = getattr(cont, '_widget', None)
                if not w: continue
                lbl_text = ""
                if hasattr(cont, '_label_widget'):
                    lbl_text = cont._label_widget.cget('text')
                is_vehicle = (type(w).__name__ == 'VehicleInspectionWidget') or ('차량' in lbl_text and '점검' in lbl_text) or ('상시' in lbl_text and '차량' in lbl_text)
                if is_vehicle:
                    print(f"[DEBUG] Removing legacy floating vehicle box: {lbl_text}")
                    self.remove_box(key)

        if hasattr(self, 'fixed_vehicle_widget'):
            if hasattr(self.fixed_vehicle_widget, 'reset_fields'):
                self.fixed_vehicle_widget.reset_fields()
                self.fixed_vehicle_widget.clear_photo()
                
        # 6. [REMOVED] 메모 패널 초기화 방지 (사용자 요청)
                    
        # 6. 작업자 섹션 초기화
        for i in range(1, 11):
            if keep_workers:
                continue
            group = getattr(self, f'worker_group{i}', None)
            if group:
                if hasattr(group, 'participation_var'):
                    group.participation_var.set(True)
                group.cb_name.set('')
                group.cb_shift.set('주간')
                group.ent_worktime.set('')
                group.ent_ot.delete(0, tk.END)

    def add_daily_usage_entry(self):
        if hasattr(self, 'inspection_batch') and self.inspection_batch.plans:
            self.inspection_batch.save()
            return
        if self.ndt_loc_type_var.get() == '핫태핑':
            if self.cb_daily_test_method.get().strip().upper() not in ('MT', 'PT'):
                messagebox.showwarning('핫태핑', '핫태핑 길이표는 MT/PT 검사에 적용됩니다.')
                return
            if not self.hot_tap_location_var.get().strip():
                messagebox.showwarning('핫태핑', '1·2차 기록을 연결할 개소명/번호를 입력하세요.')
                return
            try:
                counts = [float(v.get().strip().replace(',', '') or '0') for v in
                          (self.ndt_ori_joint_var, self.ndt_rep_joint_var)]
                if any(not math.isfinite(n) or n < 0 or not n.is_integer() for n in counts) or sum(counts) <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning('핫태핑', 'ORI/REP 조인트수에 검사 개소 수를 양의 정수로 입력하세요.')
                return
            self._refresh_ndt_quantity()
        """현장별 일일 사용량 단일 등록 (리팩토링 버전)"""
        try:
            # 1. 기본 정보 및 유효성 검사
            date_val = self.ent_daily_date.get_date()
            site = self.cb_daily_site.get().strip()
            if not site:
                messagebox.showwarning("입력 오류", "현장명을 입력해주세요.")
                return

            # [NEW VALIDATION START]
            # 1-1. 차량 운행 유효성 검사
            if hasattr(self, 'vehicle_widget'):
                veh_data = self.vehicle_widget.get_data()
                if veh_data.get('vehicle_info'):
                    if not veh_data.get('mileage'):
                        if not messagebox.askyesno("입력 확인", "차량 정보는 선택되었으나 주행거리가 입력되지 않았습니다.\n이대로 저장하시겠습니까?"):
                            return

            # 1-2. 투입 인원 대비 검사 물량 검사
            has_workers = False
            if hasattr(self, 'worker_groups'):
                from site_apps.central.src.worker_selection import worker_is_selected
                for wg in self.worker_groups:
                    if worker_is_selected(wg) and wg.get_worker() and wg.get_time():
                        has_workers = True
                        break
            
            has_ndt_qty = False
            if hasattr(self, 'ndt_company_entries'):
                for comp_dict in self.ndt_company_entries:
                    for mat, entry_widget in comp_dict.items():
                        if mat == '_company': continue
                        try:
                            qty = float(entry_widget.get().strip() or 0)
                            if qty > 0:
                                has_ndt_qty = True
                                break
                        except: pass
                    if has_ndt_qty: break

            if has_workers and not has_ndt_qty:
                if not messagebox.askyesno("입력 확인", "투입된 인원(작업시간)은 있으나 검사 물량(NDT 수량)이 모두 0입니다.\n결측치일 가능성이 있습니다. 이대로 저장하시겠습니까?"):
                    return
            # [NEW VALIDATION END]

            # 2. 품목 ID 확인 (디스플레이 명칭 기반 매칭)
            mat_display = self.cb_daily_material.get().strip()
            mat_id = ""
            current_method = self.cb_daily_test_method.get().strip().upper()
            if current_method in ('PAUT', 'MT'):
                # PAUT(길이검사) 및 MT(약품별도기입)는 MaterialID에 자재ID/장비명을 기입하지 않음
                mat_id = ""
            elif mat_display:
                from site_apps.central.src.material_selection import resolve_daily_material
                selected_material_id = resolve_daily_material(self, mat_display)
                if selected_material_id is not None:
                    selected_rows = self.materials_df[self.materials_df['MaterialID'].apply(self.normalize_id) == self.normalize_id(selected_material_id)]
                    m_name = str(selected_rows.iloc[0].get('품목명', '')).strip()
                    if self._is_consumable_material(m_name, current_method):
                        mat_id = selected_material_id

                # [FIX] Only auto-register if it's a CONSUMABLE (Drug/Film). Equipment should NOT be registered.
                if not mat_id and mat_display:
                    if self._is_consumable_material(mat_display, self.cb_daily_test_method.get().strip()):
                        # [UX IMPROVEMENT] 등록되지 않은 소모품 감지 시, 사용자의 승인을 받도록 알림 팝업 추가
                        if messagebox.askyesno("신규 자재 등록", f"등록되지 않은 소모품 자재명('{mat_display}')입니다.\n재고 마스터에 신규 자재로 등록하시겠습니까?\n\n(아니오를 누르면 작업일보에는 저장되지만 재고 연동(차감) 대상에서 제외됩니다.)"):
                            mat_id = self.register_new_material(mat_display, warehouse='현장', 규격='자동등록')
                        else:
                            mat_id = mat_display
                    else:
                        # For equipment or non-consumables, do not store in MaterialID
                        mat_id = ""

            # Prepare first; replacement and persistence share one rollback boundary.
            from site_apps.central.src.views.daily_record_save import save_prepared
            plan = self._add_single_usage_record_logic(mat_id, date_val, site, prepare_only=True)
            saved_count = save_prepared(self, plan)
            if saved_count > 0:
                # [SMART VISIBILITY] Adjust filters to ensure the NEW record is visible
                try:
                    current_start = self.ent_daily_start_date.get_date()
                    current_end = self.ent_daily_end_date.get_date()
                    
                    if date_val < current_start:
                        self.ent_daily_start_date.set_date(date_val)
                    if date_val > current_end:
                        self.ent_daily_end_date.set_date(date_val)
                except: pass
                
                # 성공 시 필드 초기화 및 뷰 갱신
                if hasattr(self, 'ndt_company_entries'):
                    for company_entries in self.ndt_company_entries:
                        for name in self.ndt_materials_all:
                            if name in company_entries:
                                company_entries[name].delete(0, tk.END)
                
                # RTK 필드 초기화
                for cat, ent in self.rtk_entries.items():
                    if cat != "총계":
                        ent.delete(0, tk.END)
                self.rtk_entries["총계"].config(state='normal')
                self.rtk_entries["총계"].delete(0, tk.END)
                self.rtk_entries["총계"].config(state='readonly')
                
                # [V13_RESET_REQUESTED] Clearing fields after successful entry based on user request
                # [FIX] Keep the current date for consecutive entries on the same day
                self.clear_daily_usage_form_all(keep_date=True, keep_workers=True)

                # Reset focus to first logical field (Company)
                self.cb_daily_company.focus_set()

                # [V13.1] Keep workers too for consecutive entries
                # for i in range(1, 11):
                #     group = getattr(self, f'worker_group{i}', None)
                #     if group:
                #         group.cb_name.set('')
                #         group.cb_shift.set('')
                #         group.ent_worktime.set('')
                #         group.ent_ot.delete(0, tk.END)

                # NDT 섹션 초기화
                if hasattr(self, 'ndt_company_entries'):
                    while len(self.ndt_company_entries) > 1:
                        self.remove_last_ndt_company()
                    if self.ndt_company_entries:
                        first = self.ndt_company_entries[0]
                        first['_company'].set('')
                        for k in self.ndt_materials_all:
                            if k in first: first[k].delete(0, tk.END)

                # 차량 점검 필드 초기화 (safeguarded)
                # [NEW] 현장별 저장 시 차량, 목록, 메모 레이블창 자동 닫기 (User request)
                for key in list(getattr(self, 'vehicle_inspections', {}).keys()):
                    self.remove_box(key)
                for key in list(getattr(self, 'checklists', {}).keys()):
                    self.remove_box(key)
                for key in list(getattr(self, 'memos', {}).keys()):
                    self.remove_box(key)

                self.update_daily_usage_view()
                self.update_recent_entries_view()
                self.refresh_inquiry_filters()
                self.update_stock_view()
                
                # [NEW] Scroll to the bottom to see new entries
                if hasattr(self, 'daily_usage_tree'):
                    self.daily_usage_tree.yview_moveto(1.0)
                
                messagebox.showinfo("성공", f"{saved_count}건의 기록이 안전하게 저장되었습니다.")
                self._recover_combobox_interaction()
        except Exception as e:
            messagebox.showerror("오류", f"기록 저장 중 오류 발생: {e}")
            self._recover_combobox_interaction()

    def _recover_combobox_interaction(self):
        """저장 직후 드롭다운이 열리지 않는 현상 방지: 남은 입력 점유(grab)·디바운스 상태를 정리한다."""
        try:
            # 1. 이미 사라졌거나 모달이 아닌 위젯이 잡고 있는 grab 해제
            holder = self.root.grab_current()
            if holder is not None:
                keep = False
                try:
                    keep = holder.winfo_exists() and holder.winfo_ismapped() and holder is not self.root
                except Exception:
                    keep = False
                if not keep:
                    try: holder.grab_release()
                    except Exception: pass
            # 2. 예약된 자동 제안(디바운스)·열림 플래그 초기화
            for name in ('cb_daily_company', 'cb_daily_site', 'cb_daily_material', 'cb_daily_equip',
                         'cb_daily_test_method', 'cb_daily_unit', 'ent_daily_inspection_item',
                         'ent_daily_applied_code'):
                cb = getattr(self, name, None)
                if cb is None:
                    continue
                try:
                    after_id = getattr(cb, '_suggest_after_id', None)
                    if after_id:
                        cb.after_cancel(after_id)
                        cb._suggest_after_id = None
                    cb._auto_opening = False
                    cb._just_selected = False
                except Exception:
                    pass
            # 3. 포커스를 메인 창으로 복귀시켜 첫 클릭이 드롭다운으로 전달되게 함
            self.root.after_idle(lambda: self.root.focus_force())
        except Exception as e:
            print(f"[WARN] 콤보박스 상태 복구 실패: {e}")

    def load_report_mapping(self):
        """Load report mapping from JSON file or return defaults"""
        mapping_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'resources', 'report_mapping.json')
        if not os.path.exists(mapping_path):
            # Return defaults matching the manager's default_mapping
            return {
                'header': {'date': 'F2'},
                'general': {
                    'company': 'E5', 'project_name': 'E6', 'standard': 'E7', 'equipment': 'E8',
                    'report_no': 'K5', 'inspection_item': 'K6', 'inspector': 'K7', 'car_no': 'K8'
                },
                'methods': {
                    'RT': {'row': '12'}, 'UT': {'row': '13'}, 'MT': {'row': '14'}, 'PT': {'row': '15'},
                    'HT': {'row': '16'}, 'VT': {'row': '17'}, 'LT': {'row': '18'}, 'ET': {'row': '19'},
                    'PAUT': {'row': '20'}
                },
                'rtk': {
                    'center_miss': 'C32', 'density': 'E32', 'marking_miss': 'G32', 'film_mark': 'I32',
                    'handling': 'K32', 'customer_complaint': 'M32', 'etc': 'O32', 'total': 'Q32'
                },
                'ot': {
                    'row1_name': 'B38', 'row1_hours': 'K38',
                    'row2_name': 'B39', 'row2_hours': 'K39'
                },
                'materials': {
                    'RT T200': '43', 'RT AA400': '44', 'MT WHITE': '46', 'MT 7C-BLACK': '47', 'MT FLUORESCENT': '48',
                    'PT Penetrant': '49', 'PT Cleaner': '50', 'PT Developer': '51'
                }
            }
        try:
            with open(mapping_path, 'r', encoding='utf-8') as f:
                import json
                mapping = json.load(f)
                
                # Robust cleaning of method rows (e.g. 'E13' -> 13)
                if 'methods' in mapping:
                    import re
                    for m_key, m_val in mapping['methods'].items():
                        if isinstance(m_val, dict) and 'row' in m_val:
                            row_str = str(m_val['row'])
                            match = re.search(r'\d+', row_str)
                            if match: m_val['row'] = int(match.group())
                return mapping
        except:
            return {}

    def save_report_mapping(self, mapping):
        """Save report mapping to JSON file"""
        resources_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'resources')
        if not os.path.exists(resources_dir): 
            try: os.makedirs(resources_dir)
            except: pass
        mapping_path = os.path.join(resources_dir, 'report_mapping.json')
        try:
            print(f"DEBUG: Saving report mapping to {mapping_path}")
            with open(mapping_path, 'w', encoding='utf-8') as f:
                json.dump(mapping, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"DEBUG: Save failed: {e}")
            messagebox.showerror("저장 오류", f"매핑 설정을 저장하지 못했습니다: {e}")
            return False

    def open_report_mapping_dialog(self):
        """Open a dialog to configure Excel mapping for Daily Work Report"""
        dialog = tk.Toplevel(self.root)
        dialog.title("작업일보 엑셀 매핑 설정")
        dialog.geometry("600x800")
        dialog.transient(self.root)
        dialog.grab_set()

        mapping = self.load_report_mapping()
        
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill='both', expand=True)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True)

        # Helper to create grid of entries
        def create_entry_grid(parent, fields, current_data):
            entries = {}
            for i, (label_text, key) in enumerate(fields):
                ttk.Label(parent, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky='w')
                ent = ttk.Entry(parent, width=15)
                # Ensure we handle missing keys or nested dicts in current_data
                val = current_data.get(key, '')
                if isinstance(val, dict): val = val.get('row', '')
                ent.insert(0, str(val))
                ent.grid(row=i, column=1, padx=5, pady=5, sticky='w')
                entries[key] = ent
            return entries

        # Tab 1: General & Header
        tab1 = ttk.Frame(notebook, padding=10)
        notebook.add(tab1, text="일반/상단")
        gen_fields = [
            ("날짜 (Date)", "date_header"),
            ("업체명 (Company)", "company"), ("공사명 (Project)", "project_name"),
            ("적용규격 (Standard)", "standard"), ("장비명 (Equipment)", "equipment"),
            ("성적서번호 (Report No)", "report_no"), ("검사품명 (Item)", "inspection_item"),
            ("검사자 (Inspector)", "inspector"), ("차량번호 (Car No)", "car_no")
        ]
        tab1_data = mapping.get('general', {}).copy()
        tab1_data['date_header'] = mapping.get('header', {}).get('date', 'F2')
        tab1_entries = create_entry_grid(tab1, gen_fields, tab1_data)

        # Tab 2: Methods (Row Numbers)
        tab2 = ttk.Frame(notebook, padding=10)
        notebook.add(tab2, text="검사공법(행)")
        method_keys = ["RT", "UT", "MT", "PT", "HT", "VT", "LT", "ET", "PAUT"]
        method_fields = [(f"{m} (시작 행 번호)", m) for m in method_keys]
        tab2_entries = create_entry_grid(tab2, method_fields, mapping.get('methods', {}))

        # Tab 3: RTK Results
        tab3 = ttk.Frame(notebook, padding=10)
        notebook.add(tab3, text="불량율(RTK)")
        rtk_keys = ['center_miss', 'density', 'marking_miss', 'film_mark', 'handling', 'customer_complaint', 'etc', 'total']
        rtk_fields = [
            ("센터미스", "center_miss"), ("농도", "density"), ("마킹미스", "marking_miss"),
            ("필름마크", "film_mark"), ("취급부주의", "handling"), ("고객불만", "customer_complaint"),
            ("기타", "etc"), ("총계", "total")
        ]
        tab3_entries = create_entry_grid(tab3, rtk_fields, mapping.get('rtk', {}))

        # Tab 4: Materials (Row Numbers)
        tab4 = ttk.Frame(notebook, padding=10)
        notebook.add(tab4, text="자재(행)")
        mat_keys = list(mapping.get('materials', {}).keys())
        if not mat_keys: mat_keys = ['RT T200', 'RT AA400', 'MT WHITE', 'MT 7C-BLACK', 'MT FLUORESCENT', 'PT Penetrant', 'PT Cleaner', 'PT Developer']
        mat_fields = [(f"{m} (행 번호)", m) for m in mat_keys]
        tab4_entries = create_entry_grid(tab4, mat_fields, mapping.get('materials', {}))

        def save_and_close():
            try:
                new_mapping = {
                    'header': {'date': tab1_entries['date_header'].get().strip()},
                    'general': {k: tab1_entries[k].get().strip() for k in ['company', 'project_name', 'standard', 'equipment', 'report_no', 'inspection_item', 'inspector', 'car_no']},
                    'methods': {m: {'row': tab2_entries[m].get().strip()} for m in method_keys},
                    'rtk': {k: tab3_entries[k].get().strip() for k in rtk_keys},
                    'ot': mapping.get('ot', {
                        'row1_name': 'B38', 'row1_hours': 'K38',
                        'row2_name': 'B39', 'row2_hours': 'K39'
                    }),
                    'materials': {k: tab4_entries[k].get().strip() for k in mat_keys}
                }
                
                if self.save_report_mapping(new_mapping):
                    messagebox.showinfo("저장 완료", "매핑 설정이 저장되었습니다.")
                    dialog.destroy()
            except Exception as e:
                messagebox.showerror("오류", f"설정 구성 중 오류가 발생했습니다: {e}")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        ttk.Button(btn_frame, text="✅ 설정 저장", command=save_and_close).pack(side='right', padx=5)
        ttk.Button(btn_frame, text="❌ 취소", command=dialog.destroy).pack(side='right', padx=5)

    def export_daily_work_report(self, *args, **kwargs):
        if "export" in "export_daily_work_report" or "excel" in "export_daily_work_report":
            from site_apps.central.src.services.excel_exporter import export_daily_work_report_impl
        else:
            from site_apps.central.src.services.data_loader import export_daily_work_report_impl
        return export_daily_work_report_impl(self, *args, **kwargs)

    def export_central_daily_work_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_central_daily_work_report_impl
        return export_central_daily_work_report_impl(self, *args, **kwargs)

    def _get_usage_session_data(self):
        """Helper to collect Site tab entry data"""
        from site_apps.central.src.aux_equipment import value as auxiliary_equipment_value
        from site_apps.central.src.material_selection import selected_id
        data = {
            'date': str(self.ent_daily_date.get_date()),
            'site': self.cb_daily_site.get().strip(),
            'company': self.cb_daily_company.get().strip(),
            'applied_code': self.ent_daily_applied_code.get().strip(),
            'report_no': self.ent_daily_report_no.get().strip(),
            'equip': self.cb_daily_equip.get().strip(),
            'aux_equipment': auxiliary_equipment_value(self),
            'material': self.cb_daily_material.get().strip(),
            'material_id': selected_id(self),
            'inspection_item': self.ent_daily_inspection_item.get().strip(),
            'method': self.cb_daily_test_method.get().strip(),
            'unit': self.cb_daily_unit.get().strip(),
            'amount': self.ent_daily_test_amount.get().strip(),
            'unit_price': self.ent_daily_unit_price.get().strip(),
            'travel_cost': self.ent_daily_travel_cost.get().strip(),
            'meal_cost': self.ent_daily_meal_cost.get().strip(),
            'test_fee': self.ent_daily_test_fee.get().strip(),
            'note': self.ent_daily_note.get().strip(),
            'main_memo': self.main_memo_text.get('1.0', 'end-1c').strip() if hasattr(self, 'main_memo_text') and self.main_memo_text else "",
            'hot_tapping': {name: getattr(self, name).get() for name in ('ndt_loc_type_var', 'hot_tap_stage_var', 'hot_tap_size_var', 'hot_tap_location_var')},
            'workers': []
        }
        for i in range(1, 11):
            group = getattr(self, f'worker_group{i}', None)
            if group:
                data['workers'].append({
                    'name': group.cb_name.get().strip(),
                    'shift': group.cb_shift.get().strip(),
                    'time': group.ent_worktime.get().strip(),
                    'ot': group.ent_ot.get().strip(),
                    'participating': group.participation_var.get() if hasattr(group, 'participation_var') else True
                })
        data['ndt_companies'] = []
        if hasattr(self, 'ndt_company_entries'):
            for entries in self.ndt_company_entries:
                co_data = {'_company': entries['_company'].get().strip()}
                for mat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]:
                    co_data[mat] = entries[mat].get().strip()
                data['ndt_companies'].append(co_data)
        if hasattr(self, 'inspection_batch'):
            data['inspection_batch'] = self.inspection_batch.serialize()
        data['rtk'] = {cat: ent.get().strip() for cat, ent in self.rtk_entries.items() if cat != '총계'}
        
        # Include the primary vehicle and every additional tab in editor snapshots.
        data['vehicles'] = []
        if hasattr(self, 'vehicle_panel'):
            data['vehicles'] = self.vehicle_panel.serialize()
        elif hasattr(self, 'vehicle_boxes'):
            for box in self.vehicle_boxes:
                data['vehicles'].append(box.get_data())
                
        return data

    def _set_usage_session_data(self, data):
        """Helper to restore Site tab entry data"""
        if not data: return
        self.ent_daily_date.set_date(data.get('date', datetime.datetime.now().strftime('%Y-%m-%d')))
        self._last_daily_date = self.ent_daily_date.get_date()
        # [FIX] Use the dedicated Company field
        if 'applied_code' in data:
            self.ent_daily_applied_code.delete(0, tk.END)
            self.ent_daily_applied_code.insert(0, data.get('applied_code', ''))
        
        from site_apps.central.src.aux_equipment import restore
        restore(self, data.get('aux_equipment', ''))
        self.cb_daily_company.set(data.get('company', ''))
        self.cb_daily_site.set(data.get('site', ''))
        self.ent_daily_report_no.delete(0, tk.END); self.ent_daily_report_no.insert(0, data.get('report_no', ''))
        if isinstance(self.cb_daily_material, ttk.Combobox):
            self.cb_daily_material.set(data.get('material', ''))
        else:
            self.cb_daily_material.delete(0, tk.END)
            self.cb_daily_material.insert(0, data.get('material', ''))
        from site_apps.central.src.material_selection import remember_selection
        remember_selection(self, data.get('material_id'))
        if 'inspection_item' in data:
            self.ent_daily_inspection_item.delete(0, tk.END)
            self.ent_daily_inspection_item.insert(0, data.get('inspection_item', ''))
        self.cb_daily_test_method.set(data.get('method', ''))
        self.cb_daily_unit.set(data.get('unit', '매'))
        self.ent_daily_test_amount.delete(0, tk.END); self.ent_daily_test_amount.insert(0, data.get('amount', ''))
        self.ent_daily_unit_price.delete(0, tk.END); self.ent_daily_unit_price.insert(0, data.get('unit_price', ''))
        self.ent_daily_travel_cost.delete(0, tk.END); self.ent_daily_travel_cost.insert(0, data.get('travel_cost', ''))
        self.ent_daily_meal_cost.delete(0, tk.END); self.ent_daily_meal_cost.insert(0, data.get('meal_cost', ''))
        self.ent_daily_test_fee.delete(0, tk.END); self.ent_daily_test_fee.insert(0, data.get('test_fee', ''))
        self.ent_daily_note.delete(0, tk.END); self.ent_daily_note.insert(0, data.get('note', ''))
        if hasattr(self, 'main_memo_text') and self.main_memo_text:
            try:
                self.main_memo_text.delete('1.0', tk.END)
                if 'main_memo' in data and data['main_memo']:
                    self.main_memo_text.insert(tk.END, data['main_memo'])
            except: pass
        for name, value in data.get('hot_tapping', {}).items():
            if name in ('ndt_loc_type_var', 'hot_tap_stage_var', 'hot_tap_size_var', 'hot_tap_location_var'):
                getattr(self, name).set(value)
        if hasattr(self, 'inspection_batch') and 'inspection_batch' in data:
            self.inspection_batch.restore(data['inspection_batch'])
        workers = data.get('workers', [])
        for i, w_data in enumerate(workers, 1):
            group = getattr(self, f'worker_group{i}', None)
            if group:
                if hasattr(group, 'participation_var'):
                    group.participation_var.set(w_data.get('participating', True))
                group.cb_name.set(w_data.get('name', ''))
                group.cb_shift.set(w_data.get('shift', ''))
                group.ent_worktime.set(w_data.get('time', ''))
                group.ent_ot.delete(0, tk.END); group.ent_ot.insert(0, w_data.get('ot', ''))
        ndt_data = data.get('ndt_companies', [])
        while len(self.ndt_company_entries) > 1: self.remove_last_ndt_company()
        for i, co_data in enumerate(ndt_data):
            if i >= len(self.ndt_company_entries): self.add_ndt_company_section()
            entries = self.ndt_company_entries[i]
            entries['_company'].set(co_data.get('_company', ''))
            for mat in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]:
                entries[mat].delete(0, tk.END); entries[mat].insert(0, co_data.get(mat, ''))
        rtk = data.get('rtk', {})
        for cat, val in rtk.items():
            if cat in self.rtk_entries:
                ent = self.rtk_entries[cat]; ent.delete(0, tk.END); ent.insert(0, val)
        self.calculate_rtk_total()
        
        # [NEW] Restore vehicle inspection data
        vehicle_data = data.get('vehicles', [])
        if hasattr(self, 'vehicle_panel'):
            self.vehicle_panel.restore(vehicle_data)
        elif vehicle_data:
            # Match current boxes or add new ones if needed
            for i, v_state in enumerate(vehicle_data):
                if i < len(self.vehicle_boxes):
                    self.vehicle_boxes[i].set_data(v_state)
                else:
                    # If more boxes are needed than current, we could add them,
                    # but usually there's 1 or the count is already correct.
                    pass

    def _get_inout_session_data(self):
        """Helper to collect In/Out tab entry data"""
        data = {
            'reg': {
                'co_code': self.cb_co_code.get().strip(),
                'eq_code': self.cb_eq_code.get().strip(),
                'item_name': self.cb_item_name.get().strip(),
                'model': self.cb_model.get().strip(),
                'sn': self.ent_sn.get().strip(),
                'class': self.cb_class.get().strip(),
                'spec': self.cb_spec.get().strip(),
                'unit': self.cb_unit.get().strip(),
                'supplier': self.cb_supplier.get().strip(),
                'mfr': self.cb_mfr.get().strip(),
                'origin': self.cb_origin.get().strip(),
                'reorder': self.ent_reorder.get().strip(),
                'init': self.ent_init.get().strip(),
                'price': self.ent_price.get().strip(),
                'cost': self.ent_cost.get().strip()
            },
            'trans': {
                'material': self.cb_material.get().strip(),
                'type': self.cb_type.get().strip(),
                'qty': self.ent_qty.get().strip(),
                'site': self.cb_trans_site.get().strip(),
                'warehouse': self.cb_warehouse.get().strip(),
                'user': self.ent_user.get().strip(),
                'note': self.ent_note.get().strip()
            }
        }
        return data

    def _set_inout_session_data(self, data):
        """Helper to restore In/Out tab entry data"""
        if not data: return
        reg = data.get('reg', {})
        self.cb_co_code.set(reg.get('co_code', ''))
        self.cb_eq_code.set(reg.get('eq_code', ''))
        self.cb_item_name.set(reg.get('item_name', ''))
        self.cb_model.set(reg.get('model', ''))
        self.ent_sn.delete(0, tk.END); self.ent_sn.insert(0, reg.get('sn', ''))
        self.cb_class.set(reg.get('class', ''))
        self.cb_spec.set(reg.get('spec', ''))
        self.cb_unit.set(reg.get('unit', ''))
        self.cb_supplier.set(reg.get('supplier', ''))
        self.cb_mfr.set(reg.get('mfr', ''))
        self.cb_origin.set(reg.get('origin', ''))
        self.ent_reorder.delete(0, tk.END); self.ent_reorder.insert(0, reg.get('reorder', '0'))
        self.ent_init.delete(0, tk.END); self.ent_init.insert(0, reg.get('init', '0'))
        self.ent_price.delete(0, tk.END); self.ent_price.insert(0, reg.get('price', '0'))
        self.ent_cost.delete(0, tk.END); self.ent_cost.insert(0, reg.get('cost', '0'))
        
        trans = data.get('trans', {})
        self.cb_material.set(trans.get('material', ''))
        self.cb_type.set(trans.get('type', 'OUT'))
        self.ent_qty.delete(0, tk.END); self.ent_qty.insert(0, trans.get('qty', ''))
        self.cb_trans_site.set(trans.get('site', ''))
        self.cb_warehouse.set(trans.get('warehouse', ''))
        self.ent_user.set(trans.get('user', ''))
        self.ent_note.delete(0, tk.END); self.ent_note.insert(0, trans.get('note', ''))

    def _get_budget_session_data(self):
        """Helper to collect Budget tab entry data"""
        data = {
            'main': {k: w.get().strip() for k, w in self.budget_widgets.items()},
            'labor_detail': self.labor_detail_widget.get_data(),
            'material_detail': self.material_detail_widget.get_data(),
            'expense_detail': self.expense_detail_widget.get_data()
        }
        return data

    def _set_budget_session_data(self, data):
        """Helper to restore Budget tab entry data"""
        if not data: return
        main = data.get('main', {})
        for k, val in main.items():
            if k in self.budget_widgets:
                w = self.budget_widgets[k]
                if hasattr(w, 'set'): w.set(val)
                else: w.delete(0, tk.END); w.insert(0, val)
        self.labor_detail_widget.set_data(data.get('labor_detail', {}))
        self.material_detail_widget.set_data(data.get('material_detail', {}))
        self.expense_detail_widget.set_data(data.get('expense_detail', {}))
        self._update_budget_kpis()

    def save_form_session(self):
        """Orchestrate saving session data from all entry tabs into one JSON"""
        try:
            global_data = {
                'usage': self._get_usage_session_data(),
                'inout': self._get_inout_session_data(),
                'budget': self._get_budget_session_data(),
                'saved_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            file_path = filedialog.asksaveasfilename(
                title="전체 세션 저장 (일괄)",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialfile=f"PMI_Global_Session_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(global_data, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("저장 완료", f"현장/입출고/실행예산 탭의 입력 내용이 일괄 저장되었습니다.\n{os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("오류", f"전체 세션 저장 중 오류 발생: {e}")

    def load_previous_day_data(self):
        """선택된 현장의 가장 최근 기록(전일 데이터)을 불러와 폼을 채웁니다."""
        site = self.cb_daily_site.get().strip()
        if not site:
            messagebox.showwarning("입력 오류", "먼저 현장명을 선택하거나 입력해주세요.")
            return
            
        if self.daily_usage_df.empty:
            messagebox.showinfo("알림", "저장된 기록이 없습니다.")
            return
            
        df_site = self.daily_usage_df[self.daily_usage_df['Site'] == site].copy()
        if df_site.empty:
            messagebox.showinfo("알림", f"'{site}' 현장의 이전 기록이 없습니다.")
            return
            
        # 가장 최근 날짜 찾기
        df_site['Date'] = pd.to_datetime(df_site['Date'])
        recent_date = df_site['Date'].max()
        df_recent = df_site[df_site['Date'] == recent_date]
        
        date_str = recent_date.strftime('%Y-%m-%d')
        if not messagebox.askyesno("전일 데이터 불러오기", f"'{site}' 현장의 가장 최근 기록({date_str})을 불러오시겠습니까?\n(기존 입력된 작업자/차량 정보는 덮어쓰기 됩니다.)"):
            return
            
        try:
            # 1. 작업자 정보 복원 (최초 레코드 기준)
            first_record = df_recent.iloc[0]
            # 0. 일반 정보 복원 (업체명, 장비명 등)
            def _set_val(widget, val):
                if hasattr(widget, 'set'): widget.set(val)
                else:
                    widget.delete(0, 'end')
                    widget.insert(0, val)

            _set_val(self.cb_daily_company, self.clean_nan(first_record.get('업체명', '')))
            _set_val(self.cb_daily_equip, self.clean_nan(first_record.get('장비명', '')))
            from site_apps.central.src.aux_equipment import restore
            restore(self, first_record.get('보조장비', ''))
            
            self.ent_daily_applied_code.delete(0, 'end')
            self.ent_daily_applied_code.insert(0, self.clean_nan(first_record.get('적용코드', '')))
            
            self.ent_daily_inspection_item.delete(0, 'end')
            self.ent_daily_inspection_item.insert(0, self.clean_nan(first_record.get('검사품명', '')))
            
            _set_val(self.cb_daily_test_method, self.clean_nan(first_record.get('검사방법', '')))
            _set_val(self.cb_daily_unit, self.clean_nan(first_record.get('Unit', first_record.get('단위', ''))))
            
            self.ent_daily_unit_price.delete(0, 'end')
            unit_price = self.clean_nan(first_record.get('단가', ''))
            if str(unit_price) in ["0", "0.0"]: unit_price = ""
            self.ent_daily_unit_price.insert(0, unit_price)

            self.ent_daily_report_no.delete(0, 'end')
            self.ent_daily_report_no.insert(0, self.clean_nan(first_record.get('성적서번호', '')))
            
            self.ent_daily_note.delete(0, 'end')
            self.ent_daily_note.insert(0, self.clean_nan(first_record.get('Note', first_record.get('비고', ''))))

            for i in range(1, 11):
                group = getattr(self, f'worker_group{i}', None)
                if group:
                    u_col = 'User' if i == 1 else f'User{i}'
                    t_col = 'WorkTime' if i == 1 else f'WorkTime{i}'
                    o_col = 'OT' if i == 1 else f'OT{i}'
                    m_col = 'Meal' if i == 1 else f'Meal{i}'
                    
                    worker = first_record.get(u_col, first_record.get('작업자' if i == 1 else f'작업자{i}', ''))
                    worktime = first_record.get(t_col, first_record.get('작업시간' if i == 1 else f'작업시간{i}', ''))
                    ot = first_record.get(o_col, first_record.get('OT시간' if i == 1 else f'OT시간{i}', ''))
                    meal = first_record.get(m_col, first_record.get('일비' if i == 1 else f'일비{i}', ''))
                    
                    group.set_worker(self.clean_nan(worker))
                    group.set_time(self.clean_nan(worktime))
                    group.set_ot(self.clean_nan(ot))
                    if hasattr(group, 'set_meal'):
                        group.set_meal(self.clean_nan(meal))
                
            # 2. 차량 정보 복원
            if hasattr(self, 'fixed_vehicle_widget'):
                v_no_str = self.clean_nan(first_record.get('VehicleNo', first_record.get('차량번호', '')))
                v_nos = [x.strip() for x in v_no_str.split("||")] if "||" in v_no_str else [v_no_str]
                
                if v_nos and v_nos[0]:
                    self.fixed_vehicle_widget.cb_vehicle_info.set(v_nos[0])
                    
                if len(v_nos) > 1:
                    existing_boxes = getattr(self, 'vehicle_boxes', [])
                    for idx, extra_v in enumerate(v_nos[1:]):
                        if not extra_v: continue
                        if idx < len(existing_boxes):
                            existing_boxes[idx].cb_vehicle_info.set(extra_v)
                        elif hasattr(self, 'add_vehicle_inspection_box'):
                            new_box = self.add_vehicle_inspection_box()
                            if new_box: new_box.cb_vehicle_info.set(extra_v)
                
            messagebox.showinfo("완료", f"{date_str}의 작업자 및 차량 정보가 성공적으로 불러와졌습니다.\n(검사 물량 및 자재 소모량은 오늘 기준에 맞게 새로 입력해주세요.)")
        except Exception as e:
            messagebox.showerror("오류", f"데이터를 불러오는 중 오류가 발생했습니다: {e}")

    def load_form_session(self):
        """Orchestrate loading session data back to all entry tabs from a JSON file"""
        try:
            file_path = filedialog.askopenfilename(
                title="전체 세션 불러오기 (일괄)",
                filetypes=[("JSON files", "*.json")]
            )
            if not file_path: return
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # Restore all tabs
            self._set_usage_session_data(data.get('usage'))
            self._set_inout_session_data(data.get('inout'))
            self._set_budget_session_data(data.get('budget'))
            messagebox.showinfo("불러오기 완료", "현장/입출고/실행예산 탭의 데이터가 일괄 복원되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"전체 세션 불러오기 중 오류 발생: {e}")

    def _auto_reconcile_and_register_ndt(self, date_val, site, ndt_data, workers, company_code=""):
        """Unified helper to auto-register missing consumables and create 'OUT' transactions"""
        ndt_product_map = self._load_ndt_product_map()
        ndt_groups = self.ndt_groups
        
        note_suffix = f" ({company_code})" if company_code else ""
        note = f"{site} 현장 사용 (자동 차감){note_suffix}"
        
        # Ensure date_val is datetime and handle date vs datetime objects
        if isinstance(date_val, str):
            date_obj = pd.to_datetime(date_val)
        else:
            date_obj = date_val
            
        # extract date part safely
        pure_date = date_obj.date() if hasattr(date_obj, 'date') else date_obj
        trans_date = datetime.datetime.combine(pure_date, datetime.datetime.now().time())
        
        for name, qty in ndt_data.items():
            if qty <= 0: continue
            
            # 1. Resolve Item Name (Group) and Model
            db_item_name = ndt_product_map.get(name, '')
            db_model_name = name
            if not db_item_name:
                for grp, members in ndt_groups.items():
                    if name in members:
                        db_item_name = grp
                        db_model_name = name # [FIX] Remove prefix for consistency with display logic
                        break
                if not db_item_name: db_item_name = "기타소모품"
            
            # 2. Find or Register (Highly flexible matching for NDT)
            target_item = db_item_name.replace(' ', '')
            
            # [FIX] Fill NaN with empty string before comparison to ensure reliable matching
            m_df = self.materials_df.copy()
            for col in ['품목명', '모델명']:
                if col in m_df.columns:
                    m_df[col] = m_df[col].fillna('').astype(str).str.replace(' ', '', regex=False)

            # Try to find existing item
            # First, check for exact (item+model) match
            target_model = db_model_name.replace(' ', '')
            exact_match = m_df[(m_df['품목명'] == target_item) & (m_df['모델명'] == target_model)]
            
            if not exact_match.empty:
                mat_id = exact_match.iloc[0]['MaterialID']
            else:
                # If no exact match, try matching by item name ONLY for NDT groups
                item_only_match = m_df[m_df['품목명'] == target_item]
                if not item_only_match.empty:
                    # Use the first available item in this group (e.g., generic "PT약품")
                    mat_id = item_only_match.iloc[0]['MaterialID']
                else:
                    # Truly missing - [REFINED] Do NOT register new one automatically
                    print(f"DEBUG: Skipping auto-registration for missing NDT item '{db_item_name}' (Model: {db_model_name})")
                    mat_id = None
            
            # 3. Create Transaction (Only if mat_id exists)
            if mat_id:
                self._create_manual_stock_transaction(trans_date, mat_id, 'OUT', qty, site, workers, note)

    def _create_manual_stock_transaction(self, date_val, mat_id, trans_type, qty, site, user, note):
        # [V23.1_WATERTIGHT_FILTER] Default to block everything unless it matches the whitelist
        try:
            if self.materials_df.empty:
                return # Can't identify item, don't save
                
            # Robust ID matching
            str_target_id = str(mat_id).strip().replace('.0', '')
            # Vectorized lookup for speed
            mask = self.materials_df['MaterialID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) == str_target_id
            mat_info = self.materials_df[mask]
            
            if mat_info.empty:
                # If we can't find it in DB, we don't know what it is. To be safe, DON'T save it to history.
                return
                
            row = mat_info.iloc[0]
            item_name = str(row.get('품목명', '')).strip()
            model_name = str(row.get('모델명', '')).strip()
            item_name_up = item_name.upper()
            model_name_up = model_name.upper()
            
            # 1. Check for Carestream (Partial Match, Case Insensitive)
            is_carestream = "CARESTREAM" in item_name_up or "CARESTREAM" in model_name_up
            
            # 2. Check for NDT Chemicals (More inclusive partial match for PT/MT/NDT)
            ndt_keywords = ["PT", "MT", "NDT", "침투제", "세척제", "현상제", "자분", "페인트"]
            is_ndt_chem = any(kw in item_name_up or kw in model_name_up for kw in ndt_keywords)
            
            # [CRITICAL] If not a whitelist item, exit NOW
            if not (is_carestream or is_ndt_chem):
                return
                
        except Exception as e:
            print(f"Filter error in _create_manual_stock_transaction: {e}")
            return # Block on error
            
        # --- IF WE GOT HERE, THE ITEM IS ALLOWED ---
        try:
            # Ensure quantity is negative for OUT
            signed_qty = -abs(float(qty)) if trans_type == 'OUT' else abs(float(qty))
            
            new_trans = {
                'Date': pd.to_datetime(date_val),
                'MaterialID': mat_id,
                'Type': trans_type,
                'Quantity': signed_qty,
                'User': user,
                'Site': site,
                'Note': note
            }
            self.transactions_df = pd.concat([self.transactions_df, pd.DataFrame([new_trans])], ignore_index=True)
        except: pass
    
    def refresh_inquiry_filters(self):
        """Unified method to populate inquiry filter dropdowns and sync with autocomplete lists"""
        try:
            import re

            marker_pattern = MARKER_PATTERN
            
            # 1. Collect and Merge Site Info
            raw_sites = set()
            if not self.daily_usage_df.empty and 'Site' in self.daily_usage_df.columns:
                raw_sites.update(self.daily_usage_df['Site'].dropna().astype(str).apply(self.normalize_site_name).tolist())
            if hasattr(self, 'sites'):
                raw_sites.update(self.sites)
            # [NEW] budget_df에 저장된 현장명도 포함
            if hasattr(self, 'budget_df') and not self.budget_df.empty and 'Site' in self.budget_df.columns:
                raw_sites.update(self.budget_df['Site'].dropna().astype(str).apply(self.normalize_site_name).tolist())
            unique_sites = sorted([s for s in raw_sites 
                                     if s and str(s).lower() != 'nan'
                                    and s not in getattr(self, 'hidden_sites', [])])  # [FIX] 숨긴 현장 제외
            
            # [ROBUST] Do NOT overwrite self.sites Master List with history.
            # self.sites is the "Managed List". unique_sites is for "Filters".
            # self.sites[:] = unique_sites # REMOVED: This was causing deleted sites to reappear
            
            if hasattr(self, 'cb_daily_filter_site'):
                self.cb_daily_filter_site['values'] = ['전체'] + unique_sites
                if not self.cb_daily_filter_site.get(): self.cb_daily_filter_site.set('전체')
            if hasattr(self, 'cb_filter_site_monthly'):
                self.cb_filter_site_monthly['values'] = ['전체'] + unique_sites
            # [FIX] 공사실행예산서 상단 현장 콤보박스 → budget_df 현장만 표시
            # (삭제 후 목록에서 즉시 사라지게 하기 위해 daily_usage_df 현장은 제외)
            if hasattr(self, 'cb_budget_site'):
                budget_sites = []
                if hasattr(self, 'budget_df') and not self.budget_df.empty and 'Site' in self.budget_df.columns:
                    budget_sites = sorted([
                        s for s in self.budget_df['Site'].dropna().astype(str).apply(self.normalize_site_name).unique()
                        if s and str(s).lower() != 'nan'
                    ])
                self.budget_sites[:] = budget_sites  # [FIX] 자동완성 목록 동기화
                current_val = self.cb_budget_site.get()
                self.cb_budget_site['values'] = budget_sites
                if current_val: self.cb_budget_site.set(current_val)
            # 하단 실적 조회 콤보박스 → 모든 현장(daily_usage_df + budget_df) 표시
            if hasattr(self, 'cb_budget_view_site'):
                self.cb_budget_view_site['values'] = unique_sites

            # 2. Collect Material Info (History Only)
            raw_materials = set()
            if not self.daily_usage_df.empty and 'MaterialID' in self.daily_usage_df.columns:
                unique_mat_ids = self.daily_usage_df['MaterialID'].dropna().unique()
                for mat_id in unique_mat_ids:
                    name = self.get_material_display_name(mat_id)
                    if name: raw_materials.add(name)
            
            unique_materials = sorted(list(raw_materials))
            
            if hasattr(self, 'cb_daily_filter_material'):
                # [RESTORED] Populate with history names to avoid 'empty' confusion, while still excluding master list
                self.cb_daily_filter_material['values'] = ['전체'] + unique_materials
                if not self.cb_daily_filter_material.get(): self.cb_daily_filter_material.set('전체')
            
            if hasattr(self, 'cb_daily_material'):
                # [NEW] Only update values if it's still a combobox (Site tab uses Entry now)
                if isinstance(self.cb_daily_material, ttk.Combobox):
                    self.cb_daily_material['values'] = unique_materials
            
            if hasattr(self, 'cb_daily_filter_site'):
                self.cb_daily_filter_site['values'] = ['전체'] + unique_sites
            if hasattr(self, 'cb_filter_material_monthly'):
                self.cb_filter_material_monthly['values'] = ['전체'] + unique_materials

            # 3. Equipment dropdown
            if hasattr(self, 'cb_daily_filter_equipment'):
                raw_equip = []
                if not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
                    raw_equip = self.daily_usage_df['장비명'].dropna().astype(str).str.strip().unique().tolist()
                
                # Combine with catalog if needed
                unique_equip = sorted(list(set([e for e in raw_equip if e and str(e).lower() != 'nan'])))
                self.cb_daily_filter_equipment['values'] = ['전체'] + unique_equip
                if not self.cb_daily_filter_equipment.get(): self.cb_daily_filter_equipment.set('전체')

            # 4. Worker dropdown
            if hasattr(self, 'cb_daily_filter_worker'):
                worker_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']
                all_workers_raw = set()
                if not self.daily_usage_df.empty:
                    for col in worker_cols:
                        if col in self.daily_usage_df.columns:
                            all_workers_raw.update(self.daily_usage_df[col].dropna().astype(str).unique().tolist())
                
                # Include static list
                if hasattr(self, 'users'): all_workers_raw.update(self.users)
                
                clean_workers_set = set()
                for w in all_workers_raw:
                    w_str = str(w).strip()
                    if not w_str or w_str.lower() in ['nan', '0.0', 'none']: continue
                    cleaned = marker_pattern.sub('', w_str).strip()
                    cleaned = " ".join(cleaned.split())
                    if cleaned: clean_workers_set.add(cleaned)

                unique_workers = sorted(list(clean_workers_set))
                # self.users[:] = unique_workers # REMOVED: This was causing deleted workers to reappear
                self.cb_daily_filter_worker['values'] = ['전체'] + unique_workers
                if not self.cb_daily_filter_worker.get(): self.cb_daily_filter_worker.set('전체')

            # 5. Vehicle dropdown
            raw_vehicles = set()
            if not self.daily_usage_df.empty and '차량번호' in self.daily_usage_df.columns:
                v_list = self.daily_usage_df['차량번호'].dropna().astype(str).str.strip().tolist()
                for v in v_list:
                    cleaned = self.clean_nan(v)
                    if cleaned: raw_vehicles.add(cleaned)
            if hasattr(self, 'vehicles'):
                raw_vehicles.update(self.vehicles)
                
            unique_vehicles = sorted(list([v for v in raw_vehicles if v and str(v).lower() != 'nan']))
            if hasattr(self, 'cb_daily_filter_vehicle'):
                self.cb_daily_filter_vehicle['values'] = ['전체'] + unique_vehicles
                if not self.cb_daily_filter_vehicle.get(): self.cb_daily_filter_vehicle.set('전체')
                
            # 6. Company dropdown (for entry form)
            raw_companies = set()
            if not self.daily_usage_df.empty and '업체명' in self.daily_usage_df.columns:
                raw_companies.update(self.daily_usage_df['업체명'].dropna().astype(str).str.strip().tolist())
            if hasattr(self, 'companies'):
                raw_companies.update(self.companies)
                
            unique_companies = sorted(list([c for c in raw_companies if c and str(c).lower() != 'nan']))
            if hasattr(self, 'cb_daily_company'):
                # New entries use the managed catalog; historical labels remain
                # available only in inquiry filters and restored record values.
                self.cb_daily_company['values'] = sorted({
                    str(c).strip() for c in getattr(self, 'companies', [])
                    if str(c).strip() and str(c).strip().lower() not in ('nan', 'none')
                })
            if hasattr(self, 'cb_daily_filter_company'):
                self.cb_daily_filter_company['values'] = ['전체'] + unique_companies
                if not self.cb_daily_filter_company.get(): self.cb_daily_filter_company.set('전체')
            if hasattr(self, 'cb_daily_filter_vehicle'):
                self.cb_daily_filter_vehicle['values'] = ['전체'] + unique_vehicles
                if not self.cb_daily_filter_vehicle.get(): self.cb_daily_filter_vehicle.set('전체')
        except Exception as e:
            print(f"ERROR in refresh_inquiry_filters: {e}")
        if hasattr(self, 'cb_sales_filter_site'):
            self.cb_sales_filter_site['values'] = ['전체'] + self.sites
            if not self.cb_sales_filter_site.get(): self.cb_sales_filter_site.set('전체')
        if hasattr(self, 'cb_trans_filter_vehicle'):
            self.cb_trans_filter_vehicle['values'] = ['전체'] + unique_vehicles

    def update_daily_usage_view(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_view import update_daily_usage_view_impl
        return update_daily_usage_view_impl(self, *args, **kwargs)

    def reset_daily_usage_filters(self):
        """Reset all daily usage history filters to default values"""
        try:
            # 1. Reset Dates
            start_date = datetime.datetime(2024, 1, 1)
            self.ent_daily_start_date.set_date(start_date)
            self.ent_daily_end_date.set_date(datetime.datetime.now())
            
            # 2. Reset Comboboxes
            filter_combos = [
                self.cb_daily_filter_company, self.cb_daily_filter_site, 
                self.cb_daily_filter_material, self.cb_daily_filter_equipment,
                self.cb_daily_filter_worker, self.cb_daily_filter_vehicle,
                self.cb_daily_filter_shift, getattr(self, 'cb_daily_filter_method', None)
            ]
            for combo in filter_combos:
                if combo is not None and hasattr(combo, 'set'):
                    combo.set('전체')
            
            # 3. Refresh View
            self.update_daily_usage_view()
            
        except Exception as e:
            print(f"Error resetting daily usage filters: {e}")

    def _auto_adjust_tree_columns(self, tree, expand_only=False):
        """Automatically adjust column widths to fit content"""
        import tkinter.font as tkfont
        
        # Use the actual font used in the Treeview (12pt Malgun Gothic)
        # This ensures measurement matches display perfectly.
        content_font = tkfont.Font(family="Malgun Gothic", size=12)
        heading_font = tkfont.Font(family="Malgun Gothic", size=12, weight="bold")
        
        # We only care about visible columns
        visible_cols = tree['displaycolumns']
        if not visible_cols or visible_cols == ('#all'):
             visible_cols = tree['columns']

        for col in visible_cols:
            # Maximum Overkill: Character Count Heuristic
            # Assume 40px per character + 200px buffer
            # This GUARANTEES visibility in any environment
            
            # Measure heading
            w = len(col) * 40 + 200
            
            # Measure content
            col_index = list(tree['columns']).index(col)
            for item in tree.get_children():
                val = tree.item(item, 'values')
                if val and col_index < len(val):
                    text_len = len(str(val[col_index]))
                    content_w = text_len * 40 + 200
                    if content_w > w:
                        w = content_w
            
            # Apply specific constraints
            # Ultra-Safe: No max limit for potentially long text columns
            if col in ['품목명', '작업자', '비고']:
                w = max(w, 300) # Minimum base
                if col == '품목명': w = max(w, 700)
                elif col == '작업자': w = max(w, 300)
                elif col == '비고': w = max(w, 600)
            else:
                # Other columns get a relaxed global max
                min_w = 200
                if col == '날짜': min_w = 400
                elif col == '현장': min_w = 300
                elif col == '장비명': min_w = 250
                elif col == '검사방법': min_w = 200
                
                w = max(min_w, min(w, 2000)) 
            
            # Smart Auto-Expand: Only grow if content requires it
            if expand_only:
                current_w = int(tree.column(col, 'width'))
                if w <= current_w:
                    continue  # Keep user's manual adjustment
            
            # All columns are now user-resizable
            tree.column(col, width=w, minwidth=20, stretch=False, anchor='center')
        
        # Force layout update
        tree.update_idletasks()

    def delete_daily_usage_entry(self):
        """Delete only selected daily records, retaining unrelated stock transactions."""
        selection = self.daily_usage_tree.selection()
        if not selection:
            return
        try:
            indices = set()
            for item in selection:
                tags = self.daily_usage_tree.item(item, 'tags')
                if not tags or int(tags[0]) not in self.daily_usage_df.index:
                    raise ValueError('기록이 변경되었습니다. 새로고침 후 다시 선택해주세요.')
                indices.add(int(tags[0]))
        except (ValueError, TypeError):
            messagebox.showwarning('삭제 오류', '선택한 기록을 확인할 수 없습니다. 새로고침해주세요.', parent=self.root)
            return
        if not messagebox.askyesno('삭제 확인',
                f'선택한 일일 작업 기록 {len(indices)}건을 삭제할까요?\n'
                '검사량·작업자·메모가 함께 삭제되고 재고 및 집계에 반영됩니다.\n'
                '별도 입출고 기록은 유지됩니다.', parent=self.root):
            return
        old_daily = self.daily_usage_df.copy(deep=True)
        old_transactions = self.transactions_df.copy(deep=True)
        cloud_deletes = []
        for idx in indices:
            row = old_daily.loc[idx]
            etime = str(row.get('EntryTime', '')).split('.')[0].strip()
            if etime and etime.lower() not in ('', 'nan', 'nat', 'none'):
                cloud_deletes.append(etime)
        try:
            # Auto-deduction transactions are already excluded by stock calculation.
            # Date/site matches cannot identify which daily job created a legacy transaction.
            self.daily_usage_df = self.daily_usage_df.drop(list(indices))
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.daily_usage_df = old_daily
            self.transactions_df = old_transactions
            messagebox.showerror('삭제 저장 실패', f'{error}\n원본 기록을 복원했습니다.', parent=self.root)
            return

        # [CLOUD SYNC] Propagate deletion to Google Apps Script and update local state immediately
        if cloud_deletes:
            try:
                import threading
                from site_apps.central.src.sync_google_sheet import push_batch_to_web_app, DEFAULT_INVENTORY_PATH, _load_sync_state, _save_sync_state
                try:
                    state = _load_sync_state(DEFAULT_INVENTORY_PATH)
                    cur_times = set(state.get("cloud_entry_times", [])) - set(cloud_deletes)
                    _save_sync_state(DEFAULT_INVENTORY_PATH, cur_times)
                except Exception as ex:
                    print(f"DEBUG: Failed to update local sync state on delete: {ex}")

                def _async_cloud_delete(deletes_list):
                    try:
                        push_batch_to_web_app(deletes=deletes_list)
                    except Exception as err:
                        print(f"DEBUG: Async cloud delete failed: {err}")
                threading.Thread(target=_async_cloud_delete, args=(cloud_deletes,), daemon=True).start()
            except Exception as e:
                print(f"DEBUG: Cloud delete dispatch error: {e}")
        self.update_daily_usage_view()
        self.update_recent_entries_view()
        self.update_stock_view()
        self.update_transaction_view()
        self.refresh_inquiry_filters()
        self.update_monthly_usage_view()
        if hasattr(self, 'budget_view_tree'):
            self.update_budget_site_view()

    def _calculate_ot_from_worktime(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import _calculate_ot_from_worktime_impl
        return _calculate_ot_from_worktime_impl(self, *args, **kwargs)

    def calculate_and_update_ot_manual(self, worktime_value, ot_entry, calculation_date=None):
        """Helper for personal calculation during editing without complex widget lookups"""
        try:
            current_date = calculation_date if calculation_date else self.ent_daily_date.get_date()
            for index in range(1, 11):
                group = getattr(self, f'worker_group{index}', None)
                if group is not None and group.ent_worktime is worktime_entry:
                    worktime_value = group.get_time()
                    break
            ot_hours, amount = self._calculate_ot_from_worktime(worktime_value, current_date)
            
            if amount > 0:
                ot_entry.delete(0, tk.END)
                ot_entry.insert(0, f"{amount:,}")
            elif ot_hours == 0:
                # If no OT, clear it or set to 0
                ot_entry.delete(0, tk.END)
                ot_entry.insert(0, "0")
        except: pass

    def _safe_set_daily_df(self, idx, col, val, is_numeric=False):
        """Defensively set a value in daily_usage_df handling all possible dtype conflicts"""
        if col not in self.daily_usage_df.columns:
            self.daily_usage_df.at[idx, col] = val
            return

        col_dtype = self.daily_usage_df[col].dtype
        is_stringy = pd.api.types.is_string_dtype(col_dtype) or pd.api.types.is_object_dtype(col_dtype)
        
        v_to_set = val
        if is_stringy and not is_numeric:
            # Force to cleaned string for string columns
            v_to_set = str(val).strip() if pd.notna(val) else ""
            if v_to_set in ('0.0', '0', 'nan', 'None', 'NaT'): v_to_set = ""
        elif is_numeric:
            # Force to float for numeric columns
            try:
                if isinstance(val, str):
                    v_to_set = float(val.replace(',', '').strip()) if val.strip() else 0.0
                else:
                    v_to_set = float(val) if pd.notna(val) else 0.0
            except:
                v_to_set = 0.0

        try:
            self.daily_usage_df.at[idx, col] = v_to_set
        except:
            try:
                # Type conflict (e.g. float into strict string column or vice versa)
                # Cast the column to object to allow the mixed/new type
                self.daily_usage_df[col] = self.daily_usage_df[col].astype(object)
                self.daily_usage_df.at[idx, col] = v_to_set
            except:
                # Final fallback
                self.daily_usage_df.at[idx, col] = str(v_to_set)

    def _is_consumable_material(self, *args, **kwargs):
        from site_apps.central.src.models.material_model import _is_consumable_material_impl
        return _is_consumable_material_impl(self, *args, **kwargs)

    def save_daily_usage_edits(self, df_idx, new_data):
        """Save edited daily usage and reconcile stock"""
        frame_names = ('daily_usage_df', 'transactions_df', 'materials_df', 'monthly_usage_df')
        snapshots = {name: getattr(self, name).copy(deep=True) for name in frame_names}
        saved = False
        try:
            old_entry = self.daily_usage_df.loc[df_idx]
            
            # Separate stock transactions cannot be linked by date/site alone.
            # Inventory deducts the edited DailyUsage row and excludes automatic mirrors.
            # 2. Update entry in DF
            # [FIX] Handle column name aliases (Date/날짜, Site/현장) to ensure existing columns are updated
            site_col = 'Site' if 'Site' in self.daily_usage_df.columns else '현장'
            date_col = 'Date' if 'Date' in self.daily_usage_df.columns else '날짜'
            
            # [NEW] Pre-collect existing column names to avoid duplicate-like column creation
            existing_cols = self.daily_usage_df.columns
            
            for k, v in new_data.items():
                target_key = k
                if k == 'Site': target_key = site_col
                elif k == 'Date': target_key = date_col
                
                # [FIX] For NDT and RTK fields, ensure they match normalized column names in the DF
                if k.startswith('NDT_') or k.startswith('RTK_'):
                    # Search for case-insensitive and space-insensitive match in existing columns
                    k_norm = k.replace(' ', '').upper()
                    for ex_col in existing_cols:
                        if str(ex_col).replace(' ', '').upper() == k_norm:
                            target_key = ex_col
                            break
                
                # Comprehensive list of columns that MUST be numeric
                numeric_cols = ['Usage', '수량', '검사량', '단가', '출장비', '일식', '검사비', 'FilmCount', 'OT', 'OT금액']
                is_numeric = k.startswith('NDT_') or k.startswith('RTK_') or k in numeric_cols or any(f"OT{i}" == k or f"OT금액{i}" == k for i in range(1, 11))
                
                # Use safe setter to avoid all dtype crashes
                self._safe_set_daily_df(df_idx, target_key, v, is_numeric=(is_numeric or k == 'MaterialID'))

            from site_apps.central.src.models.stock_calculation import synchronize_legacy_chemicals
            for column, value in synchronize_legacy_chemicals(old_entry, new_data).items():
                self._safe_set_daily_df(df_idx, column, value, is_numeric=False)

            # Ensure Usage is consistently updated from Inspection Amount (검사량)
            usage_col = 'Usage' if 'Usage' in self.daily_usage_df.columns else '수량'
            self._safe_set_daily_df(df_idx, usage_col, new_data.get('검사량', 0.0), is_numeric=True)
            if '검사량' in self.daily_usage_df.columns:
                self._safe_set_daily_df(df_idx, '검사량', new_data.get('검사량', 0.0), is_numeric=True)
            
            # [FIX] Initialize new_mat_display for stock reconciliation logic
            new_mat_display = new_data.get('Material', '')
            if not new_mat_display:
                new_mat_display = self.get_material_display_name(new_data.get('MaterialID', ''))
            
            # 3. Apply New Deduction (Selective Auto-deduction)
            new_date = pd.to_datetime(new_data['Date'])
            new_site = new_data['Site']
            # [NEW] PAUT 및 장비류는 재고 차감에서 제외
            method_raw = new_data.get('검사방법', '')
            method = str(method_raw).strip().upper()
            is_excluded = (method == 'PAUT')
            
            if not is_excluded:
                # Resolve or Register MaterialID
                new_mat_id = ""
                if new_mat_display:
                    for _, row in self.materials_df.iterrows():
                        if self.get_material_display_name(row['MaterialID']) == new_mat_display:
                            new_mat_id = row['MaterialID']
                            if row.get('Active', 1) == 0:
                                self.materials_df.loc[self.materials_df['MaterialID'] == new_mat_id, 'Active'] = 1
                            break
                    
                    if not new_mat_id and new_mat_display:
                        if self._is_consumable_material(new_mat_display, method):
                            # [UX IMPROVEMENT] 등록되지 않은 소모품 감지 시 승인 팝업 추가 (수정/저장 동일)
                            if messagebox.askyesno("신규 자재 등록", f"등록되지 않은 소모품 자재명('{new_mat_display}')입니다.\n재고 마스터에 신규 자재로 등록하시겠습니까?\n\n(아니오를 누르면 기록에는 변경되지만 재고 연동(차감) 대상에서 제외됩니다.)"):
                                new_mat_id = self.register_new_material(new_mat_display, warehouse='현장', 규격='자동등록', persist=False)
                            else:
                                new_mat_id = new_mat_display
                        else:
                            new_mat_id = new_mat_display
                
                self._safe_set_daily_df(df_idx, 'MaterialID', new_mat_id, is_numeric=False)
                new_qty = float(new_data.get('검사량', 0))
                new_note_pattern = f'{new_site} 현장 사용 (자동 차감)'
                
                workers_names = [new_data.get(f'User{i}' if i > 1 else 'User', '').strip() for i in range(1, 11)]
                all_workers = ", ".join([n for n in workers_names if n])
                
                if new_mat_id and new_qty > 0:
                    if self._is_consumable_material(new_mat_display, method):
                        mat_info = self.get_material_info(new_mat_id)
                        full_item_name = str(mat_info.get('품목명', '')).replace(' ', '').upper()
                        # PT/MT 약품 부모 항목 중복 차감 방지
                        if full_item_name not in ["PT약품", "MT약품", "NDT약품"]:
                            self._create_manual_stock_transaction(new_date, new_mat_id, 'OUT', new_qty, new_site, all_workers, new_note_pattern)
                
                # NDT 약품(소모품) 일괄 정산
                ndt_data = {
                    name: float(new_data.get(f'NDT_{"".join(name.split())}', 0))
                    for name in ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
                }
                self._auto_reconcile_and_register_ndt(new_date, new_site, ndt_data, all_workers, new_data.get('회사코드', ''))

            # 4. Save & Refresh
            if self.save_data(require_primary=True):
                saved = True
                self.update_daily_usage_view()
                self.refresh_inquiry_filters()
                self.update_stock_view()
                self.update_transaction_view()
                self.update_material_combo() # Refresh history suggestions
                self.update_recent_entries_view()
                self.update_monthly_usage_view()
                return True
            raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as e:
            if saved:
                messagebox.showwarning('저장 완료 · 화면 갱신 확인', f'기록은 저장되었습니다. 다시 조회해주세요.\n{e}')
                return True
            for name, snapshot in snapshots.items():
                setattr(self, name, snapshot)
            messagebox.showerror("수정 오류", f"기록 수정 중 오류가 발생했습니다: {e}")
            return False
                
    def export_daily_usage_history(self):
        """Export current filtered daily usage history to Excel"""
        try:
            # Get selected items only, or all if none selected
            selected_items = self.daily_usage_tree.selection()
            if not selected_items:
                # If no items selected, export all visible items
                items_to_export = self.daily_usage_tree.get_children()
            else:
                # Export only selected items
                items_to_export = selected_items
            
            data = []
            for item in items_to_export:
                data.append(self.daily_usage_tree.item(item, 'values'))
            
            if not data:
                messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
                return
            
            # Export columns currently visible in the Treeview
            display_cols = self.daily_usage_tree['displaycolumns']
            if not display_cols or display_cols == ('#all'):
                selected_cols = list(self.daily_usage_tree['columns'])
            else:
                selected_cols = list(display_cols)
            
            # Map selected column names to their indices in the Treeview values
            all_cols = list(self.daily_usage_tree['columns'])
            col_indices = [all_cols.index(col) for col in selected_cols]
            
            # Reconstruct data
            filtered_data = []
            for item_values in data:
                row = [item_values[i] for i in col_indices if i < len(item_values)]
                filtered_data.append(row)
            
            filename = f"일일사용내역_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                cols = list(self.daily_usage_tree['columns'])
                col_to_idx = {col: i for i, col in enumerate(cols)}
                
                # Rebuild data from 'data' (which has ALL columns including hidden ones)
                # and swap summarized '작업자' with full list from '(Full작업자)'
                final_filtered_data = []
                for item_values in data:
                    row_list = list(item_values)
                    
                    # Perform the swap if both columns exist
                    if '작업자' in cols and '(Full작업자)' in col_to_idx:
                        full_idx = col_to_idx['(Full작업자)']
                        if full_idx < len(row_list):
                            full_val = row_list[full_idx]
                            if full_val:
                                # Find index of '작업자' in ALL columns to swap in row_list
                                idx_worker = cols.index('작업자')
                                row_list[idx_worker] = full_val
                    
                    # Extract only the selected display columns for final export
                    final_row = [row_list[i] for i in col_indices if i < len(row_list)]
                    final_filtered_data.append(final_row)
                
                df = pd.DataFrame(final_filtered_data, columns=selected_cols)
                df = self.clean_df_export(df)
                self.save_df_to_excel_autofit(df, save_path, "일일사용내역")
                messagebox.showinfo("완료", f"데이터가 엑셀로 저장되었습니다.\n{save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")

    def show_column_visibility_dialog(self, tree=None):
        """Open dialog to manually show/hide columns in the specified tree view"""
        if tree is None:
            tree = self.daily_usage_tree
            
        all_cols = list(tree['columns'])
        
        # Mandatory columns identification
        # NDT/RTK columns are at the end, workers in middle, basic info at start
        selectable_cols = [c for c in all_cols if c != '(Full작업자)']
        
        # [FIX] Use stored manual preferences instead of current filtered displaycolumns
        # so that hidden-by-data columns still appear checked if user wants them.
        preferred_defaults = ['날짜', '구분', '검사방법', '품목명', '수량', '단위', '단가', '검사비', '작업자', '작업시간', 'OT금액', '장비명', '업체명', '비고']
        if tree == self.daily_usage_tree:
            active_cols = getattr(self, 'manual_visible_cols', [])
            if not active_cols:
                active_cols = list(preferred_defaults)
        else:
            active_cols = tree['displaycolumns']
            if not active_cols or active_cols == ('#all'):
                 active_cols = all_cols
        
        dialog = ColumnSelectionDialog(self.root, selectable_cols, title="표시 컬럼 관리")
        # Overwrite vars with current visibility state
        for col, var in dialog.vars.items():
            var.set(col in active_cols)
            
        dialog.wait_window()
        
        if dialog.result is not None:
            # Result set from dialog
            final_selection = list(dialog.result)
            
            # Reconstruct final display based on user result + mandatory ones
            mandatory_lookup = {
                self.daily_usage_tree: ['날짜', '구분', '검사방법', '품목명', '수량', '작업자'],
                self.inout_tree: ['날짜', '구분', '품목명', '장비명']
            }
            
            mandatory_cols = mandatory_lookup.get(tree, ['날짜'])
            for mc in mandatory_cols:
                if mc in all_cols and mc not in final_selection:
                    final_selection.append(mc)
            
            # Re-order according to preferred order first, then remaining
            if tree == self.daily_usage_tree:
                sorted_selection = [c for c in preferred_defaults if c in final_selection]
                for c in all_cols:
                    if c in final_selection and c not in sorted_selection:
                        sorted_selection.append(c)
            else:
                sorted_selection = [c for c in all_cols if c in final_selection]
            
            tree['displaycolumns'] = sorted_selection
            
            # Save configuration to the appropriate key
            if tree == self.daily_usage_tree:
                self.manual_visible_cols = sorted_selection
                if not hasattr(self, 'tab_config'): self.tab_config = {}
                self.tab_config['daily_usage_visible_cols'] = sorted_selection
            elif tree == self.inout_tree:
                if not hasattr(self, 'tab_config'): self.tab_config = {}
                self.tab_config['inout_visible_cols'] = sorted_selection
                
            self.save_tab_config()
            
            # Immediate refresh to apply selection
            if tree == self.daily_usage_tree:
                self.update_daily_usage_view()
            elif tree == self.inout_tree:
                self.update_transaction_view()

    def export_weekly_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_weekly_report_impl
        return export_weekly_report_impl(self, *args, **kwargs)

    def export_monthly_ndt_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_monthly_ndt_report_impl
        return export_monthly_ndt_report_impl(self, *args, **kwargs)

    def export_all_daily_usage(self):
        """Export all daily usage records to Excel"""
        try:
            if self.daily_usage_df.empty:
                messagebox.showinfo("알림", "기록된 데이터가 없습니다.")
                return
            
            # Export all columns from the tree structure
            selected_cols = list(self.daily_usage_tree['columns'])
            
            filename = f"전체사용기록_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=[("Excel files", "*.xlsx")])
            
            if save_path:
                # Ensure we only try to export columns that exist in the dataframe
                valid_cols = [c for c in selected_cols if c in self.daily_usage_df.columns]
                # If dataframe has extra columns not in tree (like raw data), include them? 
                # Better to stick to what the user sees/defines in tree interface + critical data
                # Actually, export_all usually implies backing up the raw dataframe. 
                # But previous logic used tree columns. Let's stick to tree columns if they exist in DF.
                
                # Fallback: if tree columns are not in DF (e.g. calculated columns), we might have issues.
                # The daily_usage_df has raw data. The tree has formatted data.
                # daily_usage_df columns: Date, Site, MaterialID, Usage, Note, etc. + RTK_..., NDT_...
                # Tree columns: 날짜, 현장, 작업자... OT합계...
                
                # If we want "All Daily Usage", we should probably export the RAW dataframe for backup purposes, 
                # OR the formatted view for reporting.
                # The previous code seemed to try to export based on tree columns but mapped from DF?
                # "export_df = self.daily_usage_df[selected_cols].copy()"
                # This implies 'selected_cols' MUST exist in daily_usage_df.
                # Let's check daily_usage_df columns again.
                # It has 'Date', 'Site', 'MaterialID'...
                # It does NOT have '날짜', '현장' (Korean names).
                # So the previous code 'self.daily_usage_df[selected_cols]' would have FAILED if selected_cols came from tree columns!
                
                # WAIT. The user said "remove column selection". 
                # The previous working code (before my changes) probably did:
                # export_df = self.daily_usage_df.copy()
                # Let's revert to a safe "Backup" style export for "All Data".
                
                export_df = self.daily_usage_df.copy()
                export_df = self.clean_df_export(export_df)
                self.save_df_to_excel_autofit(export_df, save_path, "전체사용기록")
                
                export_df = self.clean_df_export(export_df)
                self.save_df_to_excel_autofit(export_df, save_path, "전체기록")
                messagebox.showinfo("완료", f"전체 기록이 엑셀로 저장되었습니다.\n{save_path}")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 내보내기 중 오류가 발생했습니다: {e}")

    def export_invoice_excel(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_invoice_excel_impl
        return export_invoice_excel_impl(self, *args, **kwargs)



    def normalize_site_name(self, *args, **kwargs):
        from site_apps.central.src.utils.helpers import normalize_site_name_impl
        return normalize_site_name_impl(self, *args, **kwargs)

    def on_tab_changed(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_tab_changed_impl
        return on_tab_changed_impl(self, *args, **kwargs)
    
    def auto_save_to_list(self, event, combobox, data_list, config_key):
        """Helper to auto-save new entry from combobox to a list and update all related UI"""
        new_val = combobox.get().strip()
        if not new_val:
            return
            
        if new_val not in data_list:
            data_list.append(new_val)
            data_list.sort()
            
            # Trigger app-wide update
            self.refresh_ui_for_list_change(config_key)
            self.save_tab_config()


    def auto_save_worktime(self, event, entry, config_key):
        """Helper to auto-save worktime values and support copy functionality"""
        worktime_value = entry.get().strip()
        if not worktime_value:
            return
            
        # Initialize worktimes list if not exists
        if not hasattr(self, 'worktimes'):
            self.worktimes = []
            
        # Check if this value already exists
        if worktime_value not in self.worktimes:
            self.worktimes.append(worktime_value)
            self.worktimes.sort()
            
            # Save to config and update UI
            try:
                self.refresh_ui_for_list_change('worktimes')
                self.save_tab_config()
            except Exception as e:
                print(f"Failed to save worktime config: {e}")
        
        # Calculate and update corresponding OT field
        self.calculate_and_update_ot(worktime_value, entry)
        
        # Add copy functionality - select all text on focus and bind Ctrl+C (only once)
        entry.select_range(0, tk.END)
        if not hasattr(entry, '_copy_bound'):
            entry.bind('<Control-c>', lambda e: self.copy_worktime(entry))
            entry._copy_bound = True
        
        # Store current value for copy functionality
        entry.last_value = worktime_value

    def calculate_and_update_ot(self, worktime_value, worktime_entry):
        """Calculate OT amount based on worktime and update corresponding OT field"""
        try:
            current_date = self.ent_daily_date.get_date()
            for index in range(1, 11):
                group = getattr(self, f'worker_group{index}', None)
                if group is not None and getattr(group, 'ent_worktime', None) is worktime_entry:
                    worktime_value = group.get_time()
                    break
            ot_hours, amount = self._calculate_ot_from_worktime(worktime_value, current_date)
            
            ot_entry = self.get_corresponding_ot_field(worktime_entry)
            if ot_entry:
                if amount > 0:
                    ot_entry.delete(0, tk.END)
                    ot_entry.insert(0, f"{amount:,}")
                else:
                    ot_entry.delete(0, tk.END)
                    # If it was cleared, but hours > 0 (very small amount?), maybe set to 0
                    if ot_hours > 0: ot_entry.insert(0, "0")
        except Exception as e:
            print(f"Error calculating OT: {e}")

    def get_corresponding_ot_field(self, worktime_entry):
        """Get corresponding OT field for a worktime entry"""
        # Map worktime entries to OT entries
        worktime_to_ot = {
            'ent_worktime1': 'ent_ot1', 'ent_worktime2': 'ent_ot2', 
            'ent_worktime3': 'ent_ot3', 'ent_worktime4': 'ent_ot4',
            'ent_worktime5': 'ent_ot5', 'ent_worktime6': 'ent_ot6',
            'ent_worktime7': 'ent_ot7', 'ent_worktime8': 'ent_ot8',
            'ent_worktime9': 'ent_ot9', 'ent_worktime10': 'ent_ot10'
        }
            
        # Find worktime entry name by checking object IDs
        worktime_id = id(worktime_entry)
            
        worktime_attrs = [f'ent_worktime{i}' for i in range(1, 11)]
        for attr_name in worktime_attrs:
            if hasattr(self, attr_name):
                attr_value = getattr(self, attr_name)
                if id(attr_value) == worktime_id:
                    ot_attr_name = worktime_to_ot.get(attr_name)
                    if ot_attr_name and hasattr(self, ot_attr_name):
                        return getattr(self, ot_attr_name)
                    break
            
        return None

    def copy_worktime(self, entry):
        """Copy worktime value to clipboard"""
        try:
            worktime_value = entry.get().strip()
            if worktime_value:
                self.root.clipboard_clear()
                self.root.clipboard_append(worktime_value)
                # Show brief feedback
                entry.delete(0, tk.END)
                entry.insert(0, worktime_value)
        except Exception as e:
            pass  # Silently handle clipboard errors

    def auto_save_ot(self, event, entry, config_key):
        """Helper to auto-save OT values and support copy functionality"""
        ot_value = entry.get().strip()
        if not ot_value:
            return
                
        # Initialize ot_times list if not exists
        if not hasattr(self, 'ot_times'):
            self.ot_times = []
            
        # Calculate OT amount for any input
        calculated_amount = self.calculate_ot_amount(ot_value)
        if calculated_amount:
            # Update entry with calculated amount
            entry.delete(0, tk.END)
            display_value = f"{ot_value} ({calculated_amount:,}원)"
            entry.insert(0, display_value)
            ot_value = display_value
            print(f"OT calculated: {ot_value}")  # Debug print
            
        # Check if this value already exists
        if ot_value not in self.ot_times:
            self.ot_times.append(ot_value)
            self.ot_times.sort()
            
        # Save to config with error handling - only save when list changes
        try:
            self.save_tab_config()
        except Exception as e:
            print(f"Failed to save OT config: {e}")
            
        # Add copy functionality - select all text on focus and bind Ctrl+C (only once)
        entry.select_range(0, tk.END)
        if not hasattr(entry, '_copy_bound'):
            entry.bind('<Control-c>', lambda e: self.copy_ot(entry))
            entry._copy_bound = True
            
        # Store current value for copy functionality
        entry.last_value = ot_value

    def calculate_ot_amount(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import calculate_ot_amount_impl
        return calculate_ot_amount_impl(self, *args, **kwargs)

    def _parse_ot_hours(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import _parse_ot_hours_impl
        return _parse_ot_hours_impl(self, *args, **kwargs)

    def _calculate_split_ot_hours(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import _calculate_split_ot_hours_impl
        return _calculate_split_ot_hours_impl(self, *args, **kwargs)

    def copy_ot(self, entry):
        """Copy OT value to clipboard"""
        try:
            ot_value = entry.get().strip()
            if ot_value:
                self.root.clipboard_clear()
                self.root.clipboard_append(ot_value)
                # Show brief feedback
                entry.delete(0, tk.END)
                entry.insert(0, ot_value)
        except Exception as e:
            pass  # Silently handle clipboard errors


    def on_closing(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_closing_impl
        return on_closing_impl(self, *args, **kwargs)
    
    def export_stock_to_excel(self):
        """Export current stock data to Excel"""
        try:
            # Get current filtered data from treeview
            stock_data = []
            
            cols = list(self.stock_tree['columns'])
            for item in self.stock_tree.get_children():
                values = self.stock_tree.item(item, 'values')
                row_dict = {}
                for idx, col in enumerate(cols):
                    if idx < len(values):
                        heading_text = self.stock_tree.heading(col, 'text') or col
                        row_dict[heading_text] = values[idx]
                stock_data.append(row_dict)
            
            if not stock_data:
                messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
                return
            
            # Prepare filename with current date
            current_date = datetime.datetime.now().strftime('%Y%m%d')
            filename = f"재고현황_{current_date}.xlsx"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=filename,
                title="재고 현황 저장",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if save_path:
                stock_df = pd.DataFrame(stock_data)
                stock_df = self.clean_df_export(stock_df)
                self.save_df_to_excel_autofit(stock_df, save_path, "재고현황")
                messagebox.showinfo("완료", f"재고 현황이 저장되었습니다.\n저장 위치: {save_path}")
                
        except Exception as e:
            messagebox.showerror("오류", f"내보내기 실패: {e}")

    def enable_tree_column_drag(self, tree, context_menu_handler=None):
        """Ultra-robust drag-and-drop column reordering using root-level event tracking"""
        tree._drag_info = {"col_id": None, "col_name": None, "start_x": 0, "active": False, "ghost": None}
        
        def on_press(event):
            region = tree.identify_region(event.x, event.y)
            if region == "heading":
                col_id = tree.identify_column(event.x)
                col_name = self._get_column_name_from_id(tree, col_id)
                if col_name:
                    tree._drag_info.update({
                        "col_id": col_id,
                        "col_name": col_name,
                        "start_x": event.x_root,
                        "active": False,
                        "motion_id": self.root.bind("<B1-Motion>", on_motion, add="+"),
                        "release_id": self.root.bind("<ButtonRelease-1>", on_release, add="+")
                    })
                
        def on_motion(event):
            di = tree._drag_info
            if not di["col_name"]: return
            
            # Start drag if moved > 5 pixels
            if not di["active"] and abs(event.x_root - di["start_x"]) > 5:
                di["active"] = True
                tree.config(cursor="fleur")
                
                # Visual ghost label
                try:
                    if di["ghost"]: di["ghost"].destroy()
                    di["ghost"] = tk.Toplevel(self.root)
                    di["ghost"].overrideredirect(True)
                    di["ghost"].attributes("-alpha", 0.8)
                    di["ghost"].attributes("-topmost", True)
                    
                    lbl = tk.Label(di["ghost"], text=f" ↔ {di['col_name']} 이동 중 ", 
                                   bg="#424242", fg="white", relief="raised", borderwidth=1, 
                                   padx=8, pady=4, font=('Malgun Gothic', 10, 'bold'))
                    lbl.pack()
                except: pass
            
            if di["active"] and di["ghost"]:
                try:
                    # Follow mouse with offset to stay visible
                    di["ghost"].geometry(f"+{event.x_root + 20}+{event.y_root + 10}")
                except: pass

        def on_release(event):
            di = tree._drag_info
            if di["active"] and di["col_name"]:
                # Convert screen X back to tree-relative X
                tree_x = event.x_root - tree.winfo_rootx()
                target_id = tree.identify_column(tree_x)
                
                if target_id and target_id != di["col_id"]:
                    self._reorder_tree_columns(tree, di["col_id"], target_id)
            
            # Cleanup all bindings and state
            tree.config(cursor="")
            if di["ghost"]:
                try: di["ghost"].destroy()
                except: pass
                di["ghost"] = None
            
            # Unbind specific root-level drag listeners securely
            try:
                if di.get("motion_id"):
                    self.root.unbind("<B1-Motion>", di["motion_id"])
                if di.get("release_id"):
                    self.root.unbind("<ButtonRelease-1>", di["release_id"])
            except: pass
            
            di.update({"col_id": None, "col_name": None, "active": False, "motion_id": None, "release_id": None})

        # Primary entry point: tree-level press
        tree.bind("<Button-1>", on_press, add="+")
        # Direct right-click menu as reliable alternative
        if context_menu_handler is not None:
            tree.bind("<Button-3>", context_menu_handler, add="+")
        else:
            tree.bind("<Button-3>", lambda e: self._show_heading_context_menu(e, tree), add="+")

    def _show_heading_context_menu(self, event, tree):
        """Show context menu on Treeview header to move columns"""
        region = tree.identify_region(event.x, event.y)
        if region == "heading":
            column_id = tree.identify_column(event.x)
            
            # Identify the column name from visual ID
            col_name = self._get_column_name_from_id(tree, column_id)
            if not col_name: return

            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label=f"⬅️ '{col_name}' 왼쪽으로 이동", 
                             command=lambda: self._move_column_visual(tree, column_id, -1))
            menu.add_command(label=f"➡️ '{col_name}' 오른쪽으로 이동", 
                             command=lambda: self._move_column_visual(tree, column_id, 1))
            menu.add_separator()
            menu.add_command(label="⚙️ 컬럼 관리(숨기기/보이기)...", command=self.show_column_visibility_dialog)
            
            menu.post(event.x_root, event.y_root)

    def _show_generic_tree_heading_context_menu(self, event, tree):
        """Generic header context menu for trees that support only column move."""
        try:
            if tree.identify_region(event.x, event.y) != "heading":
                return
            column_id = tree.identify_column(event.x)
            col_name = self._get_column_name_from_id(tree, column_id)
            if not col_name:
                return
        except Exception:
            return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"⬅️ '{col_name}' 왼쪽으로 이동",
                         command=lambda: self._move_column_visual(tree, column_id, -1))
        menu.add_command(label=f"➡️ '{col_name}' 오른쪽으로 이동",
                         command=lambda: self._move_column_visual(tree, column_id, 1))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _get_column_name_from_id(self, tree, column_id):
        """Helper to get actual column name from visual ID like #1"""
        try:
            if not column_id or not column_id.startswith("#"): return None
            # Extract N from #N
            vis_idx = int(column_id[1:]) - 1
            
            # Get displaycolumns accurately
            disp_tuple = tree.cget("displaycolumns")
            if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
                visible_names = list(tree.cget("columns"))
            else:
                visible_names = list(disp_tuple)
            
            if 0 <= vis_idx < len(visible_names):
                return visible_names[vis_idx]
        except: pass
        return None

    def _move_column_visual(self, tree, source_id, direction):
        """Move a column left (-1) or right (1) based on its current position"""
        col_name = self._get_column_name_from_id(tree, source_id)
        if not col_name: return
        
        try:
            disp_tuple = tree.cget("displaycolumns")
            if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
                visible_names = list(tree.cget("columns"))
            else:
                visible_names = list(disp_tuple)

            if col_name not in visible_names: return
            
            src_idx = visible_names.index(col_name)
            tgt_idx = src_idx + direction
            
            if 0 <= tgt_idx < len(visible_names):
                visible_names.pop(src_idx)
                visible_names.insert(tgt_idx, col_name)
                tree["displaycolumns"] = visible_names
                
                # Update persistence
                if tree == self.daily_usage_tree:
                    self.manual_visible_cols = visible_names
                    self.save_tab_config()
                elif tree == self.monthly_usage_tree:
                    self.monthly_visible_cols = visible_names
                    self.save_tab_config()
                elif hasattr(self, 'budget_view_tree') and tree == self.budget_view_tree:
                    self.budget_view_visible_cols = visible_names
                    self.save_tab_config()
        except: pass

    def _reorder_tree_columns(self, tree, source_id, target_id):
        """Internal helper to swap displaycolumns for drag-and-drop"""
        src_name = self._get_column_name_from_id(tree, source_id)
        tgt_name = self._get_column_name_from_id(tree, target_id)
        
        if not src_name or not tgt_name or src_name == tgt_name: return
        
        try:
            disp_tuple = tree.cget("displaycolumns")
            if not disp_tuple or disp_tuple == ("#all",) or disp_tuple == ["#all"] or disp_tuple == "":
                visible_names = list(tree.cget("columns"))
            else:
                visible_names = list(disp_tuple)
            
            if src_name in visible_names and tgt_name in visible_names:
                src_idx = visible_names.index(src_name)
                tgt_idx = visible_names.index(tgt_name)
                
                # Strategic insertion: move before or after target 
                # (Simple swap approach: just pop and insert)
                visible_names.pop(src_idx)
                visible_names.insert(tgt_idx, src_name)
                
                tree["displaycolumns"] = visible_names
                
                # Update persistence
                if tree == self.daily_usage_tree:
                    self.manual_visible_cols = visible_names
                    self.save_tab_config()
                elif tree == self.monthly_usage_tree:
                    self.monthly_visible_cols = visible_names
                    self.save_tab_config()
                elif hasattr(self, 'budget_view_tree') and tree == self.budget_view_tree:
                    self.budget_view_visible_cols = visible_names
                    self.save_tab_config()
        except: pass


    def treeview_sort_column(self, tv, col, reverse):
        """Sort treeview contents when a column header is clicked"""
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Separate the 'Total' row (tagged 'total') from sorting
        data_rows = []
        total_row = None
        
        for val, k in l:
            # Check tags for this item
            tags = tv.item(k, 'tags')
            # If tags is a tuple/list, check if 'total' is in it. 
            # If it's a string (though usually tuple), check equality or contains.
            if tags and ('total' in tags or tags == 'total'):
                total_row = (val, k)
            else:
                data_rows.append((val, k))
                
        # Helper for numeric conversion
        def convert(val):
            try:
                # Remove common non-numeric chars
                s = str(val).replace(',', '').replace('시간', '').replace('Hrs', '').replace('원', '').replace(' ', '').replace('(', '').replace(')', '')
                if not s: return 0.0
                return float(s)
            except ValueError:
                return str(val).lower() # Default to string sort

        try:
            data_rows.sort(key=lambda t: convert(t[0]), reverse=reverse)
        except Exception:
            data_rows.sort(key=lambda t: t[0].lower(), reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(data_rows):
            tv.move(k, '', index)
            
        # Ensure Total row is always at the bottom
        if total_row:
             tv.move(total_row[1], '', 'end')

        # Reverse sort next time
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def switch_to_daily_usage_details(self):
        """Switch to Daily Usage tab and apply current filters from Sales tab"""
        # Sync filters before switching

        site = self.cb_sales_filter_site.get().strip()
        start = self.sales_start_date.get_date()
        end = self.sales_end_date.get_date()
        
        # 1. Switch Tab (index 5)
        self.notebook.select(5)
        
        # [NEW] Force focus transition
        if hasattr(self, 'daily_usage_tree'):
            self.daily_usage_tree.focus_set()

        # 2. Sync Filters
        if hasattr(self, 'cb_daily_filter_site'):
            self.cb_daily_filter_site.set(site)
        
        if hasattr(self, 'ent_daily_start_date'):
            self.ent_daily_start_date.set_date(start)
            
        if hasattr(self, 'ent_daily_end_date'):
            self.ent_daily_end_date.set_date(end)
            
        # 3. Trigger Search
        self.update_daily_usage_view()

    def save_all_daily_usage_changes(self):
        """Unified save button for the inquiry list to satisfy spreadsheet-style workflow"""
        try:
            if self.save_data():
                # Provide clear feedback that ALL changes in the app are saved
                messagebox.showinfo("저장 완료", "모든 변경 사항이 마스터 파일(Material_Inventory.xlsx)에 안전하게 저장되었습니다.")
                # Refresh view to ensure everything is synchronized
                self.update_daily_usage_view()
                self.refresh_inquiry_filters()
        except Exception as e:
            messagebox.showerror("저장 오류", f"데이터 저장 중 오류가 발생했습니다: {e}")

    def open_detached_monthly_usage_view(self):
        """월별 집계 조회를 별도의 팝업창(모니터링 창)으로 엽니다. (메인 화면과 100% 동일한 구성)"""
        # 기존 창이 있으면 포커스만 이동
        if 'monthly' in self.detached_windows and self.detached_windows['monthly']['window'].winfo_exists():
            self.detached_windows['monthly']['window'].lift()
            return

        popup = tk.Toplevel(self.root)
        popup.title("📊 월별 집계 모니터링 (팝업)")
        popup.geometry("1600x900")
        
        main_frame = ttk.Frame(popup, padding=10)
        main_frame.pack(expand=True, fill='both')
        
        # 상단 정보/필터 안내
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill='x', pady=(0, 5))
        
        # [NEW] Popup Search Filters
        filter_row = ttk.Frame(info_frame)
        filter_row.pack(side='left', fill='x', expand=True)

        def sync_and_refresh(event=None):
            # Mirror popup filters to main window
            self.cb_filter_year.set(p_cb_year.get())
            self.cb_filter_month.set(p_cb_month.get())
            self.cb_filter_site_monthly.set(p_cb_site.get())
            if hasattr(self, 'cb_filter_method_monthly'):
                self.cb_filter_method_monthly.set(p_cb_method.get())
            self.cb_filter_material_monthly.set(p_cb_mat.get())
            self.update_monthly_usage_view()

        ttk.Label(filter_row, text="🗓️ 연도:").pack(side='left', padx=2)
        p_cb_year = ttk.Combobox(filter_row, values=self.cb_filter_year['values'], width=8)
        p_cb_year.pack(side='left', padx=2)
        p_cb_year.set(self.cb_filter_year.get())
        p_cb_year.bind("<<ComboboxSelected>>", sync_and_refresh)

        ttk.Label(filter_row, text="📅 월:").pack(side='left', padx=2)
        p_cb_month = ttk.Combobox(filter_row, values=self.cb_filter_month['values'], width=6)
        p_cb_month.pack(side='left', padx=2)
        p_cb_month.set(self.cb_filter_month.get())
        p_cb_month.bind("<<ComboboxSelected>>", sync_and_refresh)

        ttk.Label(filter_row, text="🏢 현장:").pack(side='left', padx=2)
        p_cb_site = ttk.Combobox(filter_row, values=self.cb_filter_site_monthly['values'], width=12)
        p_cb_site.pack(side='left', padx=2)
        p_cb_site.set(self.cb_filter_site_monthly.get())
        p_cb_site.bind("<<ComboboxSelected>>", sync_and_refresh)

        ttk.Label(filter_row, text="🔬 방법:").pack(side='left', padx=2)
        method_vals = getattr(self, 'cb_filter_method_monthly', {}).get('values', ['전체', 'RT', 'PAUT', 'MT', 'PT']) if hasattr(self, 'cb_filter_method_monthly') else ['전체', 'RT', 'PAUT', 'MT', 'PT']
        p_cb_method = ttk.Combobox(filter_row, values=method_vals, width=8)
        p_cb_method.pack(side='left', padx=2)
        p_cb_method.set(self.cb_filter_method_monthly.get() if hasattr(self, 'cb_filter_method_monthly') else '전체')
        p_cb_method.bind("<<ComboboxSelected>>", sync_and_refresh)

        ttk.Label(filter_row, text="📦 품목:").pack(side='left', padx=2)
        p_cb_mat = ttk.Combobox(filter_row, values=self.cb_filter_material_monthly['values'], width=18)
        p_cb_mat.pack(side='left', padx=2)
        p_cb_mat.set(self.cb_filter_material_monthly.get())
        p_cb_mat.bind("<<ComboboxSelected>>", sync_and_refresh)

        ttk.Button(filter_row, text="🔍 조회", command=sync_and_refresh, style="Primary.TButton").pack(side='left', padx=5)
        
        ttk.Button(filter_row, text="🔄 새로고침", command=sync_and_refresh, style="Secondary.TButton").pack(side='right', padx=5)
        
        paned = ttk.PanedWindow(main_frame, orient="vertical")
        paned.pack(expand=True, fill='both')
        
        # 1. 메인 트리뷰 (상단)
        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3) 
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        columns = ('연도', '월', '현장', '구분', '작업자', '작업시간', 'OT시간', 'OT금액', 'OT1', 'OT2', 'OT3', 'OT4', 'OT5', 'OT6', 'OT7', 'OT8', 'OT9', 'OT10', 
                   '수량', '단가', '출장비', '일식', '검사비', '품목명', '센터미스', '농도', '마킹미스', '필름마크', 
                   '취급부주의', '고객불만', '기타', 'RTK총계', '형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제', '비고', '(Full작업자)')
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree['displaycolumns'] = [c for c in columns if c != '(Full작업자)']
        
        vsb.config(command=tree.yview); hsb.config(command=tree.xview)
        
        col_widths = {'연도': 90, '월': 70, '현장': 140, '구분': 100, '작업자': 100, '작업시간': 100, 'OT시간': 100, 'OT금액': 110, '품목명': 200, '비고': 220}
        for col in columns:
            tree.heading(col, text=("자재명" if col == "품목명" else col), command=lambda c=col: self.treeview_sort_column(tree, c, False))
            tree.column(col, width=col_widths.get(col, 100), anchor='center', stretch=False)
        
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)

        # [NEW] 인터랙티브 기능 복제
        self.enable_tree_column_drag(tree)
        tree.tag_configure('total', background='#E8F4F8', font=('Arial', 12, 'bold'))
        tree.bind("<Button-1>", lambda e: self.show_worker_popup(e, tree), add="+")
        tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())

        # 2. 현장별 요약 (중단 - 메인 UI는 3단 수직 분할)
        site_frame = ttk.LabelFrame(paned, text="🏢 현장별 누계")
        paned.add(site_frame, weight=1)
        site_cols = ('현장', '검사방법', '품목명', '수량', '검사비', '출장비', '제경비', '기술료', '환산물량', '재료비', '인건비', '형광자분', '흑색자분', '백색페인트', 
                     '침투제', '세척제', '현상제', '형광침투제', '센터미스', '농도', '마킹미스', '필름마크', '취급부주의', '고객불만', '기타', 'RTK총계')
        site_tree = ttk.Treeview(site_frame, columns=site_cols, show='headings')
        site_vsb = ttk.Scrollbar(site_frame, orient="vertical", command=site_tree.yview)
        site_tree.configure(yscrollcommand=site_vsb.set)
        for col in site_cols:
            site_tree.heading(col, text=("자재명" if col == "품목명" else col), command=lambda c=col: self.treeview_sort_column(site_tree, c, False))
            w = 120 if col in ['현장', '검사방법', '품목명'] else (100 if col in ['검사비', '출장비', '제경비', '기술료', '재료비', '인건비'] else 80)
            site_tree.column(col, width=w, anchor='center', stretch=False)
        site_tree.pack(side='left', expand=True, fill='both')
        site_vsb.pack(side='right', fill='y')
        
        self.enable_tree_column_drag(site_tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, site_tree))
        site_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())
        
        # 3. 작업자별 요약 (하단)
        worker_frame = ttk.LabelFrame(paned, text="👷 작업자별 누계")
        paned.add(worker_frame, weight=1)
        worker_cols = ('작업자', '총공수', '연장(시간)', '야간(시간)', '휴일(시간)', '총OT(시간)', '연장(금액)', '야간(금액)', '휴일(금액)', '총OT(금액)')
        worker_tree = ttk.Treeview(worker_frame, columns=worker_cols, show='headings')
        worker_vsb = ttk.Scrollbar(worker_frame, orient="vertical", command=worker_tree.yview)
        worker_tree.configure(yscrollcommand=worker_vsb.set)
        worker_widths = {'작업자': 100, '총공수': 70, '총OT(시간)': 80, '총OT(금액)': 100}
        for col in worker_cols:
            worker_tree.heading(col, text=("자재명" if col == "품목명" else col), command=lambda c=col: self.treeview_sort_column(worker_tree, c, False))
            worker_tree.column(col, width=worker_widths.get(col, 80), anchor='center', stretch=False)
        worker_tree.pack(side='left', expand=True, fill='both')
        worker_vsb.pack(side='right', fill='y')

        self.enable_tree_column_drag(worker_tree, context_menu_handler=lambda e: self._show_generic_tree_heading_context_menu(e, worker_tree))
        worker_tree.bind('<ButtonRelease-1>', lambda e: self.save_tab_config())

        # [NEW] 메인 UI의 분할 배율 동기화
        def sync_sash():
            try:
                main_h = self.monthly_paned.winfo_height()
                if main_h > 0:
                    popup_h = paned.winfo_height()
                    if popup_h > 0:
                        s1 = self.monthly_paned.sashpos(0)
                        s2 = self.monthly_paned.sashpos(1)
                        paned.sashpos(0, int(s1 * popup_h / main_h))
                        paned.sashpos(1, int(s2 * popup_h / main_h))
            except: pass
        popup.after(500, sync_sash) # UI 렌더링 후 적용

        # 상태 등록
        self.detached_windows['monthly'] = {
            'window': popup,
            'tree': tree,
            'site_tree': site_tree,
            'worker_tree': worker_tree,
            'filters': {
                'year': p_cb_year,
                'month': p_cb_month,
                'site': p_cb_site,
                'method': p_cb_method,
                'mat': p_cb_mat
            }
        }
        
        def on_close():
            if 'monthly' in self.detached_windows:
                del self.detached_windows['monthly']
            popup.destroy()
        popup.protocol("WM_DELETE_WINDOW", on_close)
        
        # 데이터 복사
        self.update_monthly_usage_view()
        
        # [FIXED] Extraction using standardized indices (Year=0, Month=1, Site=2)
        def on_popup_select(event):
            selection = tree.selection()
            if not selection: return
            item = selection[0]
            try:
                vals = tree.item(item, 'values')
                if not vals: return
                tags = tree.item(item, 'tags')
                
                if not hasattr(self, 'current_monthly_df'): return
                
                if 'total' in tags:
                    subset = self.current_monthly_df
                else:
                    # Robust extraction
                    try:
                        y = int(vals[0]) 
                        m = int(vals[1])
                        s = str(vals[2]).strip()
                        
                        # Material name is at index 22 in our formatted tuple
                        mat_name = str(vals[22]).strip() if len(vals) > 22 else None
                        
                        mask = (self.current_monthly_df['Year'] == y) & \
                               (self.current_monthly_df['Month'] == m) & \
                               (self.current_monthly_df['Site'] == s)
                        
                        if mat_name:
                            ids = self.materials_df[self.materials_df['품목명'] == mat_name]['MaterialID'].tolist()
                            if ids:
                                mask = mask & (self.current_monthly_df['MaterialID'].isin(ids))
                        
                        subset = self.current_monthly_df[mask]
                    except:
                        subset = pd.DataFrame()
                
                # Update summaries
                self._populate_monthly_summary_trees(subset)
            except Exception as e:
                print(f"DEBUG: Detached selection error: {e}")

        tree.bind("<<TreeviewSelect>>", on_popup_select)
        
        if hasattr(self, 'current_monthly_df'):
            self._populate_monthly_summary_trees(self.current_monthly_df)

    def open_detached_budget_view(self):
        """공사실행예산서 탭을 별도의 팝업창으로 엽니다. (메인 화면과 유사한 KPI/목록 구성)"""
        if 'budget' in self.detached_windows and self.detached_windows['budget']['window'].winfo_exists():
            self.detached_windows['budget']['window'].lift()
            return

        popup = tk.Toplevel(self.root)
        popup.title("📋 공사 실행예산서 모니터링 (팝업)")
        popup.geometry("1400x800")
        
        main_frame = ttk.Frame(popup, padding=10)
        main_frame.pack(expand=True, fill='both')
        
        # 1. KPI Panel (메인 UI 그대로 재현)
        kpi_frame = tk.Frame(main_frame, bg='#f8fafc', height=80)
        kpi_frame.pack(fill='x', pady=(0, 10))
        kpi_frame.pack_propagate(False)
        
        def create_kpi_card(parent, title, color):
            frame = tk.Frame(parent, bg='white', highlightbackground='#e2e8f0', highlightthickness=1)
            frame.pack(side='left', expand=True, fill='both', padx=5, pady=5)
            tk.Label(frame, text=title, font=('Malgun Gothic', 9), fg='#64748b', bg='white').pack(pady=(5, 2))
            lbl = tk.Label(frame, text="0", font=('Malgun Gothic', 16, 'bold'), fg=color, bg='white')
            lbl.pack(pady=(0, 5))
            return lbl

        lbl_rev = create_kpi_card(kpi_frame, "총 매출액", "#2563eb")
        lbl_cost = create_kpi_card(kpi_frame, "총 실행원가", "#dc2626")
        lbl_profit = create_kpi_card(kpi_frame, "총 기대이익", "#059669")
        lbl_margin = create_kpi_card(kpi_frame, "평균 이익률", "#7c3aed")
        
        # 2. 예산 목록 트리뷰
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(expand=True, fill='both')
        
        vsb = ttk.Scrollbar(tree_frame, orient='vertical')
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal')
        
        columns = ('Site', 'Revenue', 'UnitPrice', 'LaborCost', 'MaterialCost', 'Expense', 'OutsourceCost', 'Profit', 'TargetMargin', 'ActualMargin', 'Note')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        col_map = {
            'Site': ('현장명', 150), 'Revenue': ('매출액(목표)', 120), 'UnitPrice': ('단가', 100),
            'LaborCost': ('노무비', 120), 'MaterialCost': ('자재비', 120), 'Expense': ('경비', 120),
            'OutsourceCost': ('외주비', 120), 'Profit': ('기대이익(목표)', 120), 'TargetMargin': ('목표이익률', 100), 'ActualMargin': ('현재이익률', 100), 'Note': ('비고', 200)
        }
        for col in columns:
            head, width = col_map[col]
            tree.heading(col, text=head, command=lambda c=col: self.treeview_sort_column(tree, c, False))
            tree.column(col, width=width, anchor='center', stretch=False)
            
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)
        
        vsb.config(command=tree.yview); hsb.config(command=tree.xview)
        
        # [NEW] 인터랙티브 기능 추가
        self.enable_tree_column_drag(tree)

        self.detached_windows['budget'] = {
            'window': popup,
            'tree': tree,
            'lbl_rev': lbl_rev,
            'lbl_cost': lbl_cost,
            'lbl_profit': lbl_profit,
            'lbl_margin': lbl_margin
        }
        
        def on_close():
            if 'budget' in self.detached_windows:
                del self.detached_windows['budget']
            popup.destroy()
        popup.protocol("WM_DELETE_WINDOW", on_close)
        
        self._refresh_detached_budget()

    def _refresh_detached_budget(self):
        """팝업창의 예산 리스트와 KPI를 최신 데이터로 갱신"""
        if 'budget' not in self.detached_windows: return
        detached = self.detached_windows['budget']
        tree = detached['tree']
        
        for item in tree.get_children(): tree.delete(item)
        
        if self.budget_df.empty:
            for lbl in [detached['lbl_rev'], detached['lbl_cost'], detached['lbl_profit']]: lbl.config(text="0")
            detached['lbl_margin'].config(text="0.0%")
            return
        
        total_rev = total_cost = total_profit = 0
        
        for _, row in self.budget_df.iterrows():
            rev = pd.to_numeric(row.get('Revenue', 0), errors='coerce') or 0
            unit_p = pd.to_numeric(row.get('UnitPrice', 0), errors='coerce') or 0
            lab = pd.to_numeric(row.get('LaborCost', 0), errors='coerce') or 0
            mat = pd.to_numeric(row.get('MaterialCost', 0), errors='coerce') or 0
            exp = pd.to_numeric(row.get('Expense', 0), errors='coerce') or 0
            out = pd.to_numeric(row.get('OutsourceCost', 0), errors='coerce') or 0
            prof = pd.to_numeric(row.get('Profit', 0), errors='coerce') or 0
            
            a_rev = pd.to_numeric(row.get('Actual_Revenue', 0), errors='coerce') or 0
            a_prof = pd.to_numeric(row.get('Actual_Profit', 0), errors='coerce') or 0
            
            total_rev += rev
            total_cost += (lab + mat + exp + out)
            total_profit += prof
            
            target_margin = (prof / rev * 100) if rev > 0 else 0
            actual_margin = (a_prof / a_rev * 100) if a_rev > 0 else 0
            
            tree.insert('', 'end', values=(
                row.get('Site', ''),
                f"{rev:,.0f}",
                f"{unit_p:,.0f}",
                f"{lab:,.0f}",
                f"{mat:,.0f}",
                f"{exp:,.0f}",
                f"{out:,.0f}",
                f"{prof:,.0f}",
                f"{target_margin:.1f}%",
                f"{actual_margin:.1f}%",
                row.get('Note', '')
            ))
            
        # KPI Update
        detached['lbl_rev'].config(text=f"{total_rev:,.0f}")
        detached['lbl_cost'].config(text=f"{total_cost:,.0f}")
        detached['lbl_profit'].config(text=f"{total_profit:,.0f}")
        total_margin = (total_profit / total_rev * 100) if total_rev > 0 else 0
        detached['lbl_margin'].config(text=f"{total_margin:.1f}%")


    def cleanup_auto_transactions(self):
        """Remove existing 'Automatic Deduction' transactions for non-consumable equipment (PAUT, Yokes, etc.)"""
        if self.transactions_df.empty: 
            messagebox.showinfo("알림", "처리할 입출고 기록이 없습니다.")
            return
        
        confirm = messagebox.askyesno("클린업 확인", 
            "현재 저장된 '자동 차감' 기록 중 장비류(PAUT, Yoke 등)에 해당하는 내역만 골라 삭제하시겠습니까?\n\n"
            "※ 필름, 약품 등 소모품 기록은 유지됩니다.")
        if not confirm: return

        initial_count = len(self.transactions_df)
        
        def is_bad_auto(row):
            note = str(row.get('Note', ''))
            # [FIX] Match both English and potentially broken encoding versions of "(자동 차감)"
            # (ڵ ) is a common corruption of (자동 차감) when read as different encoding
            is_auto = "(자동 차감)" in note or "(ڵ )" in note or "(\ucac0 \u0020)" in note or "(\u00e1\u00b6\u00bd\u0020)" in note

            if not is_auto:
                return False
            
            mat_id = row.get('MaterialID')
            mat_name = self.get_material_display_name(mat_id)
            
            # If it's NOT a consumable, it's a candidate for removal
            # _is_consumable_material handles empty method by returning False for equipment keywords
            return not self._is_consumable_material(mat_name, "")

        mask = self.transactions_df.apply(is_bad_auto, axis=1)
        self.transactions_df = self.transactions_df[~mask]
        
        removed_count = initial_count - len(self.transactions_df)
        if removed_count > 0:
            if self.save_data():
                messagebox.showinfo("클린업 완료", f"장비류(PAUT, Yoke 등)의 자동 차감 내역 {removed_count}건을 삭제했습니다.")
                self.update_stock_view()
                self.update_transaction_view()
        else:
            messagebox.showinfo("알림", "정리할 대상이 없습니다.")

    def open_usage_memo_history(self):
        """현장별 일일 사용량 기입 탭의 상시 메모 이력 조회 및 이동"""
        memo_rows = []
        full_memos = {}
        if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
            if 'DailyMemo' in self.daily_usage_df.columns:
                import pandas as pd
                df_memos = self.daily_usage_df.dropna(subset=['DailyMemo'])
                site = self.cb_daily_site.get().strip()
                if site and site != '전체' and 'Site' in df_memos.columns:
                    df_memos = df_memos[df_memos['Site'].fillna('').astype(str).str.strip() == site]
                df_memos = df_memos[df_memos['DailyMemo'].astype(str).str.strip() != '']
                
                from site_apps.central.src.views.memo_details import daily_memo_lookup
                for (date_str, _site), memo in daily_memo_lookup(df_memos).items():
                    if date_str in full_memos:
                        full_memos[date_str] += '\n\n' + memo
                    else:
                        full_memos[date_str] = memo
                for date_str, memo in full_memos.items():
                    preview = ' '.join(memo.split())
                    memo_rows.append((date_str, preview))

        memo_rows.sort(key=lambda item: item[0], reverse=True)

        import tkinter as tk
        from tkinter import ttk
        window = tk.Toplevel(self.root)
        window.title("현장 메모 이력")
        window.geometry("760x480")
        window.minsize(560, 320)
        window.transient(self.root)

        search_frame = ttk.Frame(window, padding=10)
        search_frame.pack(fill="x")
        ttk.Label(search_frame, text="날짜·내용 검색:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        table_frame = ttk.Frame(window, padding=(10, 0, 10, 10))
        table_frame.pack(fill="both", expand=True)
        columns = ('date', 'memo')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        tree.heading('date', text='날짜')
        tree.heading('memo', text='메모 내용')
        tree.column('date', width=110, minwidth=100, stretch=False, anchor='center')
        tree.column('memo', width=580, minwidth=300, stretch=True, anchor='w')
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        status_var = tk.StringVar()
        ttk.Label(window, textvariable=status_var, padding=(10, 0, 10, 8)).pack(fill='x')

        def refresh_rows(*_args):
            query = search_var.get().strip().lower()
            tree.delete(*tree.get_children())
            matched = [row for row in memo_rows if not query or query in row[0].lower() or query in row[1].lower()]
            for date_key, preview in matched:
                tree.insert('', 'end', values=(date_key, preview))
            status_var.set(f"메모 {len(matched)}건  ·  더블클릭하면 해당 날짜로 이동합니다.")

        def move_to_selected(_event=None):
            selected = tree.selection()
            if not selected:
                return
            date_key = str(tree.item(selected[0], 'values')[0])
            try:
                import datetime
                from site_apps.central.src.views.daily_date_context import open_memo_date
                selected_date = datetime.datetime.strptime(date_key, '%Y-%m-%d').date()
                if open_memo_date(self, selected_date, full_memos[date_key]):
                    window.destroy()
            except Exception as e:
                print(f"Date navigate error: {e}")

        tree.bind('<Double-1>', move_to_selected)
        tree.bind('<Return>', move_to_selected)
        search_var.trace_add('write', refresh_rows)
        refresh_rows()
        search_entry.focus_set()

if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialManager(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass
