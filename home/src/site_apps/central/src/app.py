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
from site_apps.central.src.runtime_config import ensure_initial_config
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
import site_apps.central.src.daily_work_report_manager as daily_work_report_manager
print(f"DEBUG: daily_work_report_manager path: {daily_work_report_manager.__file__}")

# Pre-compiled regex for performance
from site_apps.central.src.views.components import *

# Auto-install additional dependencies
from tkcalendar import DateEntry, Calendar

# [FIX] Patch tkcalendar for stability in ko_KR locale / Python 3.14.
# 각 패치를 독립적으로 적용 - 한 패치 실패가 나머지를 무효화하지 않도록 분리.
import tkinter as _tk_ref

# 1) Calendar.__init__ : select_on_nav=False 기본값 + 헤더 너비 조정 + 초기화 중 전역 update_idletasks 병목 차단
try:
    _orig_cal_init = Calendar.__init__
    def _patched_cal_init(self, *args, **kwargs):
        if 'select_on_nav' not in kwargs:
            kwargs['select_on_nav'] = False
        _orig_misc_uid = _tk_ref.Misc.update_idletasks
        _tk_ref.Misc.update_idletasks = lambda _s: None
        try:
            _orig_cal_init(self, *args, **kwargs)
        finally:
            _tk_ref.Misc.update_idletasks = _orig_misc_uid
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

        # User settings are intentionally ignored by Git, so a clean checkout
        # needs a valid first-run file before configuration is loaded.
        ensure_initial_config(self.config_path)

        # 중앙지사에서 선택한 클라우드 재고 파일을 우선 사용한다.
        from site_apps.central.src.data_sync import get_business_data_dir, get_external_data_root
        cloud_inventory = os.path.join(get_business_data_dir(), 'Material_Inventory.xlsx')
        if get_external_data_root() and not os.path.isfile(cloud_inventory):
            raise FileNotFoundError(f'중앙지사 운영 재고 파일이 없습니다: {cloud_inventory}')
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
                    self.tab_config = config
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
        from site_apps.central.src.services.data_loader import load_data_impl
        result = load_data_impl(self, *args, **kwargs)
        self.refresh_operating_data_status()
        return result


    
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
        except (ValueError, TypeError):
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
            
        old_materials = self.materials_df.copy(deep=True)
        self.materials_df = pd.concat([self.materials_df, pd.DataFrame([new_mat])], ignore_index=True)
        if persist:
            if not self.save_data(require_primary=True):
                self.materials_df = old_materials
                raise RuntimeError(f"신규 자재('{name}')를 원본 재고 파일에 저장하지 못했습니다.")
            self.update_material_combo()
        return new_id
    
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
        from site_apps.central.src.services.data_loader import save_data_impl
        result = save_data_impl(self, *args, **kwargs)
        self._inventory_save_status = (
            f"완료 {time.strftime('%Y-%m-%d %H:%M:%S')}" if result else "실패 (재고 원본 저장 미완료)"
        )
        self.refresh_operating_data_status()
        return result

    def refresh_operating_data_status(self):
        if not hasattr(self, "_inventory_path_status"):
            return
        path = os.path.abspath(self.db_path)
        if getattr(self, "_inventory_load_failed", False):
            path += " (불러오기 실패)"
        elif not os.path.isfile(path):
            path += " (파일 없음)"
        self._inventory_path_status.set(path)
        sync = getattr(self, "_last_sync_success_at", "") or "기록 없음"
        save = getattr(self, "_inventory_save_status", "이번 실행에서 저장 기록 없음")
        self._operating_data_status.set(f"마지막 동기화 성공  {sync}    |    재고 저장  {save}")
            
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
        status = ttk.Frame(self.root, padding=(10, 3))
        status.pack(side='bottom', fill='x')
        self._inventory_path_status = tk.StringVar(master=status)
        self._operating_data_status = tk.StringVar(master=status)
        path_row = ttk.Frame(status)
        path_row.pack(fill='x')
        ttk.Label(path_row, text="재고 경로").pack(side='left', padx=(0, 8))
        ttk.Entry(path_row, textvariable=self._inventory_path_status, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Label(status, textvariable=self._operating_data_status).pack(anchor='w', pady=(3, 0))
        self.refresh_operating_data_status()

        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Tab 1: Current Stock (현재 재고 현황)
        self.tab_stock = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stock, text='현재 재고 현황')
        self._load_tab_on_first_map(self.tab_stock, self._initialize_stock_tab)

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
        work_log_page = ttk.Frame(self.notebook)
        self.notebook.add(work_log_page, text='작업/감독일보')
        self._load_tab_on_first_map(work_log_page, lambda: self.setup_daily_work_log_tab(work_log_page), show_loading=True)

        # Tab 8: NDT Billing Calculator (기성 정산 (NDT))
        self.tab_ndt_billing = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_ndt_billing, text='기성 정산 (NDT)')
        self._load_tab_on_first_map(self.tab_ndt_billing, self.setup_ndt_billing_tab, show_loading=True)

        # Tab 9: Project Execution Budget (공사실행예산서)
        self.tab_budget = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_budget, text='공사실행예산서')
        self.setup_budget_tab()

        # Tab 10: NDT Inspection Report Generator (비파괴검사 성적서/보고서)
        try:
            from site_apps.central.src.views.ndt_report_tab import NDTReportTab
            self.tab_ndt_report = NDTReportTab(self.notebook, main_app=self, lazy=True)
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
                else:
                    saved_tab_name = None
            if not saved_tab_name:
                saved_index = self.tab_config.get('selected_tab')
                tabs = self.notebook.tabs()
                if isinstance(saved_index, int) and 0 <= saved_index < len(tabs):
                    init_tab = tabs[saved_index]
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

    def _load_tab_on_first_map(self, page, setup, *, show_loading=False):
        def build():
            try:
                setup()
            except Exception as error:
                ttk.Label(page, text=f"탭을 불러오지 못했습니다.\n{error}", foreground="red").pack(
                    anchor="w", padx=10, pady=10
                )

        def initialize(event):
            if event.widget is not page:
                return
            page.unbind("<Map>", binding)
            if not show_loading:
                build()
                return
            loading = ttk.Label(page, text="화면을 불러오는 중…")
            loading.pack(anchor="w", padx=10, pady=10)

            def build_after_display():
                try:
                    build()
                finally:
                    loading.destroy()

            page.after(50, build_after_display)
        binding = page.bind("<Map>", initialize, add="+")

    def setup_daily_work_log_tab(self, page):
        from site_apps.central.src.daily_work_log_tab import DailyWorkLogTab

        self.tab_daily_work_log = DailyWorkLogTab(page)
        self.tab_daily_work_log.main_app = self
        self.tab_daily_work_log.pack(fill="both", expand=True)

    def setup_stock_tab(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import setup_stock_tab_impl
        return setup_stock_tab_impl(self, *args, **kwargs)

    def _initialize_stock_tab(self):
        if hasattr(self, 'stock_tree'):
            return
        self.setup_stock_tab()
        widths = getattr(self, 'tab_config', {}).get('stock_col_widths', {})
        for column in self.stock_tree['columns']:
            try:
                width = int(widths.get(column, 0))
            except (TypeError, ValueError):
                continue
            if width > 0:
                self.stock_tree.column(column, width=width, minwidth=20, stretch=False)
        self.update_stock_view()
    
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

    def delete_selected_material(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import delete_selected_material_impl
        return delete_selected_material_impl(self, *args, **kwargs)

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

    def save_material_edits(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import save_material_edits_impl
        return save_material_edits_impl(self, *args, **kwargs)

    
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
        if hasattr(self, '_initialize_stock_tab'):
            self._initialize_stock_tab()
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
        
        
    def format_worker_summary(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import format_worker_summary_impl
        return format_worker_summary_impl(self, *args, **kwargs)

    def clean_df_export(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import clean_df_export_impl
        return clean_df_export_impl(self, *args, **kwargs)

    def save_df_to_excel_autofit(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import save_df_to_excel_autofit_impl
        return save_df_to_excel_autofit_impl(self, *args, **kwargs)


    
    def setup_import_tab(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import setup_import_tab_impl
        return setup_import_tab_impl(self, *args, **kwargs)
    
    def import_from_excel(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import import_from_excel_impl
        return import_from_excel_impl(self, *args, **kwargs)
    
    def export_materials(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_materials_impl
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
    

    def view_monthly_usage(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import view_monthly_usage_impl
        return view_monthly_usage_impl(self, *args, **kwargs)

    def generate_yearly_report(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import generate_yearly_report_impl
        return generate_yearly_report_impl(self, *args, **kwargs)

    def generate_monthly_report(self, *args, **kwargs):
        from site_apps.central.src.views.stock_management_view import generate_monthly_report_impl
        return generate_monthly_report_impl(self, *args, **kwargs)

    def setup_monthly_usage_tab(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import setup_monthly_usage_tab_impl
        return setup_monthly_usage_tab_impl(self, *args, **kwargs)
    
    def update_monthly_usage_view(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import update_monthly_usage_view_impl
        return update_monthly_usage_view_impl(self, *args, **kwargs)

    def on_monthly_usage_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_monthly_usage_select_impl
        return on_monthly_usage_select_impl(self, *args, **kwargs)

    def _populate_monthly_summary_trees(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import _populate_monthly_summary_trees_impl
        return _populate_monthly_summary_trees_impl(self, *args, **kwargs)

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

    def export_monthly_usage_history(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import export_monthly_usage_history_impl
        return export_monthly_usage_history_impl(self, *args, **kwargs)

    def create_draggable_container(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import create_draggable_container_impl
        return create_draggable_container_impl(self, *args, **kwargs)

    def bulk_register_materials_to_stock(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import bulk_register_materials_to_stock_impl
        return bulk_register_materials_to_stock_impl(self, *args, **kwargs)

    def refresh_ui_for_list_change(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import refresh_ui_for_list_change_impl
        return refresh_ui_for_list_change_impl(self, *args, **kwargs)

    def remove_box(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import remove_box_impl
        return remove_box_impl(self, *args, **kwargs)

    def destroy_custom_widget(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import destroy_custom_widget_impl
        return destroy_custom_widget_impl(self, *args, **kwargs)

    def rename_widget_label(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import rename_widget_label_impl
        return rename_widget_label_impl(self, *args, **kwargs)

    def clone_widget(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import clone_widget_impl
        return clone_widget_impl(self, *args, **kwargs)

    def _bind_recursive(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _bind_recursive_impl
        return _bind_recursive_impl(self, *args, **kwargs)

    def hide_widget(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import hide_widget_impl
        return hide_widget_impl(self, *args, **kwargs)

    def make_draggable(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import make_draggable_impl
        return make_draggable_impl(self, *args, **kwargs)

    def make_header_draggable(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import make_header_draggable_impl
        return make_header_draggable_impl(self, *args, **kwargs)
        
    def reset_widget_position(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import reset_widget_position_impl
        return reset_widget_position_impl(self, *args, **kwargs)

    def reset_all_widgets_layout(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import reset_all_widgets_layout_impl
        return reset_all_widgets_layout_impl(self, *args, **kwargs)

    def save_current_layout_as_default(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import save_current_layout_as_default_impl
        return save_current_layout_as_default_impl(self, *args, **kwargs)
            
    def toggle_layout_lock(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import toggle_layout_lock_impl
        return toggle_layout_lock_impl(self, *args, **kwargs)

    def on_drag_stop(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_drag_stop_impl
        return on_drag_stop_impl(self, *args, **kwargs)


    def add_new_memo(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import add_new_memo_impl
        return add_new_memo_impl(self, *args, **kwargs)

    def duplicate_memo(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import duplicate_memo_impl
        return duplicate_memo_impl(self, *args, **kwargs)

    def add_new_checklist(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import add_new_checklist_impl
        return add_new_checklist_impl(self, *args, **kwargs)

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

    def add_vehicle_inspection_box(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import add_vehicle_inspection_box_impl
        return add_vehicle_inspection_box_impl(self, *args, **kwargs)

    def add_checklist_item(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import add_checklist_item_impl
        return add_checklist_item_impl(self, *args, **kwargs)

    def duplicate_checklist(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import duplicate_checklist_impl
        return duplicate_checklist_impl(self, *args, **kwargs)


    def on_drag_start(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_drag_start_impl
        return on_drag_start_impl(self, *args, **kwargs)
        
    def on_resize_start(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_resize_start_impl
        return on_resize_start_impl(self, *args, **kwargs)
        
    def on_mouse_motion(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import on_mouse_motion_impl
        return on_mouse_motion_impl(self, *args, **kwargs)

    def _update_widget_position(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _update_widget_position_impl
        return _update_widget_position_impl(self, *args, **kwargs)
        
    def _apply_push_down_logic(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _apply_push_down_logic_impl
        return _apply_push_down_logic_impl(self, *args, **kwargs)

    def _adjust_parent_height(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _adjust_parent_height_impl
        return _adjust_parent_height_impl(self, *args, **kwargs)


    def _ensure_placeholder(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _ensure_placeholder_impl
        return _ensure_placeholder_impl(self, *args, **kwargs)
                
    def _remove_placeholder(self, *args, **kwargs):
        from site_apps.central.src.views.draggable_layout_controller import _remove_placeholder_impl
        return _remove_placeholder_impl(self, *args, **kwargs)

        

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
        from site_apps.central.src.sync_google_sheet import inventory_deletion_key, queue_inventory_deletions
        cloud_deletes = []
        for idx in indices:
            deletion_key = inventory_deletion_key(old_daily.loc[idx])
            if deletion_key:
                cloud_deletes.append(deletion_key)
        try:
            # Auto-deduction transactions are already excluded by stock calculation.
            # Date/site matches cannot identify which daily job created a legacy transaction.
            with queue_inventory_deletions(getattr(self, 'db_path', None), cloud_deletes):
                self.daily_usage_df = self.daily_usage_df.drop(list(indices))
                if not self.save_data(require_primary=True):
                    raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.daily_usage_df = old_daily
            self.transactions_df = old_transactions
            messagebox.showerror('삭제 저장 실패', f'{error}\n원본 기록을 복원했습니다.', parent=self.root)
            return

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

    def setup_daily_usage_query_tab(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import setup_daily_usage_query_tab_impl
        return setup_daily_usage_query_tab_impl(self, *args, **kwargs)


    def _show_daily_usage_note(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _show_daily_usage_note_impl
        return _show_daily_usage_note_impl(self, *args, **kwargs)

    def _on_daily_usage_select(self, *args, **kwargs):
        from site_apps.central.src.controllers.event_controller import _on_daily_usage_select_impl
        return _on_daily_usage_select_impl(self, *args, **kwargs)


    

    def export_budget_estimation(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import export_budget_estimation_impl
        return export_budget_estimation_impl(self, *args, **kwargs)

    def export_budget_sales_status(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import export_budget_sales_status_impl
        return export_budget_sales_status_impl(self, *args, **kwargs)

    def setup_ndt_billing_tab(self):
        from site_apps.central.src.ndt_billing_tab import NDTCalculatorTab

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
        from site_apps.central.src.ndt_billing_tab import NDTCalculatorTab

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

    def _refresh_budget_view_tree_columns(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import _refresh_budget_view_tree_columns_impl
        return _refresh_budget_view_tree_columns_impl(self, *args, **kwargs)

    def add_budget_view_custom_column(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import add_budget_view_custom_column_impl
        return add_budget_view_custom_column_impl(self, *args, **kwargs)

    def rename_budget_view_column(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import rename_budget_view_column_impl
        return rename_budget_view_column_impl(self, *args, **kwargs)

    def set_budget_view_custom_column_content(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import set_budget_view_custom_column_content_impl
        return set_budget_view_custom_column_content_impl(self, *args, **kwargs)

    def delete_budget_view_custom_column(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import delete_budget_view_custom_column_impl
        return delete_budget_view_custom_column_impl(self, *args, **kwargs)

    def show_budget_view_column_dialog(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import show_budget_view_column_dialog_impl
        return show_budget_view_column_dialog_impl(self, *args, **kwargs)

    def _show_budget_view_heading_context_menu(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import _show_budget_view_heading_context_menu_impl
        return _show_budget_view_heading_context_menu_impl(self, *args, **kwargs)

    def update_budget_site_view(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import update_budget_site_view_impl
        return update_budget_site_view_impl(self, *args, **kwargs)


    def _update_budget_kpis(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import _update_budget_kpis_impl
        return _update_budget_kpis_impl(self, *args, **kwargs)

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
                except (ValueError, TypeError):
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
                
        # 4. 업체(Company) 데이터 수집 — 발주처('발주처')와 NDT 사용 자재 제조사('회사코드')를 엄격히 분리
        company_data_list = []
        client_code = self.cb_daily_client.get().strip() if hasattr(self, 'cb_daily_client') else '한국지역난방공사'
        if not client_code:
            client_code = '한국지역난방공사'
        if hasattr(self, 'ndt_company_entries'):
            for company_entries in self.ndt_company_entries:
                is_first = (company_entries == self.ndt_company_entries[0])
                # '회사코드'에는 오직 NDT 사용 자재 제조사만 기록하며 미입력 시 빈 문자열로 유지(발주처와 혼용 차단)
                raw_ndt_co = company_entries['_company'].get().strip()
                company_code = '' if raw_ndt_co in ('한국지역난방공사', '한국지역난방공사 중앙지사') else raw_ndt_co
                ndt_data = {name: to_f(company_entries.get(name)) for name in self.ndt_materials_all}
                
                # [REFINEMENT] Create record if there is NDT data OR it is the first entry (guaranteed row)
                if any(v > 0 for v in ndt_data.values()) or is_first:
                    company_data_list.append({
                        '회사코드': company_code,
                        'ndt_data': ndt_data,
                        'is_primary': is_first
                    })
        
        if not company_data_list: # [FALLBACK] Ensure at least one row exists
             company_data_list.append({'회사코드': '', 'ndt_data': {}, 'is_primary': True})

        # 5. 공통 레코드 데이터 구성 (기본값)
        from site_apps.central.src.aux_equipment import value as auxiliary_equipment_value
        common_data = {
            'Date': date_val,
            'Site': site,
            '발주처': client_code,
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

        if not hasattr(self, 'co_code_list') or not isinstance(self.co_code_list, list):
            self.co_code_list = []
        co_changed = False
        for cd in company_data_list:
            co_val = str(cd.get('회사코드', '') or '').strip()
            if co_val and co_val not in ('한국지역난방공사', '한국지역난방공사 중앙지사'):
                if not any(c.strip() == co_val for c in self.co_code_list):
                    self.co_code_list.append(co_val)
                    co_changed = True
        if co_changed:
            self.co_code_list.sort()
            self.refresh_ui_for_list_change('co_code')

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

    def add_ndt_company_section(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import add_ndt_company_section_impl
        return add_ndt_company_section_impl(self, *args, **kwargs)
    
    def remove_last_ndt_company(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import remove_last_ndt_company_impl
        return remove_last_ndt_company_impl(self, *args, **kwargs)

    def sync_worker_times(self, *args, **kwargs):
        from site_apps.central.src.models.worker_model import sync_worker_times_impl
        return sync_worker_times_impl(self, *args, **kwargs)
    def _load_ndt_product_map(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _load_ndt_product_map_impl
        return _load_ndt_product_map_impl(self, *args, **kwargs)

    def _save_ndt_product_map(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _save_ndt_product_map_impl
        return _save_ndt_product_map_impl(self, *args, **kwargs)

    def open_ndt_product_map_dialog(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import open_ndt_product_map_dialog_impl
        return open_ndt_product_map_dialog_impl(self, *args, **kwargs)

    def open_equipment_search_dialog(self):
        """Select multiple central equipment units without losing serial numbers."""
        from site_apps.central.src.aux_equipment import select_main
        return select_main(self.root, self, self.cb_daily_equip)

    def open_preferred_item_search_dialog(self):
        """Deprecated in favor of unified open_material_search_dialog"""
        self.open_material_search_dialog(target_form='registration')

    def _get_merged_memo_and_note(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _get_merged_memo_and_note_impl
        return _get_merged_memo_and_note_impl(self, *args, **kwargs)

    def load_daily_usage_to_form(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import load_daily_usage_to_form_impl
        return load_daily_usage_to_form_impl(self, *args, **kwargs)

    def clear_daily_usage_form_all(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import clear_daily_usage_form_all_impl
        return clear_daily_usage_form_all_impl(self, *args, **kwargs)

    def add_daily_usage_entry(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import add_daily_usage_entry_impl
        return add_daily_usage_entry_impl(self, *args, **kwargs)

    def _recover_combobox_interaction(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _recover_combobox_interaction_impl
        return _recover_combobox_interaction_impl(self, *args, **kwargs)

    def load_report_mapping(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import load_report_mapping_impl
        return load_report_mapping_impl(self, *args, **kwargs)

    def save_report_mapping(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import save_report_mapping_impl
        return save_report_mapping_impl(self, *args, **kwargs)

    def open_report_mapping_dialog(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import open_report_mapping_dialog_impl
        return open_report_mapping_dialog_impl(self, *args, **kwargs)

    def export_daily_work_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_daily_work_report_impl
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

    def _get_inout_session_data(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _get_inout_session_data_impl
        return _get_inout_session_data_impl(self, *args, **kwargs)

    def _set_inout_session_data(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _set_inout_session_data_impl
        return _set_inout_session_data_impl(self, *args, **kwargs)

    def _get_budget_session_data(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _get_budget_session_data_impl
        return _get_budget_session_data_impl(self, *args, **kwargs)

    def _set_budget_session_data(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _set_budget_session_data_impl
        return _set_budget_session_data_impl(self, *args, **kwargs)

    def save_form_session(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import save_form_session_impl
        return save_form_session_impl(self, *args, **kwargs)

    def load_previous_day_data(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import load_previous_day_data_impl
        return load_previous_day_data_impl(self, *args, **kwargs)

    def load_form_session(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import load_form_session_impl
        return load_form_session_impl(self, *args, **kwargs)

    def _auto_reconcile_and_register_ndt(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _auto_reconcile_and_register_ndt_impl
        return _auto_reconcile_and_register_ndt_impl(self, *args, **kwargs)

    def _create_manual_stock_transaction(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _create_manual_stock_transaction_impl
        return _create_manual_stock_transaction_impl(self, *args, **kwargs)
    
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
                monthly_mat_names = set()
                equip_tokens = ('SCANNER', '스캐너', 'YOKE', '장비', 'CRAWLER', 'PAUT')
                if not self.daily_usage_df.empty and 'MaterialID' in self.daily_usage_df.columns:
                    for mat_id in self.daily_usage_df['MaterialID'].dropna().unique():
                        m_only = self.get_material_name_only(mat_id)
                        if m_only and not m_only.startswith('(미등록)') and not any(tok in m_only.upper() for tok in equip_tokens):
                            monthly_mat_names.add(m_only)
                if not self.daily_usage_df.empty and '검사방법' in self.daily_usage_df.columns:
                    methods_present = set(self.daily_usage_df['검사방법'].dropna().astype(str).str.strip().str.upper())
                    if 'MT' in methods_present:
                        monthly_mat_names.add('MT약품')
                    if 'PT' in methods_present:
                        monthly_mat_names.add('PT약품')
                self.cb_filter_material_monthly['values'] = ['전체'] + sorted(monthly_mat_names)

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

    def _auto_adjust_tree_columns(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _auto_adjust_tree_columns_impl
        return _auto_adjust_tree_columns_impl(self, *args, **kwargs)

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
        from site_apps.central.src.sync_google_sheet import inventory_deletion_key, queue_inventory_deletions
        cloud_deletes = []
        for idx in indices:
            deletion_key = inventory_deletion_key(old_daily.loc[idx])
            if deletion_key:
                cloud_deletes.append(deletion_key)
        try:
            # Auto-deduction transactions are already excluded by stock calculation.
            # Date/site matches cannot identify which daily job created a legacy transaction.
            with queue_inventory_deletions(getattr(self, 'db_path', None), cloud_deletes):
                self.daily_usage_df = self.daily_usage_df.drop(list(indices))
                if not self.save_data(require_primary=True):
                    raise RuntimeError('원본 파일에 저장하지 못했습니다.')
        except Exception as error:
            self.daily_usage_df = old_daily
            self.transactions_df = old_transactions
            messagebox.showerror('삭제 저장 실패', f'{error}\n원본 기록을 복원했습니다.', parent=self.root)
            return

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
            except (ValueError, TypeError):
                v_to_set = 0.0

        try:
            self.daily_usage_df.at[idx, col] = v_to_set
        except (ValueError, TypeError):
            try:
                # Type conflict (e.g. float into strict string column or vice versa)
                # Cast the column to object to allow the mixed/new type
                self.daily_usage_df[col] = self.daily_usage_df[col].astype(object)
                self.daily_usage_df.at[idx, col] = v_to_set
            except (ValueError, TypeError):
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
                
                # Comprehensive list of columns that MUST be numeric (OT~OT10은 원화 문자열/공백 호환을 위해 제외)
                numeric_cols = ['Usage', '수량', '검사량', '단가', '출장비', '일식', '검사비', 'FilmCount']
                is_numeric = k.startswith('NDT_') or k.startswith('RTK_') or k in numeric_cols
                
                # Use safe setter to avoid all dtype crashes
                self._safe_set_daily_df(df_idx, target_key, v, is_numeric=(is_numeric or k == 'MaterialID'))

            from site_apps.central.src.models.stock_calculation import synchronize_legacy_chemicals
            for column, value in synchronize_legacy_chemicals(old_entry, new_data).items():
                self._safe_set_daily_df(df_idx, column, value, is_numeric=False)

            # Ensure Usage, 검사량, 수량 stay synchronized on edits (RT REP rows keep Usage/수량 = 0.0)
            edited_amount = float(new_data.get('검사량', 0.0) or 0.0)
            is_rt_rep = (
                str(new_data.get('검사방법', old_entry.get('검사방법', ''))).strip().upper() == 'RT'
                and str(old_entry.get('검사구분', '')).strip().upper() == 'REP'
            )
            effective_usage = 0.0 if is_rt_rep else edited_amount
            usage_col = 'Usage' if 'Usage' in self.daily_usage_df.columns else '수량'
            self._safe_set_daily_df(df_idx, usage_col, effective_usage, is_numeric=True)
            if '검사량' in self.daily_usage_df.columns:
                self._safe_set_daily_df(df_idx, '검사량', effective_usage, is_numeric=True)
            if '수량' in self.daily_usage_df.columns and usage_col != '수량':
                self._safe_set_daily_df(df_idx, '수량', effective_usage, is_numeric=True)
            
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
                
    def export_daily_usage_history(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import export_daily_usage_history_impl
        return export_daily_usage_history_impl(self, *args, **kwargs)

    def show_column_visibility_dialog(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import show_column_visibility_dialog_impl
        return show_column_visibility_dialog_impl(self, *args, **kwargs)

    def export_weekly_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_weekly_report_impl
        return export_weekly_report_impl(self, *args, **kwargs)

    def export_monthly_ndt_report(self, *args, **kwargs):
        from site_apps.central.src.services.excel_exporter import export_monthly_ndt_report_impl
        return export_monthly_ndt_report_impl(self, *args, **kwargs)

    def export_all_daily_usage(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import export_all_daily_usage_impl
        return export_all_daily_usage_impl(self, *args, **kwargs)

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
    
    def export_stock_to_excel(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import export_stock_to_excel_impl
        return export_stock_to_excel_impl(self, *args, **kwargs)

    def enable_tree_column_drag(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import enable_tree_column_drag_impl
        return enable_tree_column_drag_impl(self, *args, **kwargs)

    def _show_heading_context_menu(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _show_heading_context_menu_impl
        return _show_heading_context_menu_impl(self, *args, **kwargs)

    def _show_generic_tree_heading_context_menu(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _show_generic_tree_heading_context_menu_impl
        return _show_generic_tree_heading_context_menu_impl(self, *args, **kwargs)

    def _get_column_name_from_id(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _get_column_name_from_id_impl
        return _get_column_name_from_id_impl(self, *args, **kwargs)

    def _move_column_visual(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _move_column_visual_impl
        return _move_column_visual_impl(self, *args, **kwargs)

    def _reorder_tree_columns(self, *args, **kwargs):
        from site_apps.central.src.views.daily_usage_form_controller import _reorder_tree_columns_impl
        return _reorder_tree_columns_impl(self, *args, **kwargs)


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

    def open_detached_monthly_usage_view(self, *args, **kwargs):
        from site_apps.central.src.views.monthly_usage_view import open_detached_monthly_usage_view_impl
        return open_detached_monthly_usage_view_impl(self, *args, **kwargs)

    def open_detached_budget_view(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import open_detached_budget_view_impl
        return open_detached_budget_view_impl(self, *args, **kwargs)

    def _refresh_detached_budget(self, *args, **kwargs):
        from site_apps.central.src.views.budget_view import _refresh_detached_budget_impl
        return _refresh_detached_budget_impl(self, *args, **kwargs)


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
