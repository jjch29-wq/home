import re
NAN_PATTERN = re.compile(r'^nan(\.0+)?$|^none$|^null$|^0\.0+|-0\.0+$', re.IGNORECASE)
DOT_ZERO_PATTERN = re.compile(r'\.0$')
MARKER_PATTERN = re.compile(r'\(.*?\)\s*|익일')

import sys
import subprocess
import os
import pandas as pd
import tkinter as tk

def get_appdata_dir(app_name):
    """
    Returns a safe, writable directory in the user's AppData for storing config/data.
    Ensures the directory exists.
    """
    appdata = os.getenv('APPDATA')
    if not appdata:
        appdata = os.path.expanduser('~')
    
    target_dir = os.path.join(appdata, 'PMI_Apps', app_name)
    os.makedirs(target_dir, exist_ok=True)
    return target_dir

def install_and_import(package, import_name=None):
    if import_name is None: import_name = package
    try:
        return __import__(import_name)
    except ImportError:
        try:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--break-system-packages"])
            return __import__(import_name)
        except Exception:
            # Fallback for uv-managed systems
            try:
                subprocess.check_call(["uv", "pip", "install", "--system", "--break-system-packages", package])
                return __import__(import_name)
            except:
                pass
            sys.exit(1)

def normalize_id(val):
    """Robustly normalize IDs: handle NaN, trailing .0, and whitespace."""
    if pd.isna(val) or val == '' or str(val).lower() == 'nan': return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s


def enable_column_resize(frame, num_cols, header_row=0, edge_px=6):
    """
    frame 내부 grid 테이블의 헤더 셀(header_row) 오른쪽 경계를 드래그해
    해당 컬럼과 오른쪽 인접 컬럼의 너비를 조절하는 기능을 추가합니다.

    - frame  : ttk.Frame, grid 레이아웃으로 컬럼이 구성된 프레임
    - num_cols : 컬럼 수
    - header_row : 헤더가 위치한 row 번호 (기본 0)
    - edge_px : 경계 감지 영역 (픽셀)
    """
    _drag = {'col': None, 'start_x': 0, 'start_w': 0, 'next_w': 0}

    def get_col_width(col):
        """현재 컬럼의 실제 픽셀 너비를 반환"""
        try:
            info = frame.grid_columnconfigure(col)
            minsz = info.get('minsize', 0)
            if minsz and minsz > 0:
                return minsz
            for widget in frame.grid_slaves(row=header_row, column=col):
                w = widget.winfo_width()
                if w > 1:
                    return w
            return 80
        except:
            return 80

    def on_motion(event):
        """드래그 중이 아닐 때만 커서 모양 변경"""
        if _drag['col'] is not None:
            return  # 이미 드래그 중이면 무시
        widget = event.widget
        x = event.x
        w = widget.winfo_width()
        if w - edge_px <= x <= w:
            widget.configure(cursor='sb_h_double_arrow')
        else:
            widget.configure(cursor='')

    def on_leave(event):
        if _drag['col'] is None:
            event.widget.configure(cursor='')

    def on_press(event):
        widget = event.widget
        x = event.x
        w = widget.winfo_width()
        if w - edge_px <= x <= w:
            info = widget.grid_info()
            col = info.get('column', -1)
            if col < 0 or col >= num_cols - 1:
                return
            _drag['col'] = col
            _drag['widget'] = widget
            _drag['start_x'] = event.x_root
            _drag['start_w'] = get_col_width(col)
            _drag['next_w'] = get_col_width(col + 1)
            # grab_set 제거 - 스크롤 방해 방지
        else:
            _drag['col'] = None

    def on_drag(event):
        if _drag['col'] is None:
            return
        col = _drag['col']
        dx = event.x_root - _drag['start_x']
        new_w = max(30, _drag['start_w'] + dx)
        new_next = max(30, _drag['next_w'] - dx)
        frame.grid_columnconfigure(col, minsize=new_w, weight=0)
        frame.grid_columnconfigure(col + 1, minsize=new_next, weight=0)

    def on_release(event):
        _drag['col'] = None
        # 커서 원상복구
        try:
            w = event.widget
            w.configure(cursor='')
        except:
            pass

    # 헤더 row(0) 의 모든 위젯에 바인딩
    def bind_headers():
        for col in range(num_cols):
            for widget in frame.grid_slaves(row=header_row, column=col):
                widget.bind('<Motion>', on_motion, add='+')
                widget.bind('<Leave>', on_leave, add='+')
                widget.bind('<ButtonPress-1>', on_press, add='+')
                widget.bind('<B1-Motion>', on_drag, add='+')
                widget.bind('<ButtonRelease-1>', on_release, add='+')

    # 위젯이 아직 렌더되기 전일 수 있으므로 idle 후 바인딩
    frame.after_idle(bind_headers)


# --- Extracted Helpers ---
def clean_nan_impl(self, val):
    """Robustly clean NaN, None, and other empty markers from any value"""
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if not s or NAN_PATTERN.match(s):
        return ""
    # Handle trailing .0 from numeric inputs converted to string
    s = DOT_ZERO_PATTERN.sub('', s)
    if NAN_PATTERN.match(s): return ""
    return s


def format_entry_with_commas_impl(self, event, entry):
    """Autoformat numbers with commas as the user types or leaves the field"""
    try:
        val = entry.get().strip().replace(',', '')
        if not val: return
        
        # Use float or int depending on content
        if '.' in val:
            num = float(val)
            formatted = f"{num:,.1f}"
            if formatted.endswith('.0'): formatted = formatted[:-2]
        else:
            num = int(val)
            formatted = f"{num:,}"
            
        entry.delete(0, tk.END)
        entry.insert(0, formatted)
    except: pass


def normalize_site_name_impl(self, name):
    """[ROBUST] Normalize site names: strip whitespace, unify ALL hyphen types, collapse multiple spaces."""
    if name is None or pd.isna(name): return ""
    import re
    s = str(name).strip()
    # Unify various types of hyphens/dashes: -, –, —, －
    s = re.sub(r'[\u002D\u2013\u2014\uFF0D]', '-', s)
    # Unify hyphens by removing spaces around them: "Site - A" -> "Site-A"
    s = re.sub(r'\s*-\s*', '-', s)
    # Collapse any multiple internal spaces
    s = " ".join(s.split())
    return s


