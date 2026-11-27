import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from site_apps.central.src.services.monthly_report_manager import MonthlyReportManager
from tkcalendar import DateEntry
from datetime import datetime
import os
from queue import Empty, Queue
from threading import Thread
from site_apps.central.src.data_sync import (
    SharedDataUnavailableError,
    get_data_connection_error,
    get_history_path,
    get_process_photos_dir,
)
import sys
import json
import shutil
import uuid
from copy import deepcopy
from site_apps.central.src.rt_quantities import rt_counts, sync_rt_widgets, widget_rt_values, resolve_rt_row, retest_parent, default_rt_copy, corrected_rt_history
from site_apps.central.src.services.section_line_helper import (
    extract_company_section_line_mapping,
    get_sections_for_company,
    get_lines_for_company_and_section,
)

# Import the exporter
try:
    from site_apps.central.src.daily_work_log_exporter import DailyWorkLogExporter
except ImportError:
    pass # Will be handled if run standalone

SIZE_LENGTH = {
    '1100A': 3.511,  '1000A': 3.1919, '900A': 2.8727, '850A': 2.7131,
    '800A':  2.5535, '750A':  2.3939, '700A': 2.2343, '650A': 2.0747,
    '600A':  1.9151, '550A':  1.7555, '500A': 1.5959, '450A': 1.4363,
    '400A':  1.2767, '350A':  1.1172, '300A': 1.0006, '250A': 0.8401,
    '200A':  0.6795, '150A':  0.519,  '125A': 0.4392, '100A': 0.3591,
    '80A':   0.2799, '65A':   0.2397, '50A':  0.1901, '40A':  0.1527,
    '32A':   0.1341, '25A':   0.1068, '20A':  0.0855,
}

# 용접사 관리대장에서 갱신된 최신 발급번호.
# 과거 작업일보 및 공정사진에 저장된 번호도 같은 기준으로 자동 이관한다.
LEGACY_WELDER_IDS = {
    'W-2023-A-12': 'W-2026-A-08',
    'W-2023-A-13': 'W-2026-A-05',
}

COMPANY_WORK_CATEGORIES = {
    '다림': ['', '제조', '수탁', '건가'],
    '다림건설': ['', '제조', '수탁', '건가'],
    '다림건설(주)': ['', '제조', '수탁', '건가'],
    'GS네오텍': ['', '건가-추가', '건가-개선'],
    'GS': ['', '건가-추가', '건가-개선'],
    '지에스네오텍': ['', '건가-추가', '건가-개선'],
}
ALL_WORK_CATEGORIES = ['', '제조', '수탁', '건가', '건가-추가', '건가-개선']


class _NDTCellProxy:
    """Treeview 셀 하나를 기존 Entry/Combobox 위젯 인터페이스와 호환되게 감싸는 경량 프록시."""

    __slots__ = ('_row', '_col', '_value', '_values_list', '_bindings', '_is_combo', '_readonly')

    def __init__(self, row_proxy, col_name, values=None, readonly=False):
        self._row = row_proxy
        self._col = col_name
        self._value = ''
        self._is_combo = values is not None
        self._values_list = list(values) if values is not None else []
        self._readonly = bool(readonly)
        self._bindings = {}

    def get(self):
        return self._value

    def set(self, val):
        new_val = '' if val is None else str(val)
        if self._value != new_val:
            self._value = new_val
            self._row._on_cell_value_updated(self._col)
        else:
            self._row._on_cell_value_updated(self._col)

    def delete(self, first, last=None):
        if first == 0 or first == '0':
            if last in (None, tk.END, 'end') or (isinstance(last, int) and last >= len(self._value)):
                if self._value != '':
                    self._value = ''
                    self._row._on_cell_value_updated(self._col)
                return
            if isinstance(last, int):
                self._value = self._value[last:]
                self._row._on_cell_value_updated(self._col)

    def insert(self, index, val):
        s = '' if val is None else str(val)
        if index in (tk.END, 'end'):
            self._value = self._value + s
        else:
            try:
                pos = int(index)
            except (TypeError, ValueError):
                pos = 0
            self._value = self._value[:pos] + s + self._value[pos:]
        self._row._on_cell_value_updated(self._col)

    def __getitem__(self, key):
        if key == 'values':
            return tuple(self._values_list)
        if key == 'state':
            return 'readonly' if self._readonly else 'normal'
        raise KeyError(key)

    def __setitem__(self, key, value):
        if key == 'values':
            self._values_list = list(value or [])
            self._row._tab._sync_active_combo_values(self._row._row_index, self._col)
            return
        if key == 'state':
            self._readonly = (value == 'readonly')
            return
        raise KeyError(key)

    def configure(self, **kwargs):
        if 'values' in kwargs:
            self['values'] = kwargs['values']
        if 'state' in kwargs:
            self['state'] = kwargs['state']

    config = configure

    def cget(self, key):
        return self[key]

    def bind(self, sequence, func, add=None):
        seq = '<Button-1>' if sequence == '<ButtonPress-1>' else sequence
        if add in ('+', True) and seq in self._bindings:
            self._bindings[seq].append(func)
        else:
            self._bindings[seq] = [func]

    def event_generate(self, sequence, **kwargs):
        seq = '<Button-1>' if sequence == '<ButtonPress-1>' else sequence
        evt = type('_VirtualEvent', (), {'widget': self, **kwargs})()
        for fn in list(self._bindings.get(seq, ())):
            res = fn(evt)
            if res == 'break':
                break

    def focus_set(self):
        self._row._tab._focus_ndt_cell(self._row._row_index, self._col)


class _NDTSectionInfoProxy:
    """'구간정보' 셀(RT 4분할 / PAUT·MT·PT 단일 문자열) 호환 프록시."""

    def __init__(self, row_proxy):
        self._row = row_proxy
        self._slots = ['', '', '', '']
        self._rt_retake_meta = {}
        self._in_sync = False
        self._bindings = {}

    def _method(self):
        m_cell = self._row.get('검사방법')
        return m_cell.get().strip().upper() if m_cell is not None else ''

    def get(self):
        method = self._method()
        if method in ('PAUT', 'PT', 'MT'):
            return self._slots[0]
        parts = [s.strip() for s in self._slots]
        return ','.join(parts).rstrip(',') if method == 'RT' else ','.join([p for p in parts if p])

    def set(self, val):
        s_val = '' if val is None else str(val)
        self._slots = ['', '', '', '']
        method = self._method()
        if method in ('PAUT', 'PT', 'MT'):
            self._slots[0] = s_val
        else:
            parts = s_val.split(',') if s_val else []
            for i, p in enumerate(parts[:4]):
                self._slots[i] = p
        self._row._on_cell_value_updated('구간정보')
        if not self._in_sync and not self._row._batch_depth:
            self._in_sync = True
            try:
                sync_rt_widgets(self._row, all_rows=self._row._tab._rt_context())
            except ValueError:
                pass
            finally:
                self._in_sync = False

    def set_slots_from_editor(self, slots):
        for i in range(4):
            self._slots[i] = str(slots[i]) if i < len(slots) else ''
        self._row._on_cell_value_updated('구간정보')

    def get_slots(self):
        return list(self._slots)

    def delete(self, first, last=None):
        if first == 0 or first == '0':
            self._slots = ['', '', '', '']
            self._row._on_cell_value_updated('구간정보')

    def insert(self, index, val):
        s_val = '' if val is None else str(val)
        if index in (tk.END, 'end'):
            self._slots[0] = self._slots[0] + s_val
        else:
            self._slots[0] = s_val + self._slots[0]
        self._row._on_cell_value_updated('구간정보')

    def apply_layout(self):
        self._row._on_cell_value_updated('구간정보')
        self._row._tab._refresh_active_section_info_mode(self._row._row_index)

    def bind(self, sequence, func, add=None):
        seq = '<Button-1>' if sequence == '<ButtonPress-1>' else sequence
        if add in ('+', True) and seq in self._bindings:
            self._bindings[seq].append(func)
        else:
            self._bindings[seq] = [func]

    def event_generate(self, sequence, **kwargs):
        seq = '<Button-1>' if sequence == '<ButtonPress-1>' else sequence
        evt = type('_VirtualEvent', (), {'widget': self, **kwargs})()
        for fn in list(self._bindings.get(seq, ())):
            res = fn(evt)
            if res == 'break':
                break

    def focus_set(self):
        self._row._tab._focus_ndt_cell(self._row._row_index, '구간정보')


class _NDTRowProxy(dict):
    """한 행의 컬럼별 셀 프록시를 담고 Treeview 행 표시를 동기화하는 dict 서브클래스."""

    def __init__(self, tab, row_index, iid):
        super().__init__()
        self._tab = tab
        self._row_index = row_index  # 0-based
        self._iid = iid
        self._batch_depth = 0
        self._dirty = False

    def begin_batch(self):
        self._batch_depth += 1

    def end_batch(self):
        if self._batch_depth > 0:
            self._batch_depth -= 1
        if self._batch_depth == 0 and self._dirty:
            self._dirty = False
            self._flush_to_tree()
            state = getattr(self._tab, '_active_editor_state', None)
            if state and state.get('row') == self._row_index and state.get('col') in self:
                self._tab._sync_active_editor_if_matching(self._row_index, state['col'])

    def _on_cell_value_updated(self, col_name):
        if self._batch_depth > 0:
            self._dirty = True
            return
        tree = getattr(self._tab, 'ndt_tree', None)
        if tree is not None and tree.winfo_exists():
            try:
                tree.set(self._iid, col_name, self[col_name].get())
            except tk.TclError:
                pass
        self._tab._sync_active_editor_if_matching(self._row_index, col_name)

    def _flush_to_tree(self):
        tree = getattr(self._tab, 'ndt_tree', None)
        if tree is None or not tree.winfo_exists():
            return
        checked = False
        if self._row_index < len(self._tab.ndt_selection_vars):
            checked = bool(self._tab.ndt_selection_vars[self._row_index].get())
        mark = f"☑ {self._row_index + 1}" if checked else f"☐ {self._row_index + 1}"
        vals = [mark] + [self[c].get() for c in self._tab.ndt_cols]
        try:
            tree.item(self._iid, values=vals)
        except tk.TclError:
            pass


class DailyWorkLogTab(ttk.Frame):
    def __init__(self, parent, *args, main_app=None, initial_ndt_rows=10, target_ndt_rows=10, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.main_app = main_app
        self._initial_ndt_rows = max(1, int(initial_ndt_rows))
        self._target_ndt_rows = max(self._initial_ndt_rows, int(target_ndt_rows))
        self._ndt_warmup_job = None
        self.history_path = get_history_path()
        self._history_cache = None
        self._history_cache_mtime = None
        self._migrate_legacy_welder_ids()
        self.photo_root = get_process_photos_dir()
        self.data_connection_error = get_data_connection_error()
        self.selected_ndt_row = None
        self.welder_names = dict(MonthlyReportManager.WELDER_NAMES)
        self.welder_ids_by_name = {
            name: welder_id for welder_id, name in self.welder_names.items()
        }
        self.setup_ui()
        self._loaded_log_date = self._selected_date_key()
        self._log_baseline = self._capture_log_input()
        if self.data_connection_error:
            self.after_idle(
                lambda: messagebox.showerror(
                    "공유 작업자료 연결 오류",
                    self.data_connection_error,
                    parent=self.winfo_toplevel(),
                )
            )

    def _normalize_welder_id(self, value):
        """화면 표시값 또는 이름을 최신 용접사 번호 형식으로 변환한다."""
        value = str(value or '').strip()
        if not value:
            return ''
        if '|' in value:
            display_name, welder_id = value.rsplit('|', 1)
            welder_id = welder_id.strip()
            if welder_id:
                return LEGACY_WELDER_IDS.get(welder_id, welder_id)
            value = display_name.strip()
        welder_id = self.welder_ids_by_name.get(value, value)
        return LEGACY_WELDER_IDS.get(welder_id, welder_id)

    def _migrate_legacy_welder_ids(self):
        """기존 작업일보 JSON의 과거 용접사 번호를 최신 번호로 저장한다."""
        if not os.path.exists(self.history_path):
            return
        try:
            source_mtime = os.path.getmtime(self.history_path)
            with open(self.history_path, 'r', encoding='utf-8') as history_file:
                history = json.load(history_file)
            if not isinstance(history, dict) or any(not isinstance(day, dict) for day in history.values()):
                return  # The normal loader reports invalid history to the user.

            changed_count = 0

            def replace_legacy_ids(value):
                nonlocal changed_count
                if isinstance(value, dict):
                    return {key: replace_legacy_ids(item) for key, item in value.items()}
                if isinstance(value, list):
                    return [replace_legacy_ids(item) for item in value]
                if isinstance(value, str) and value.strip() in LEGACY_WELDER_IDS:
                    changed_count += 1
                    return LEGACY_WELDER_IDS[value.strip()]
                return value

            migrated_history = replace_legacy_ids(history)
            if os.path.getmtime(self.history_path) != source_mtime:
                return

            if changed_count:
                temp_path = self.history_path + '.tmp'
                with open(temp_path, 'w', encoding='utf-8') as history_file:
                    json.dump(migrated_history, history_file, ensure_ascii=False, indent=4)
                os.replace(temp_path, self.history_path)
                source_mtime = os.path.getmtime(self.history_path)
                history = migrated_history
                print(f"[DailyWorkLog] 용접사 번호 {changed_count}건 최신화 완료")
            self._rt_parent_history = [
                row for day in sorted(history) for row in history[day].get('ndt_results', [])
            ]
            self._history_cache = history
            self._history_cache_mtime = source_mtime
        except Exception as exc:
            print(f"[DailyWorkLog] 용접사 번호 최신화 실패: {exc}")

    def _format_welder_display(self, value):
        """저장된 용접사 번호를 '이름 | 번호' 화면 표시 형식으로 변환한다."""
        welder_id = self._normalize_welder_id(value)
        name = self.welder_names.get(welder_id)
        return f"{name} | {welder_id}" if name else welder_id
        
    def setup_ui(self):
        self._left_pane_visible = True
        self._left_pane_width = None
        self._saved_sash_ratio = 0.5
        self._last_paned_width = 0
        self._sash_press_pos = None
        self._resizing_ndt_col = False

        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=5, pady=(4, 0))
        self.left_pane_toggle = ttk.Button(
            toolbar, text='◀ 좌측 숨기기', command=self._toggle_left_pane,
            style="Secondary.TButton"
        )
        self.left_pane_toggle.pack(side='left')

        self._build_basic_info(self)

        # Create PanedWindow for Left/Right split
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill="both", expand=True)
        
        # --- LEFT PANE ---
        self.left_container = ttk.Frame(self.paned)
        self.paned.add(self.left_container, weight=4) # 40% width
        
        self.left_canvas = tk.Canvas(self.left_container)
        self.left_scrollbar = ttk.Scrollbar(self.left_container, orient="vertical", command=self.left_canvas.yview)
        self.left_frame = ttk.Frame(self.left_canvas)
        
        _update_left_scroll = self._bind_scroll_region(self.left_canvas)
        self.left_frame.bind("<Configure>", _update_left_scroll)
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
        
        self._last_left_win_size = (None, None)
        def _on_left_canvas_configure(event):
            new_width = max(event.width, self.left_frame.winfo_reqwidth())
            new_height = max(event.height, self.left_frame.winfo_reqheight())
            if (new_width, new_height) != self._last_left_win_size:
                self._last_left_win_size = (new_width, new_height)
                self.left_canvas.itemconfig(self.left_window, width=new_width, height=new_height)
                _update_left_scroll()
            
        self.left_canvas.bind("<Configure>", _on_left_canvas_configure)
        
        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_scrollbar.pack(side="right", fill="y")
        
        self._build_left_pane(self.left_frame)
        
        # --- RIGHT PANE (Treeview + In-Place Cell Editor) ---
        self.right_container = ttk.Frame(self.paned)
        self.paned.add(self.right_container, weight=6) # 60% width
        self.right_canvas = None
        self.right_frame = self.right_container
        self._build_right_pane(self.right_container)

        self.paned.bind("<Configure>", self._on_paned_configure, add="+")
        self.paned.bind("<ButtonPress-1>", self._on_paned_press, add="+")
        self.paned.bind("<ButtonRelease-1>", self._on_paned_release, add="+")

        # 저장된 좌우 프레임 비율·좌측 숨김 상태·결과서 열 너비를 복원한다.
        self._load_work_log_layout_prefs()

        initial_jobs = {
            self.after(200, self._apply_initial_pane_layout),
            self.after(100, self.on_date_change),
        }

        def cancel_initial_jobs(event):
            if event.widget is self:
                for job in initial_jobs:
                    self.after_cancel(job)
                initial_jobs.clear()

        self.bind("<Destroy>", cancel_initial_jobs, add="+")

    @staticmethod
    def _bind_scroll_region(canvas):
        pending = None
        previous_region = None

        def refresh():
            nonlocal pending, previous_region
            pending = None
            if not canvas.winfo_exists():
                return
            bbox = canvas.bbox("all")
            if bbox:
                region = (0, 0, bbox[2], max(bbox[3], canvas.winfo_height()))
                if region != previous_region:
                    canvas.configure(scrollregion=region)
                    previous_region = region

        def schedule(_event=None):
            nonlocal pending
            if pending is None:
                pending = canvas.after_idle(refresh)

        def cancel(_event):
            nonlocal pending
            if pending is not None:
                canvas.after_cancel(pending)
                pending = None

        canvas.bind("<Destroy>", cancel, add="+")
        return schedule

    def _get_host_tab_config(self):
        app = getattr(self, 'main_app', None)
        cfg = getattr(app, 'tab_config', None) if app is not None else None
        return cfg if isinstance(cfg, dict) else None

    def _load_work_log_layout_prefs(self):
        cfg = self._get_host_tab_config()
        if not cfg:
            return
        try:
            ratio = float(cfg.get('work_log_sash_ratio', 0.5))
            if 0.08 <= ratio <= 0.92:
                self._saved_sash_ratio = ratio
        except (TypeError, ValueError):
            pass

        saved_pos = cfg.get('work_log_sash_pos')
        if isinstance(saved_pos, (int, float)) and saved_pos > 50:
            self._left_pane_width = int(saved_pos)

        col_widths = cfg.get('work_log_ndt_col_widths', {})
        if isinstance(col_widths, dict) and hasattr(self, 'ndt_tree') and self.ndt_tree.winfo_exists():
            valid_cols = set(self.ndt_tree['columns'])
            for col, width in col_widths.items():
                if col in valid_cols:
                    try:
                        w = int(width)
                        if w >= 36:
                            self.ndt_tree.column(col, width=w, minwidth=36, stretch=False)
                    except (TypeError, ValueError, tk.TclError):
                        pass

        want_visible = bool(cfg.get('work_log_left_visible', True))
        if not want_visible and self._left_pane_visible:
            try:
                self.paned.forget(self.left_container)
                self._left_pane_visible = False
                self.left_pane_toggle.configure(text='▶ 좌측 보이기')
            except tk.TclError:
                pass
        elif want_visible and not self._left_pane_visible:
            try:
                self.paned.insert(0, self.left_container, weight=4)
                self._left_pane_visible = True
                self.left_pane_toggle.configure(text='◀ 좌측 숨기기')
            except tk.TclError:
                pass

    def restore_work_log_layout(self):
        """외부 설정 복원(load_tab_config / 창 최대화 직후) 시 프레임 상태를 재적용한다."""
        if not self.winfo_exists():
            return
        self._load_work_log_layout_prefs()
        self._apply_initial_pane_layout()

    def sync_layout_to_config(self):
        """현재 좌우 프레임 상태와 결과서 컬럼 너비를 main_app.tab_config에 반영한다."""
        cfg = self._get_host_tab_config()
        if cfg is None:
            return
        cfg['work_log_left_visible'] = bool(self._left_pane_visible)
        cfg['work_log_sash_ratio'] = round(float(self._saved_sash_ratio), 4)
        if self._left_pane_width is not None:
            cfg['work_log_sash_pos'] = int(self._left_pane_width)
        if hasattr(self, 'ndt_tree') and self.ndt_tree.winfo_exists():
            widths = {}
            for col in self.ndt_tree['columns']:
                try:
                    widths[col] = int(self.ndt_tree.column(col, 'width'))
                except (TypeError, ValueError, tk.TclError):
                    pass
            if widths:
                cfg['work_log_ndt_col_widths'] = widths

    def capture_visible_sash_before_close(self):
        """창 종료 시점에 화면에 표시 중인 좌우 분할선이 있으면 최신 비율을 캡처한다."""
        try:
            if (self.winfo_exists() and self.paned.winfo_ismapped()
                    and self._left_pane_visible and self.paned.winfo_width() > 200):
                pos = self.paned.sashpos(0)
                if pos > 50:
                    self._left_pane_width = pos
                    self._saved_sash_ratio = max(0.08, min(0.92, pos / self.paned.winfo_width()))
        except tk.TclError:
            pass
        self.sync_layout_to_config()

    def _save_work_log_layout(self, force=False):
        self.sync_layout_to_config()
        app = getattr(self, 'main_app', None)
        save_fn = getattr(app, 'save_tab_config', None) if app is not None else None
        if callable(save_fn):
            try:
                save_fn(force=force)
            except TypeError:
                save_fn()

    def _on_paned_configure(self, event):
        if event.widget is not self.paned:
            return
        width = event.width
        if width <= 200 or width == self._last_paned_width:
            return
        self._last_paned_width = width
        if self._left_pane_visible and self._sash_press_pos is None:
            target = max(120, min(int(round(width * self._saved_sash_ratio)), width - 120))
            try:
                self.paned.sashpos(0, target)
            except tk.TclError:
                pass

    def _on_paned_press(self, event):
        if event.widget is self.paned and self._left_pane_visible:
            try:
                self._sash_press_pos = self.paned.sashpos(0)
            except tk.TclError:
                self._sash_press_pos = None

    def _on_paned_release(self, event):
        prev_pos = self._sash_press_pos
        self._sash_press_pos = None
        if prev_pos is None or not self._left_pane_visible:
            return

        def _capture_sash():
            if not self.winfo_exists() or not self._left_pane_visible:
                return
            try:
                total_w = self.paned.winfo_width()
                cur_pos = self.paned.sashpos(0)
            except tk.TclError:
                return
            if total_w > 200 and abs(cur_pos - prev_pos) >= 2:
                self._left_pane_width = cur_pos
                self._saved_sash_ratio = max(0.08, min(0.92, cur_pos / total_w))
                self._save_work_log_layout(force=True)

        _capture_sash()
        try:
            self.after_idle(_capture_sash)
        except tk.TclError:
            pass

    def _apply_initial_pane_layout(self):
        try:
            self.update_idletasks()
            total_width = self.paned.winfo_width()
            if total_width > 200 and self._left_pane_visible:
                target = max(120, min(int(round(total_width * self._saved_sash_ratio)), total_width - 120))
                self.paned.sashpos(0, target)
                self._last_paned_width = total_width
        except (tk.TclError, AttributeError):
            pass

    def _ensure_calendar_work_summary(self, event=None):
        """달력 팝업이 처음 열릴 때 CalendarWorkSummary를 지연 로드한다."""
        if getattr(self, 'calendar_work_summary', None) is None:
            from site_apps.central.src.views.calendar_work_summary import CalendarWorkSummary
            self.calendar_work_summary = CalendarWorkSummary(
                self.date_entry, self,
                site_getter=lambda: '중앙지사',
                frame_getter=lambda: getattr(getattr(self, 'main_app', None), 'daily_usage_df', None),
                history_getter=self.load_history)
        return self.calendar_work_summary

    def _toggle_left_pane(self):
        if self._left_pane_visible:
            width = self.paned.winfo_width()
            pos = self.paned.sashpos(0)
            if width > 200 and pos > 50:
                self._left_pane_width = pos
                self._saved_sash_ratio = max(0.08, min(0.92, pos / width))
            self.paned.forget(self.left_container)
            self._left_pane_visible = False
            self.left_pane_toggle.configure(text='▶ 좌측 보이기')
            self._save_work_log_layout(force=True)
        else:
            self.paned.insert(0, self.left_container, weight=4)
            self._left_pane_visible = True
            self.paned.update_idletasks()
            width = self.paned.winfo_width()
            if width > 200:
                position = int(round(width * self._saved_sash_ratio))
                self.paned.sashpos(0, max(120, min(position, width - 120)))
                self._last_paned_width = width
            elif self._left_pane_width is not None:
                self.paned.sashpos(0, max(0, self._left_pane_width))
            self.left_pane_toggle.configure(text='◀ 좌측 숨기기')
            self._save_work_log_layout(force=True)

    def _build_basic_info(self, parent):
        # --- Top Section: General Info & Unified Work Toolbar ---
        top_frame = ttk.LabelFrame(parent, text="📋 기본 정보 및 통합 작업 도구", padding=6)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        # 1. 좌측: 기본 정보 입력 및 저장
        left_box = ttk.LabelFrame(top_frame, text="📌 기본 정보", padding=(6, 3))
        
        ttk.Label(left_box, text="검사일자:").grid(row=0, column=0, padx=(0, 4), pady=2, sticky="e")
        self.date_entry = DateEntry(left_box, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=(0, 6), pady=2, sticky="w")
        self.date_entry.bind("<<DateEntrySelected>>", self.on_date_change)
        self.date_entry.bind("<Return>", self.on_date_change)
        self.calendar_work_summary = None
        self.date_entry._top_cal.bind('<Map>', self._ensure_calendar_work_summary, add='+')
        self.date_entry.bind('<Button-1>', self._ensure_calendar_work_summary, add='+')

        self.date_status_label = tk.Label(
            left_box, text="○ 저장되지 않은 날짜", fg="#777777",
            font=('맑은 고딕', 9, 'bold')
        )
        self.date_status_label.grid(row=1, column=0, columnspan=2, padx=(0, 6), sticky="w")
        
        ttk.Label(left_box, text="날씨:").grid(row=0, column=2, padx=(0, 4), pady=2, sticky="e")
        self.weather_entry = ttk.Entry(left_box, width=8)
        self.weather_entry.grid(row=0, column=3, padx=(0, 6), pady=2, sticky="w")
        
        btn_calc = ttk.Button(left_box, text="💾 자동 집계/저장", command=self.auto_calculate_and_save, style="Action.TButton")
        btn_calc.grid(row=0, column=4, padx=(0, 3), pady=2)
        
        btn_delete = ttk.Button(left_box, text="🗑️ 삭제", command=self.delete_current_date_data, style="Danger.TButton")
        btn_delete.grid(row=0, column=5, padx=0, pady=2)
        
        btn_export = ttk.Button(left_box, text="📊 엑셀 출력 (일보)", command=self.export_excel, style="Action.TButton")
        btn_export.grid(row=1, column=4, columnspan=2, padx=0, pady=(2, 0), sticky="ew")
        
        # 2. 중앙: 보고서 발행
        mid_box = ttk.LabelFrame(top_frame, text="📑 보고서 발행", padding=(6, 3))
        
        btn_export_ndt = ttk.Button(mid_box, text="📗 NDT 누계대장", command=self.export_ndt_summary, style="Primary.TButton")
        btn_export_ndt.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        btn_export_monthly = ttk.Button(mid_box, text="📈 월간진도보고서", command=self.export_monthly_report, style="Info.TButton")
        btn_export_monthly.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        btn_weekly_report = ttk.Button(mid_box, text="🗓️ 주간업무보고", command=lambda: self.main_app.export_weekly_report() if hasattr(self, 'main_app') else None, style="Secondary.TButton")
        btn_weekly_report.grid(row=1, column=0, columnspan=2, padx=2, pady=(2, 0), sticky="ew")

        # 3. 중앙 우측: 비파괴검사결과서 편집 도구 (2단 배치)
        right_box = ttk.LabelFrame(top_frame, text="✏️ 비파괴검사결과서 편집", padding=(6, 3))

        # 1단: 행 추가/삭제 및 작업내역 검색
        ttk.Button(right_box, text="➕ 10행 추가", command=lambda: self._append_ndt_rows(10, focus=True), style="Primary.TButton").grid(row=0, column=0, padx=2, pady=1, sticky="w")
        ttk.Button(right_box, text="🗑️ 선택 행 삭제", command=self._delete_selected_ndt_row, style="Danger.TButton").grid(row=0, column=1, padx=2, pady=1, sticky="w")
        ttk.Button(right_box, text="🔍 작업내역 검색·수정", command=self.open_history_search_dialog, style="Info.TButton").grid(row=0, column=2, padx=2, pady=1, sticky="w")

        # 2단: 일괄 변경 도구 3종 (라인번호, 구간, 공사구분)
        ttk.Button(right_box, text="✏️ 라인번호 일괄", command=self._edit_ndt_lines, style="Secondary.TButton").grid(row=1, column=0, padx=2, pady=1, sticky="w")
        ttk.Button(right_box, text="✏️ 구간 일괄", command=self._edit_ndt_sections, style="Secondary.TButton").grid(row=1, column=1, padx=2, pady=1, sticky="w")
        ttk.Button(right_box, text="🏷️ 공사구분 일괄", command=self._edit_ndt_work_categories, style="Warning.TButton").grid(row=1, column=2, padx=2, pady=1, sticky="w")

        self.ndt_selection_vars = []
        self.ndt_selection_text = tk.StringVar(master=right_box, value="")
        self.ndt_selection_label = ttk.Label(
            right_box, textvariable=self.ndt_selection_text, foreground="#666666", wraplength=300
        )
        self.ndt_selection_label.grid(
            row=2, column=0, columnspan=3, padx=2, pady=1, sticky="w"
        )
        self.ndt_selection_label.grid_remove()

        # 4. 우측: 선택 행 핫태핑 및 공정사진 (2단 배치)
        extra_box = ttk.LabelFrame(top_frame, text="🔥 선택 행 핫태핑 · 📸 공정사진", padding=(6, 3))

        hot_bar = ttk.Frame(extra_box)
        hot_bar.grid(row=0, column=0, sticky="w", pady=1)
        ttk.Button(
            hot_bar, text="🔥 선택 행 핫태핑 입력/수정", command=self._edit_hot_tapping, style="Warning.TButton"
        ).pack(side="left", padx=(2, 6))
        self.hot_summary_var = tk.StringVar()
        hot_summary = ttk.Label(extra_box, textvariable=self.hot_summary_var, foreground="#444444", wraplength=260)
        hot_summary.grid(row=2, column=0, sticky="w", padx=2)

        photo_bar = ttk.Frame(extra_box)
        photo_bar.grid(row=1, column=0, sticky="w", pady=1)
        ttk.Button(
            photo_bar, text="📷 사진 추가", command=self.add_process_photo, style="Primary.TButton"
        ).pack(side="left", padx=2)
        ttk.Button(
            photo_bar, text="🖼️ 사진 관리", command=self.manage_process_photos, style="Secondary.TButton"
        ).pack(side="left", padx=(2, 6))
        self.photo_selection_var = tk.StringVar(value="")
        photo_selection = ttk.Label(extra_box, textvariable=self.photo_selection_var, foreground="#666666", wraplength=260)
        photo_selection.grid(row=3, column=0, sticky="w", padx=2)

        def refresh_extra_labels(*_):
            has_selection = bool(self.photo_selection_var.get())
            if has_selection:
                photo_selection.grid()
            else:
                photo_selection.grid_remove()
            if has_selection and self.hot_summary_var.get():
                hot_summary.grid()
            else:
                hot_summary.grid_remove()

        self.photo_selection_var.trace_add("write", refresh_extra_labels)
        self.hot_summary_var.trace_add("write", refresh_extra_labels)
        refresh_extra_labels()

        groups = (left_box, mid_box, right_box, extra_box)
        columns = None

        def arrange_toolbar(event):
            nonlocal columns
            widths = [box.winfo_reqwidth() + 6 for box in groups]
            available = event.width - 16  # Label frame border and padding
            if sum(widths) <= available:
                new_columns = 4
            elif max(widths[0], widths[2]) + max(widths[1], widths[3]) <= available:
                new_columns = 2
            else:
                new_columns = 1
            if new_columns == columns:
                return
            columns = new_columns
            for index, box in enumerate(groups):
                row, column = divmod(index, columns)
                box.grid(row=row, column=column, sticky="nsew", padx=(0, 6), pady=(0, 3))

        # Keep the initial requested size small so the window can resize freely.
        for index, box in enumerate(groups):
            box.grid(row=index, column=0, sticky="nsew", padx=(0, 6), pady=(0, 3))
        top_frame.bind("<Configure>", arrange_toolbar)

    def _build_left_pane(self, parent):
        # --- Middle Section: Work Quantities ---
        mid_frame = ttk.LabelFrame(parent, text="1. 작업 물량 및 누계 현황", padding=10)
        mid_frame.pack(fill="x", padx=5, pady=5)
        
        headers = ['방법', '규격', '예상량', '전일누계', '금일작업', '총누계', '공정률(%)', '불량', '불량률(%)', '비고']
        for col, h in enumerate(headers):
            ttk.Label(mid_frame, text=h, font=('맑은 고딕', 9, 'bold')).grid(row=0, column=col, padx=1, pady=2)
            
        self.qty_rows = [
            ('PAUT', '300A이상'), ('PAUT', '300A이상-야간'), ('PAUT', '250A'), ('PAUT', '250A-야간'), ('PAUT', '200A'), ('PAUT', '200A-야간'), 
            ('PAUT', '150A~125A'), ('PAUT', '150A~125A-야간'), ('PAUT', '100A이하'), ('PAUT', '100A이하-야간'), ('PAUT', '소계'),
            ('RT', '150A~100A'), ('RT', '150A~100A-야간'), ('RT', '80A이하'), ('RT', '80A이하-야간'), ('RT', '소계'),
            ('MT', '전체(주간)'), ('MT', '전체(야간)'),
            ('PT', '전체(주간)'), ('PT', '전체(야간)')
        ]
        
        self.qty_entries = {}
        self.default_qty = {
            ('PAUT', '300A이상'): '129', ('PAUT', '300A이상-야간'): '624',
            ('PAUT', '250A'): '4', ('PAUT', '250A-야간'): '1', ('PAUT', '200A'): '4',
            ('PAUT', '200A-야간'): '2', 
            ('PAUT', '150A~125A'): '1', ('PAUT', '150A~125A-야간'): '1',
            ('PAUT', '100A이하'): '1', ('PAUT', '100A이하-야간'): '1',
            ('PAUT', '소계'): '768',
            ('RT', '150A~100A'): '293', ('RT', '150A~100A-야간'): '43',
            ('RT', '80A이하'): '105', ('RT', '80A이하-야간'): '49',
            ('RT', '소계'): '490', ('MT', '전체(주간)'): '26',
            ('MT', '전체(야간)'): '1', ('PT', '전체(주간)'): '26',
            ('PT', '전체(야간)'): '1',
        }
        for row_idx, (method, spec) in enumerate(self.qty_rows, start=1):
            ttk.Label(mid_frame, text=method).grid(row=row_idx, column=0, padx=1, pady=2)
            ttk.Label(mid_frame, text=spec).grid(row=row_idx, column=1, padx=1, pady=2)
            
            row_dict = {}
            for col_idx, key in enumerate(['예상량', '전일누계', '금일작업', '총누계', '공정률', '불량', '불량률', '비고'], start=2):
                ent = ttk.Entry(mid_frame, width=7)
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=2)
                if key == '예상량':
                    val = self.default_qty.get((method, spec), '')
                    if val:
                        ent.insert(0, val)
                row_dict[key] = ent
            self.qty_entries[f"{method}_{spec}"] = row_dict
            
        # --- Bottom Section: Equip, Personnel, Remarks ---
        bot_frame = ttk.Frame(parent)
        bot_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Equipment
        eq_frame = ttk.LabelFrame(bot_frame, text="🚜 장비투입 현황", padding=10)
        eq_frame.pack(side="left", fill="y", padx=5)
        
        ttk.Label(eq_frame, text="장비명").grid(row=0, column=0)
        ttk.Label(eq_frame, text="금일").grid(row=0, column=1)
        ttk.Label(eq_frame, text="누계").grid(row=0, column=2)
        
        self.equip_rows = ['PAUT장비', 'PAUT프로브', 'PAUT스캐너', 'RT장비', 'MT장비']
        self.equip_entries = {}
        for i, eq in enumerate(self.equip_rows, start=1):
            ttk.Label(eq_frame, text=eq).grid(row=i, column=0, sticky="w")
            e_today = ttk.Entry(eq_frame, width=7)
            e_cum = ttk.Entry(eq_frame, width=7)
            e_today.grid(row=i, column=1, padx=2, pady=2)
            e_cum.grid(row=i, column=2, padx=2, pady=2)
            self.equip_entries[eq] = {'금일': e_today, '누계': e_cum}
            
        # Personnel
        pe_frame = ttk.LabelFrame(bot_frame, text="👥 금일 투입 인원", padding=10)
        pe_frame.pack(side="left", fill="y", padx=5)
        
        ttk.Label(pe_frame, text="구분").grid(row=0, column=0)
        ttk.Label(pe_frame, text="검사원").grid(row=0, column=1)
        ttk.Label(pe_frame, text="안전").grid(row=0, column=2)
        
        self.personnel_entries = {}
        for i, lbl in enumerate(['인원', '현장대리인', '누계'], start=1):
            ttk.Label(pe_frame, text=lbl).grid(row=i, column=0, sticky="w")
            ent1 = ttk.Entry(pe_frame, width=8)
            ent1.grid(row=i, column=1, padx=2, pady=2)
            self.personnel_entries[f'검사원_{lbl}'] = ent1
            
            ent2 = ttk.Entry(pe_frame, width=8)
            ent2.grid(row=i, column=2, padx=2, pady=2)
            self.personnel_entries[f'안전_{lbl}'] = ent2
            
        # Remarks
        rm_frame = ttk.LabelFrame(bot_frame, text="📝 특이사항 및 계획", padding=10)
        rm_frame.pack(side="left", fill="both", expand=True, padx=5)
        memo_toolbar = ttk.Frame(rm_frame)
        memo_toolbar.pack(fill="x", pady=(0, 5))
        ttk.Label(memo_toolbar, text="선택한 검사일자의 메모").pack(side="left")
        ttk.Button(memo_toolbar, text="📜 메모 이력", command=self.open_memo_history, style="Secondary.TButton").pack(side="right")
        self.remarks_text = tk.Text(rm_frame, width=20, height=8)
        self.remarks_text.pack(fill="both", expand=True)

    def open_memo_history(self):
        """저장된 일보 메모를 날짜/키워드로 찾아 해당 날짜로 이동한다."""
        history = self.load_history()
        memo_rows = []
        for date_key, day_data in history.items():
            memo = str(day_data.get('remarks', '') or '').strip()
            if memo:
                preview = ' '.join(memo.split())
                memo_rows.append((str(date_key), preview))
        memo_rows.sort(key=lambda item: item[0], reverse=True)

        window = tk.Toplevel(self)
        window.title("일보 메모 이력")
        window.geometry("760x480")
        window.minsize(560, 320)
        window.transient(self.winfo_toplevel())

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
                self.date_entry.set_date(datetime.strptime(date_key, '%Y-%m-%d').date())
                self.on_date_change()
                window.destroy()
            except ValueError:
                messagebox.showerror("날짜 오류", f"저장된 날짜 형식을 읽을 수 없습니다: {date_key}", parent=window)

        search_var.trace_add('write', refresh_rows)
        search_entry.bind('<Return>', lambda _event: move_to_selected())
        tree.bind('<Double-1>', move_to_selected)
        refresh_rows()
        search_entry.focus_set()

    def _extract_history_choices(self, history=None):
        """이력 데이터에서 업체/구간/라인/용접사 선택 목록을 추출한다."""
        if history is None:
            history = self.load_history()

        sections = set()
        lines = set()
        companies = set()
        welders = set(self.welder_names)
        for day_data in history.values():
            for row in day_data.get('ndt_results', []):
                section = str(row.get('구간', '') or '').strip()
                line_no = str(row.get('라인번호', '') or '').strip()
                company = str(row.get('업체', '') or '').strip()
                welder = self._normalize_welder_id(row.get('용접사', ''))
                if section:
                    sections.add(section)
                if line_no:
                    lines.add(line_no)
                if company:
                    companies.add(company)
                if welder:
                    welders.add(welder)

        self.company_section_mapping = extract_company_section_line_mapping(history)
        self.history_sections = [''] + sorted(sections)
        self.history_lines = [''] + sorted(lines)
        self.history_companies = [''] + sorted(companies)
        self.history_welders = [''] + [
            self._format_welder_display(welder_id)
            for welder_id in sorted(welders)
        ]

    def _build_right_pane(self, parent):
        # --- NDT Results Grid (Treeview + In-Place Cell Editor) ---
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", padx=5, pady=(4, 2))
        ttk.Label(title_frame, text="2. 비파괴검사결과서", font=("맑은 고딕", 11, "bold")).pack(side="left", padx=(5, 10))
        ttk.Label(
            title_frame,
            text="※ 셀 클릭 즉시 편집 · Tab/Shift+Tab 칸 이동 · Enter 윗줄 복사 및 다음 줄 이동 · 우클릭 빠른 메뉴",
            foreground="#555555"
        ).pack(side="left", padx=5)

        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        grid_frame.grid_rowconfigure(1, weight=1)
        grid_frame.grid_columnconfigure(1, weight=1)

        self.ndt_cols = (
            '업체', '공사구분', '검사방법', '구간', '작업구분',
            '핫태핑개소', '핫태핑차수', '핫태핑개소수', '라인번호', 'Joint No.',
            '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분',
            'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'
        )
        history = self.load_history()
        self._extract_history_choices(history)

        selection_button = ttk.Menubutton(grid_frame, text='☑ 선택 ▾', width=8, padding=(4, 1))
        selection_menu = tk.Menu(selection_button, tearoff=False)
        selection_menu.add_command(label='입력된 행 선택', command=self._select_filled_ndt_rows)
        selection_menu.add_command(label='전체 해제', command=self._clear_ndt_selection)
        selection_button.configure(menu=selection_menu)
        selection_button.grid(row=0, column=0, padx=(2, 6), pady=(0, 3), sticky='w')
        self._ndt_selection_menu = selection_menu

        ttk.Label(
            grid_frame,
            text="왼쪽 '선택' 열 클릭으로 개별 체크(Shift+클릭 범위 선택), 열 경계 드래그로 너비 조절이 가능합니다.",
            foreground="#64748b",
            font=("맑은 고딕", 8)
        ).grid(row=0, column=1, sticky='w', pady=(0, 3))

        top_win = self.winfo_toplevel()
        if not getattr(top_win, '_ndt_grid_style_configured', False):
            style = ttk.Style(top_win)
            style.configure("NDTGrid.Treeview", rowheight=26, font=("맑은 고딕", 9))
            style.configure("NDTGrid.Treeview.Heading", font=("맑은 고딕", 9, "bold"))
            top_win._ndt_grid_style_configured = True

        tree_columns = ('선택',) + self.ndt_cols
        self.ndt_tree = ttk.Treeview(
            grid_frame,
            columns=tree_columns,
            show='headings',
            selectmode='none',
            style="NDTGrid.Treeview",
        )

        def _on_yscroll(*args):
            self.right_yscroll.set(*args)
            self._reposition_active_cell_editor()

        def _on_xscroll(*args):
            self.right_xscroll.set(*args)
            self._reposition_active_cell_editor()

        self.right_yscroll = ttk.Scrollbar(
            grid_frame, orient="vertical",
            command=lambda *a: (self.ndt_tree.yview(*a), self._reposition_active_cell_editor())
        )
        self.right_xscroll = ttk.Scrollbar(
            grid_frame, orient="horizontal",
            command=lambda *a: (self.ndt_tree.xview(*a), self._reposition_active_cell_editor())
        )
        self.ndt_tree.configure(yscrollcommand=_on_yscroll, xscrollcommand=_on_xscroll)

        col_widths = {
            '선택': 60,
            '업체': 125,
            '공사구분': 85,
            '검사방법': 70,
            '구간': 85,
            '작업구분': 85,
            '핫태핑개소': 80,
            '핫태핑차수': 75,
            '핫태핑개소수': 80,
            '라인번호': 195,
            'Joint No.': 95,
            '관경': 65,
            '두께': 60,
            '용접사': 165,
            '구간정보': 230,
            '결과': 65,
            '규격': 105,
            '근무구분': 70,
            'RT_OR': 80,
            'RT_RE': 80,
            'PAUT': 72,
            'MT': 72,
            'PT': 72,
        }
        left_aligned = {'라인번호', '구간정보'}
        for col in tree_columns:
            label = {'RT_OR': 'RT_OR 전체', 'RT_RE': 'RT_RE 불량'}.get(col, col)
            if col == '선택':
                self.ndt_tree.heading(col, text='선택 ▾', command=self._popup_ndt_selection_menu)
            else:
                self.ndt_tree.heading(col, text=label)
            anchor = 'w' if col in left_aligned else 'center'
            self.ndt_tree.column(col, width=col_widths.get(col, 75), minwidth=48, stretch=False, anchor=anchor)

        self.ndt_tree.tag_configure('odd_row', background='#ffffff')
        self.ndt_tree.tag_configure('even_row', background='#f8fafc')
        self.ndt_tree.tag_configure('focused_row', background='#fef9c3')
        self.ndt_tree.tag_configure('checked_row', background='#dbeafe')
        self.ndt_tree.tag_configure('checked_focused_row', background='#bfdbfe')

        self.ndt_tree.grid(row=1, column=0, columnspan=2, sticky='nsew')
        self.right_yscroll.grid(row=1, column=2, sticky='ns')
        self.right_xscroll.grid(row=2, column=0, columnspan=2, sticky='ew')

        self._ndt_grid_frame = grid_frame
        self.ndt_grid_entries = []
        self._active_editor_state = None  # {'row': int, 'col': str, 'widget': Widget, 'full_values': list}
        self._init_ndt_inplace_editors()

        self.ndt_tree.bind('<ButtonPress-1>', self._on_ndt_tree_click)
        self.ndt_tree.bind('<ButtonRelease-1>', self._on_ndt_tree_release, add='+')
        self.ndt_tree.bind('<Button-3>', self._on_ndt_tree_right_click)
        self.ndt_tree.bind('<Configure>', lambda _e: self._reposition_active_cell_editor(), add='+')
        self.ndt_tree.bind('<Escape>', self._on_ndt_escape, add='+')

        self._append_ndt_rows(self._initial_ndt_rows)

    def _popup_ndt_selection_menu(self):
        try:
            x = self.ndt_tree.winfo_rootx() + 12
            y = self.ndt_tree.winfo_rooty() + 26
            self._ndt_selection_menu.tk_popup(x, y)
        finally:
            self._ndt_selection_menu.grab_release()

    def _init_ndt_inplace_editors(self):
        """Treeview 셀 위에 재사용할 단일 Entry, Combobox, 4분할 RT 구간정보 에디터를 1회 생성한다."""
        self._cell_entry_var = tk.StringVar(master=self)
        self._cell_entry = tk.Entry(
            self.ndt_tree,
            textvariable=self._cell_entry_var,
            font=("맑은 고딕", 9),
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightcolor="#2563eb",
            highlightbackground="#2563eb",
            bg="#ffffff",
        )

        self._cell_combo_var = tk.StringVar(master=self)
        self._cell_combo = ttk.Combobox(
            self.ndt_tree,
            textvariable=self._cell_combo_var,
            font=("맑은 고딕", 9),
        )

        self._cell_rt_frame = tk.Frame(
            self.ndt_tree,
            bg="#2563eb",
            bd=0,
            highlightthickness=1,
            highlightcolor="#2563eb",
            highlightbackground="#2563eb",
        )
        self._cell_rt_vars = [tk.StringVar(master=self) for _ in range(4)]
        self._cell_rt_entries = []
        for idx in range(4):
            sub_ent = tk.Entry(
                self._cell_rt_frame,
                textvariable=self._cell_rt_vars[idx],
                width=4,
                justify='center',
                font=("맑은 고딕", 9),
                relief="flat",
                bd=1,
                bg="#ffffff",
            )
            sub_ent.pack(side='left', fill='both', expand=True, padx=(0, 1 if idx < 3 else 0), pady=0)
            self._cell_rt_entries.append(sub_ent)

        # Single Entry bindings
        self._cell_entry.bind('<KeyRelease>', self._on_cell_entry_keyrelease)
        self._cell_entry.bind('<Return>', self._on_editor_return)
        self._cell_entry.bind('<Tab>', lambda e: self._on_editor_tab(e, forward=True))
        self._cell_entry.bind('<Shift-Tab>', lambda e: self._on_editor_tab(e, forward=False))
        self._cell_entry.bind('<ISO_Left_Tab>', lambda e: self._on_editor_tab(e, forward=False))
        self._cell_entry.bind('<Up>', lambda e: self._on_editor_vertical(e, delta=-1))
        self._cell_entry.bind('<Down>', lambda e: self._on_editor_vertical(e, delta=1))
        self._cell_entry.bind('<Escape>', self._on_ndt_escape)
        self._cell_entry.bind('<FocusOut>', self._on_editor_focus_out)
        self._cell_entry.bind('<Button-3>', self._on_editor_right_click)

        # Combobox bindings
        self._cell_combo.bind('<<ComboboxSelected>>', self._on_cell_combo_selected)
        self._cell_combo.bind('<KeyRelease>', self._on_cell_combo_keyrelease)
        self._cell_combo.bind('<Button-1>', self._on_cell_combo_click, add='+')
        self._cell_combo.bind('<Return>', self._on_editor_return)
        self._cell_combo.bind('<Tab>', lambda e: self._on_editor_tab(e, forward=True))
        self._cell_combo.bind('<Shift-Tab>', lambda e: self._on_editor_tab(e, forward=False))
        self._cell_combo.bind('<ISO_Left_Tab>', lambda e: self._on_editor_tab(e, forward=False))
        self._cell_combo.bind('<Control-Up>', lambda e: self._on_editor_vertical(e, delta=-1))
        self._cell_combo.bind('<Control-Down>', lambda e: self._on_editor_vertical(e, delta=1))
        self._cell_combo.bind('<Escape>', self._on_ndt_escape)
        self._cell_combo.bind('<FocusOut>', self._on_editor_focus_out)
        self._cell_combo.bind('<Button-3>', self._on_editor_right_click)

        # 4-slot RT section info bindings
        for sub_idx, sub_ent in enumerate(self._cell_rt_entries):
            sub_ent.bind('<KeyRelease>', self._on_rt_sub_keyrelease)
            sub_ent.bind('<Return>', self._on_editor_return)
            sub_ent.bind('<Tab>', lambda e, idx=sub_idx: self._on_rt_sub_tab(e, idx, forward=True))
            sub_ent.bind('<Shift-Tab>', lambda e, idx=sub_idx: self._on_rt_sub_tab(e, idx, forward=False))
            sub_ent.bind('<ISO_Left_Tab>', lambda e, idx=sub_idx: self._on_rt_sub_tab(e, idx, forward=False))
            sub_ent.bind('<Up>', lambda e: self._on_editor_vertical(e, delta=-1))
            sub_ent.bind('<Down>', lambda e: self._on_editor_vertical(e, delta=1))
            sub_ent.bind('<Escape>', self._on_ndt_escape)
            sub_ent.bind('<FocusOut>', self._on_rt_sub_focus_out)
            sub_ent.bind('<Button-3>', self._on_editor_right_click)

    def _update_ndt_row_tags(self, row_index):
        tree = getattr(self, 'ndt_tree', None)
        if tree is None or not tree.winfo_exists() or row_index >= len(self.ndt_grid_entries):
            return
        row = self.ndt_grid_entries[row_index]
        checked = row_index < len(self.ndt_selection_vars) and bool(self.ndt_selection_vars[row_index].get())
        focused = (getattr(self, 'selected_ndt_row', None) is row)
        if checked and focused:
            tag = 'checked_focused_row'
        elif checked:
            tag = 'checked_row'
        elif focused:
            tag = 'focused_row'
        else:
            tag = 'even_row' if (row_index % 2 == 1) else 'odd_row'
        mark = f"☑ {row_index + 1}" if checked else f"☐ {row_index + 1}"
        try:
            tree.set(row._iid, '선택', mark)
            tree.item(row._iid, tags=(tag,))
        except tk.TclError:
            pass

    def _sync_active_editor_if_matching(self, row_index, col_name):
        state = getattr(self, '_active_editor_state', None)
        if not state or state.get('in_live_edit'):
            return
        if state['row'] != row_index or state['col'] != col_name:
            return
        cell = self.ndt_grid_entries[row_index][col_name]
        if state['widget'] is self._cell_entry:
            self._cell_entry_var.set(cell.get())
        elif state['widget'] is self._cell_combo:
            self._cell_combo_var.set(cell.get())
        elif state['widget'] is self._cell_rt_frame and hasattr(cell, 'get_slots'):
            slots = cell.get_slots()
            for i in range(4):
                self._cell_rt_vars[i].set(slots[i])

    def _sync_active_combo_values(self, row_index, col_name):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['row'] != row_index or state['col'] != col_name:
            return
        if state['widget'] is self._cell_combo:
            vals = list(self.ndt_grid_entries[row_index][col_name]['values'])
            state['full_values'] = vals
            self._cell_combo['values'] = vals

    def _refresh_active_section_info_mode(self, row_index):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['row'] != row_index or state['col'] != '구간정보':
            return
        self._open_cell_editor(row_index, '구간정보', select_all=False)

    def _reposition_active_cell_editor(self):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return
        tree = getattr(self, 'ndt_tree', None)
        if tree is None or not tree.winfo_exists():
            return
        row_idx = state['row']
        col_name = state['col']
        if row_idx >= len(self.ndt_grid_entries):
            self._hide_active_cell_editor()
            return
        iid = self.ndt_grid_entries[row_idx]._iid
        col_num = self.ndt_cols.index(col_name) + 2
        bbox = tree.bbox(iid, f'#{col_num}')
        widget = state['widget']
        if not bbox:
            widget.place_forget()
            return
        x, y, w, h = bbox
        widget.place(x=x, y=y, width=max(w, 36), height=max(h, 22))

    def _hide_active_cell_editor(self):
        state = getattr(self, '_active_editor_state', None)
        self._active_editor_state = None
        if state and state.get('widget') is not None:
            try:
                state['widget'].place_forget()
            except tk.TclError:
                pass

    def _commit_active_cell_editor(self, run_triggers=True):
        state = getattr(self, '_active_editor_state', None)
        if not state or state.get('committing'):
            return
        state['committing'] = True
        try:
            row_idx = state['row']
            col_name = state['col']
            if row_idx >= len(self.ndt_grid_entries):
                return
            row = self.ndt_grid_entries[row_idx]
            cell = row[col_name]
            widget = state['widget']
            state['in_live_edit'] = True
            try:
                if widget is self._cell_entry:
                    new_val = self._cell_entry_var.get()
                    if col_name == '구간정보':
                        cell.set(new_val)
                    else:
                        cell.set(new_val)
                elif widget is self._cell_combo:
                    cell.set(self._cell_combo_var.get())
                elif widget is self._cell_rt_frame:
                    slots = [v.get() for v in self._cell_rt_vars]
                    cell.set_slots_from_editor(slots)
                    try:
                        sync_rt_widgets(row, from_details=True, all_rows=self._rt_context())
                    except ValueError:
                        pass
            finally:
                state['in_live_edit'] = False

            if run_triggers and cell.get() != state.get('initial_value') or state.get('combo_picked'):
                cell.event_generate('<FocusOut>')
        finally:
            state['committing'] = False

    def _prepare_dynamic_cell_options(self, row_idx, col_name):
        """셀 에디터를 열기 직전 현재 행의 업체/구간/검사방법에 맞춰 드롭다운 후보를 최신화한다."""
        row = self.ndt_grid_entries[row_idx]
        if col_name == '공사구분':
            comp = row['업체'].get().strip()
            matched_opts = None
            for k, opts in COMPANY_WORK_CATEGORIES.items():
                if k in comp:
                    matched_opts = opts
                    break
            row['공사구분']['values'] = matched_opts if matched_opts else ALL_WORK_CATEGORIES
        elif col_name == '구간':
            comp = row['업체'].get().strip()
            mapping = getattr(self, 'company_section_mapping', None)
            if mapping and comp:
                row['구간']['values'] = [''] + get_sections_for_company(mapping, comp)
            else:
                row['구간']['values'] = self.history_sections
        elif col_name == '라인번호':
            comp = row['업체'].get().strip()
            sec = row['구간'].get().strip()
            mapping = getattr(self, 'company_section_mapping', None)
            if mapping and (comp or sec):
                row['라인번호']['values'] = [''] + get_lines_for_company_and_section(mapping, comp, sec)
            else:
                row['라인번호']['values'] = self.history_lines
        elif col_name == '규격':
            self._sync_ndt_spec(row)

    def _open_cell_editor(self, row_idx, col_name, select_all=True, sub_idx=0):
        if row_idx < 0 or row_idx >= len(self.ndt_grid_entries):
            return
        if col_name not in self.ndt_cols:
            return
        prev = getattr(self, '_active_editor_state', None)
        if prev and (prev['row'] != row_idx or prev['col'] != col_name):
            self._commit_active_cell_editor()
        self._hide_active_cell_editor()

        tree = getattr(self, 'ndt_tree', None)
        if tree is None or not tree.winfo_exists():
            return

        row = self.ndt_grid_entries[row_idx]
        self._select_ndt_row(row)
        iid = row._iid
        col_num = self.ndt_cols.index(col_name) + 2
        col_id = f'#{col_num}'

        try:
            tree.see(iid)
            # Ensure target column is horizontally visible when off-screen
            bbox = tree.bbox(iid, col_id)
            if not bbox:
                self.update_idletasks()
                bbox = tree.bbox(iid, col_id)
            if not bbox:
                total_w = sum(int(tree.column(c, 'width')) for c in ('선택',) + self.ndt_cols)
                before_w = sum(int(tree.column(c, 'width')) for c in (('선택',) + self.ndt_cols)[:col_num - 1])
                if total_w > 0:
                    tree.xview_moveto(max(0.0, (before_w - 80) / total_w))
                    self.update_idletasks()
                    bbox = tree.bbox(iid, col_id)
        except tk.TclError:
            bbox = None

        if not bbox:
            return

        self._prepare_dynamic_cell_options(row_idx, col_name)
        cell = row[col_name]
        x, y, w, h = bbox

        if col_name == '구간정보':
            method = row['검사방법'].get().strip().upper()
            if method in ('PAUT', 'PT', 'MT'):
                widget = self._cell_entry
                self._cell_entry.configure(justify='left')
                self._cell_entry_var.set(cell.get())
                self._active_editor_state = {
                    'row': row_idx, 'col': col_name, 'widget': widget,
                    'initial_value': cell.get(), 'full_values': [],
                }
                widget.place(x=x, y=y, width=max(w, 40), height=max(h, 22))
                widget.focus_set()
                if select_all:
                    widget.selection_range(0, tk.END)
                    widget.icursor(tk.END)
            else:
                widget = self._cell_rt_frame
                slots = cell.get_slots()
                for i in range(4):
                    self._cell_rt_vars[i].set(slots[i])
                self._active_editor_state = {
                    'row': row_idx, 'col': col_name, 'widget': widget,
                    'initial_value': cell.get(), 'full_values': [],
                }
                widget.place(x=x, y=y, width=max(w, 100), height=max(h, 22))
                target_sub = max(0, min(3, sub_idx))
                ent = self._cell_rt_entries[target_sub]
                ent.focus_set()
                if select_all:
                    ent.selection_range(0, tk.END)
                    ent.icursor(tk.END)
            return

        if getattr(cell, '_is_combo', False):
            widget = self._cell_combo
            full_vals = list(cell['values'])
            widget.configure(
                state='readonly' if cell._readonly else 'normal',
                values=full_vals,
                justify='left' if col_name == '라인번호' else 'center',
            )
            self._cell_combo_var.set(cell.get())
            self._active_editor_state = {
                'row': row_idx, 'col': col_name, 'widget': widget,
                'initial_value': cell.get(), 'full_values': full_vals,
            }
            widget.place(x=x, y=y, width=max(w, 48), height=max(h, 22))
            widget.focus_set()
            if not cell._readonly and select_all:
                try:
                    widget.selection_range(0, tk.END)
                    widget.icursor(tk.END)
                except tk.TclError:
                    pass
        else:
            widget = self._cell_entry
            self._cell_entry.configure(justify='center')
            self._cell_entry_var.set(cell.get())
            self._active_editor_state = {
                'row': row_idx, 'col': col_name, 'widget': widget,
                'initial_value': cell.get(), 'full_values': [],
            }
            widget.place(x=x, y=y, width=max(w, 40), height=max(h, 22))
            widget.focus_set()
            if select_all:
                widget.selection_range(0, tk.END)
                widget.icursor(tk.END)

    def _on_cell_entry_keyrelease(self, event):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['widget'] is not self._cell_entry:
            return
        if event.keysym in ('Return', 'Tab', 'ISO_Left_Tab', 'Escape', 'Up', 'Down', 'Shift_L', 'Shift_R'):
            return
        row = self.ndt_grid_entries[state['row']]
        col = state['col']
        state['in_live_edit'] = True
        try:
            if col == '구간정보':
                row[col].set_slots_from_editor([self._cell_entry_var.get(), '', '', ''])
            else:
                row[col].set(self._cell_entry_var.get())
        finally:
            state['in_live_edit'] = False

    def _on_cell_combo_click(self, _event=None):
        state = getattr(self, '_active_editor_state', None)
        if state and state['widget'] is self._cell_combo:
            self._cell_combo['values'] = state.get('full_values', [])

    def _on_cell_combo_keyrelease(self, event):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['widget'] is not self._cell_combo:
            return
        if event.keysym in ('Return', 'Tab', 'ISO_Left_Tab', 'Escape', 'Up', 'Down', 'Left', 'Right', 'Alt_L', 'Alt_R', 'Control_L', 'Control_R', 'Shift_L', 'Shift_R'):
            return
        row = self.ndt_grid_entries[state['row']]
        col = state['col']
        cell = row[col]
        if cell._readonly:
            return
        typed = self._cell_combo_var.get()
        state['in_live_edit'] = True
        try:
            cell.set(typed)
        finally:
            state['in_live_edit'] = False
        q = typed.strip().lower()
        full_vals = state.get('full_values', [])
        if q:
            filtered = [v for v in full_vals if not v or q in str(v).lower()]
            self._cell_combo['values'] = filtered if filtered else full_vals
        else:
            self._cell_combo['values'] = full_vals

    def _on_cell_combo_selected(self, _event=None):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['widget'] is not self._cell_combo:
            return
        row_idx = state['row']
        col = state['col']
        row = self.ndt_grid_entries[row_idx]
        cell = row[col]
        state['combo_picked'] = True
        state['in_live_edit'] = True
        try:
            cell.set(self._cell_combo_var.get())
        finally:
            state['in_live_edit'] = False
        cell.event_generate('<<ComboboxSelected>>')

    def _on_rt_sub_keyrelease(self, event):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['widget'] is not self._cell_rt_frame:
            return
        if event.keysym in ('Return', 'Tab', 'ISO_Left_Tab', 'Escape', 'Up', 'Down'):
            return
        row = self.ndt_grid_entries[state['row']]
        state['in_live_edit'] = True
        try:
            row['구간정보'].set_slots_from_editor([v.get() for v in self._cell_rt_vars])
            try:
                sync_rt_widgets(row, from_details=True, all_rows=self._rt_context())
            except ValueError:
                pass
        finally:
            state['in_live_edit'] = False

    def _on_rt_sub_tab(self, _event, sub_idx, forward=True):
        if forward and sub_idx < 3:
            nxt = self._cell_rt_entries[sub_idx + 1]
            nxt.focus_set()
            nxt.selection_range(0, tk.END)
            return 'break'
        if not forward and sub_idx > 0:
            prv = self._cell_rt_entries[sub_idx - 1]
            prv.focus_set()
            prv.selection_range(0, tk.END)
            return 'break'
        return self._on_editor_tab(_event, forward=forward)

    def _on_rt_sub_focus_out(self, _event):
        state = getattr(self, '_active_editor_state', None)
        if not state or state['widget'] is not self._cell_rt_frame:
            return
        focused = self.focus_get()
        if focused in self._cell_rt_entries:
            return
        self._commit_active_cell_editor()

    def _on_editor_focus_out(self, event):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return
        if event is not None and getattr(event, 'widget', None) is not state.get('widget'):
            return
        self._commit_active_cell_editor()

    def _on_editor_tab(self, _event, forward=True):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return 'break'
        row_idx = state['row']
        col_idx = self.ndt_cols.index(state['col'])
        self._commit_active_cell_editor()
        if forward:
            col_idx += 1
            if col_idx >= len(self.ndt_cols):
                col_idx = 0
                row_idx += 1
                if row_idx >= len(self.ndt_grid_entries):
                    self._append_ndt_rows(10)
            self._open_cell_editor(row_idx, self.ndt_cols[col_idx], sub_idx=0)
        else:
            col_idx -= 1
            if col_idx < 0:
                if row_idx > 0:
                    row_idx -= 1
                    col_idx = len(self.ndt_cols) - 1
                else:
                    col_idx = 0
            self._open_cell_editor(row_idx, self.ndt_cols[col_idx], sub_idx=3)
        return 'break'

    def _on_editor_vertical(self, _event, delta=1):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return 'break'
        row_idx = state['row'] + delta
        col_name = state['col']
        self._commit_active_cell_editor()
        if row_idx < 0:
            row_idx = 0
        elif row_idx >= len(self.ndt_grid_entries):
            self._append_ndt_rows(10)
        self._open_cell_editor(row_idx, col_name)
        return 'break'

    def _on_editor_return(self, _event):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return 'break'
        row_idx = state['row']
        col_name = state['col']
        self._commit_active_cell_editor()
        self._hide_active_cell_editor()
        self.ndt_grid_entries[row_idx][col_name].event_generate('<Return>')
        return 'break'

    def _on_ndt_escape(self, _event=None):
        self._commit_active_cell_editor()
        self._hide_active_cell_editor()
        self._clear_ndt_selection()
        return 'break'

    def _on_ndt_tree_click(self, event):
        region = self.ndt_tree.identify_region(event.x, event.y)
        if region == 'separator':
            self._commit_active_cell_editor()
            self._hide_active_cell_editor()
            self._resizing_ndt_col = True
            return
        if region == 'heading':
            return
        self._commit_active_cell_editor()
        iid = self.ndt_tree.identify_row(event.y)
        col_id = self.ndt_tree.identify_column(event.x)
        if not iid or not col_id:
            self._hide_active_cell_editor()
            return
        try:
            row_idx = int(iid.split('_', 1)[1]) - 1
            col_num = int(col_id[1:])
        except (IndexError, ValueError):
            return
        if row_idx < 0 or row_idx >= len(self.ndt_grid_entries):
            return

        if col_num == 1:
            self._hide_active_cell_editor()
            if event.state & 0x0001:
                return self._select_ndt_range(row_idx)
            var = self.ndt_selection_vars[row_idx]
            var.set(not var.get())
            self._toggle_ndt_row(self.ndt_grid_entries[row_idx], var)
            return 'break'

        col_idx = col_num - 2
        if 0 <= col_idx < len(self.ndt_cols):
            col_name = self.ndt_cols[col_idx]
            self._open_cell_editor(row_idx, col_name)
            return 'break'

    def _on_ndt_tree_release(self, _event):
        if getattr(self, '_resizing_ndt_col', False):
            self._resizing_ndt_col = False

            def _commit_resize():
                if not self.winfo_exists():
                    return
                self._reposition_active_cell_editor()
                self._save_work_log_layout(force=True)

            _commit_resize()
            try:
                self.after_idle(_commit_resize)
            except tk.TclError:
                pass

    def _on_ndt_tree_right_click(self, event):
        self._commit_active_cell_editor()
        self._hide_active_cell_editor()
        iid = self.ndt_tree.identify_row(event.y)
        if not iid:
            return
        try:
            row_idx = int(iid.split('_', 1)[1]) - 1
        except (IndexError, ValueError):
            return
        if 0 <= row_idx < len(self.ndt_grid_entries):
            self._show_ndt_context_menu_for_row(event, self.ndt_grid_entries[row_idx])

    def _on_editor_right_click(self, event):
        state = getattr(self, '_active_editor_state', None)
        if not state:
            return
        row_idx = state['row']
        self._commit_active_cell_editor()
        if 0 <= row_idx < len(self.ndt_grid_entries):
            self._show_ndt_context_menu_for_row(event, self.ndt_grid_entries[row_idx])
        return 'break'

    def _show_ndt_context_menu_for_row(self, event, row_entries):
        self._select_ndt_row(row_entries)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="➕ 10행 추가", command=lambda: self._append_ndt_rows(10, focus=True))
        menu.add_command(label="🗑️ 선택 행 삭제", command=self._delete_selected_ndt_row)
        menu.add_separator()
        menu.add_command(label="✏️ 라인번호 일괄 변경...", command=self._edit_ndt_lines)
        menu.add_command(label="✏️ 구간 일괄 변경...", command=self._edit_ndt_sections)
        menu.add_command(label="🏷️ 공사구분 일괄 변경...", command=self._edit_ndt_work_categories)
        menu.add_command(label="🔍 작업내역 검색·수정...", command=self.open_history_search_dialog)
        menu.add_separator()
        menu.add_command(label="💾 자동 집계 및 저장", command=self.auto_calculate_and_save)
        menu.add_command(label="📊 엑셀 출력 (일보)", command=self.export_excel)
        menu.add_command(label="📗 NDT 누계대장 엑셀 출력", command=self.export_ndt_summary)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _append_ndt_rows(self, count=10, focus=False):
        first = len(self.ndt_grid_entries)
        for row_idx in range(first + 1, first + count + 1):
            self._create_ndt_row(row_idx)
        if focus and count:
            self._focus_ndt_joint(first)

    def _ensure_ndt_capacity(self, required):
        if required <= len(self.ndt_grid_entries):
            return
        target = max(10, ((required + 9) // 10) * 10)
        self._append_ndt_rows(target - len(self.ndt_grid_entries))

    def _focus_ndt_cell(self, index, col_name='Joint No.'):
        if index < 0 or index >= len(self.ndt_grid_entries):
            return
        row = self.ndt_grid_entries[index]
        self._select_ndt_row(row)
        tree = getattr(self, 'ndt_tree', None)
        if tree is not None and tree.winfo_exists():
            try:
                tree.see(row._iid)
            except tk.TclError:
                pass
            if tree.winfo_ismapped():
                self._open_cell_editor(index, col_name)

    def _focus_ndt_joint(self, index):
        self._focus_ndt_cell(index, 'Joint No.')

    def _create_ndt_row(self, row_idx):
        iid = f"row_{row_idx}"
        r_zero = row_idx - 1
        row_entries = _NDTRowProxy(self, r_zero, iid)

        selected = tk.BooleanVar(value=False)
        self.ndt_selection_vars.append(selected)
        selected.trace_add('write', lambda *_a, idx=r_zero: self._update_ndt_row_tags(idx))

        for c in self.ndt_cols:
            if c in ('작업구분', '핫태핑차수'):
                values = ['', '일반', '핫태핑', 'PAD', 'PAD(C심)', 'PAD(L심)', 'PAD(CAP)'] if c == '작업구분' else ['', '1차', '2차']
                row_entries[c] = _NDTCellProxy(row_entries, c, values=values, readonly=True)
            elif c == '검사방법':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'])
            elif c == '구간정보':
                row_entries[c] = _NDTSectionInfoProxy(row_entries)
            elif c == '관경':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=[''] + list(SIZE_LENGTH.keys()))
            elif c == '결과':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=['', '합격', '불합격', '재촬영', '보류'])
            elif c == '규격':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=[''])
            elif c == '근무구분':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=['', '주간', '야간', '재검'])
            elif c == '용접사':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=self.history_welders)
            elif c == '업체':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=self.history_companies)
            elif c == '공사구분':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=ALL_WORK_CATEGORIES)
            elif c == '구간':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=self.history_sections)
            elif c == '라인번호':
                row_entries[c] = _NDTCellProxy(row_entries, c, values=self.history_lines)
            else:
                row_entries[c] = _NDTCellProxy(row_entries, c)

        def on_ndt_change(event, r=row_entries):
            method = r['검사방법'].get().strip().upper()
            pipe = r['관경'].get().strip().upper()
            self._sync_ndt_spec(r)

            thk_mapping = {
                "1100": "11.1", "1000": "11.1", "900": "10.3", "850": "10.3", "800": "9.5",
                "750": "8.7", "700": "8.7", "650": "8.7", "600": "9.5", "550": "9.5",
                "500": "6.4", "450": "6.4", "400": "6.4", "350": "6.4", "300": "6.4", "250": "6.4", "200": "6.4",
                "150": "5.5", "125": "5.1", "100": "4.9", "80": "4.5", "65": "4.5", "50": "3.9",
                "40": "3.7", "32": "3.6", "25": "3.4", "20": "2.9"
            }
            lookup_size = str(pipe).replace("A", "").strip()
            if lookup_size in thk_mapping and r['작업구분'].get() != '핫태핑':
                r['두께'].set(thk_mapping[lookup_size])

            if pipe and pipe.isdigit():
                pipe += 'A'

            if r['작업구분'].get() == '핫태핑':
                from site_apps.central.src.hot_tapping import inspection_length
                try:
                    count = int(r['핫태핑개소수'].get())
                    if method in ('MT', 'PT') and count > 0:
                        length = inspection_length(pipe, r['핫태핑차수'].get()) * count
                        for col in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'):
                            r[col].set('')
                        r[method].set(f'{length:.4f}')
                except ValueError:
                    pass
                r['구간정보'].apply_layout()
                self._refresh_hot_summary()
                return

            target_col = None
            if method == 'PAUT': target_col = 'PAUT'
            elif method == 'RT': target_col = 'RT_OR'
            elif method == 'MT': target_col = 'MT'
            elif method == 'PT': target_col = 'PT'

            if r['작업구분'].get() == 'PAD(L심)':
                length_str = "1.0000"
            elif pipe in SIZE_LENGTH:
                length = SIZE_LENGTH[pipe]
                length_str = f"{length:.4f}"
                if method == 'RT':
                    try:
                        size_val = int(''.join(filter(str.isdigit, pipe)))
                        if size_val <= 150 and not retest_parent(r['Joint No.'].get()):
                            length_str = "3"
                        else:
                            length_str = ""
                    except ValueError:
                        length_str = ""
            else:
                length_str = ""

            if target_col and length_str and not r[target_col].get().strip():
                for col in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'):
                    r[col].set('')
                r[target_col].set(length_str)

            r['구간정보'].apply_layout()

        def update_work_category_options(r=row_entries):
            comp = r['업체'].get().strip()
            cat_cb = r.get('공사구분')
            if not cat_cb:
                return
            matched_opts = None
            for k, opts in COMPANY_WORK_CATEGORIES.items():
                if k in comp:
                    matched_opts = opts
                    break
            if matched_opts:
                cat_cb['values'] = matched_opts
                curr_val = cat_cb.get().strip()
                if curr_val and curr_val not in matched_opts:
                    cat_cb.set('')
            else:
                cat_cb['values'] = ALL_WORK_CATEGORIES

        def on_company_selected(r=row_entries):
            update_work_category_options(r)
            comp = r['업체'].get().strip()
            mapping = getattr(self, 'company_section_mapping', None)
            if mapping:
                sec_cb = r.get('구간')
                if sec_cb is not None:
                    sections = [''] + get_sections_for_company(mapping, comp)
                    sec_cb['values'] = sections
                    curr_sec = sec_cb.get().strip()
                    if curr_sec and curr_sec not in sections:
                        sec_cb.set('')
            on_section_selected(r=r)

        def on_section_selected(r=row_entries):
            comp = r['업체'].get().strip()
            sec = r['구간'].get().strip()
            mapping = getattr(self, 'company_section_mapping', None)
            if mapping:
                line_cb = r.get('라인번호')
                if line_cb is not None:
                    lines = [''] + get_lines_for_company_and_section(mapping, comp, sec)
                    line_cb['values'] = lines

        row_entries['업체'].bind('<<ComboboxSelected>>', lambda _e, r=row_entries: on_company_selected(r), add='+')
        row_entries['업체'].bind('<FocusOut>', lambda _e, r=row_entries: on_company_selected(r), add='+')
        row_entries['구간'].bind('<<ComboboxSelected>>', lambda _e, r=row_entries: on_section_selected(r), add='+')
        row_entries['구간'].bind('<FocusOut>', lambda _e, r=row_entries: on_section_selected(r), add='+')

        def on_hot_change(event=None, r=row_entries):
            if r['작업구분'].get() == '핫태핑':
                self._edit_hot_tapping(r)
            elif str(r['작업구분'].get()).startswith('PAD'):
                on_ndt_change(event, r)

        row_entries['작업구분'].bind('<<ComboboxSelected>>', on_hot_change)
        row_entries['검사방법'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['검사방법'].bind('<FocusOut>', on_ndt_change)
        row_entries['관경'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['관경'].bind('<FocusOut>', on_ndt_change)
        row_entries['Joint No.'].bind('<FocusOut>', lambda _e, r=row_entries: self._sync_rt_row(r), add='+')

        def on_joint_enter(event, r_idx=r_zero, r=row_entries):
            import re
            if r_idx == 0:
                if len(self.ndt_grid_entries) > 1:
                    self._focus_ndt_joint(1)
                return 'break'

            prev_row = self.ndt_grid_entries[r_idx - 1]
            r.begin_batch()
            try:
                curr_joint = r['Joint No.'].get().strip()
                if not curr_joint:
                    prev_joint = prev_row['Joint No.'].get().strip()
                    if prev_joint:
                        match = re.search(r'(\d+)$', prev_joint)
                        if match:
                            num_str = match.group(1)
                            prefix = prev_joint[:-len(num_str)]
                            next_num = str(int(num_str) + 1).zfill(len(num_str))
                            r['Joint No.'].set(prefix + next_num)

                for col in ['업체', '공사구분', '검사방법', '구간', '작업구분', '핫태핑개소', '핫태핑차수', '핫태핑개소수', '라인번호', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분']:
                    if col == '규격' and r['검사방법'].get().strip().upper() != prev_row['검사방법'].get().strip().upper():
                        continue
                    if col in ('구간정보', '결과') and (r['검사방법'].get().strip().upper() == 'RT' or prev_row['검사방법'].get().strip().upper() == 'RT'):
                        continue
                    if not r[col].get().strip():
                        prev_val = prev_row[col].get().strip()
                        if prev_val:
                            r[col].set(prev_val)
                            if col == '업체':
                                on_company_selected(r)
                            elif col == '구간':
                                on_section_selected(r)

                on_ndt_change(None, r)
                defaults = default_rt_copy(widget_rt_values(r), self._rt_context())
                if defaults.get('구간정보', '') != r['구간정보'].get():
                    r['구간정보'].set(defaults['구간정보'])
                self._sync_rt_row(r)
            finally:
                r.end_batch()

            if r_idx + 1 == len(self.ndt_grid_entries):
                self._append_ndt_rows(10)
            if r_idx + 1 < len(self.ndt_grid_entries):
                self._focus_ndt_joint(r_idx + 1)
            return 'break'

        for widget in row_entries.values():
            widget.bind('<Return>', on_joint_enter)
            widget.bind('<Button-1>', lambda _e, r=row_entries: self._select_ndt_row(r), add='+')
            widget.bind('<Escape>', self._clear_ndt_selection, add='+')

        tag = 'even_row' if (r_zero % 2 == 1) else 'odd_row'
        self.ndt_tree.insert('', tk.END, iid=iid, values=(f"☐ {row_idx}",) + ('',) * len(self.ndt_cols), tags=(tag,))
        self.ndt_grid_entries.append(row_entries)


    def _delete_selected_ndt_row(self):
        if hasattr(self, '_commit_active_cell_editor'):
            self._commit_active_cell_editor()
            self._hide_active_cell_editor()
        targets = [(i, row) for i, row in enumerate(self.ndt_grid_entries)
                   if self.ndt_selection_vars[i].get()]
        if not targets:
            messagebox.showwarning('행 삭제', '삭제할 행의 왼쪽 체크박스를 선택하세요.')
            return
        numbers = ', '.join(str(i + 1) for i, _ in targets)
        if not messagebox.askyesno(
            '행 삭제',
            f'선택한 {len(targets)}개 행 ({numbers}행)의 입력 내용을 삭제하시겠습니까?\n'
            '상단 저장 버튼을 눌러야 저장된 내역에 반영됩니다.'
        ):
            return
        self._compact_ndt_rows({i for i, _ in targets})
        self._refresh_hot_summary()
        self._clear_ndt_selection()
        messagebox.showinfo('행 삭제', '선택한 행을 삭제하고 나머지 행을 위로 정리했습니다. 상단 저장 버튼을 눌러 주세요.')

    def _compact_ndt_rows(self, deleted_indices):
        # Snapshot before writing: destination widgets may also be source rows.
        remaining = []
        for index, row in enumerate(self.ndt_grid_entries):
            if index in deleted_indices:
                continue
            values = {key: widget.get() for key, widget in row.items()}
            if any(str(value).strip() for value in values.values()):
                metadata = dict(getattr(row.get('구간정보'), '_rt_retake_meta', {}))
                remaining.append((values, metadata))
        for index, row in enumerate(self.ndt_grid_entries):
            values, metadata = remaining[index] if index < len(remaining) else ({}, {})
            # Set the method before the detail widget chooses its layout.
            for key, widget in row.items():
                if key == '구간정보':
                    continue
                value = values.get(key, '')
                if hasattr(widget, 'set'):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, value)
            details = row.get('구간정보')
            if details is not None:
                details.set(values.get('구간정보', ''))
                details._rt_retake_meta = metadata
            # Detail setters can recalculate RT counts; preserve the saved inputs.
            for key in ('RT_OR', 'RT_RE'):
                if key in row:
                    row[key].delete(0, tk.END)
                    row[key].insert(0, values.get(key, ''))
            self._sync_ndt_spec(row)

    def _replace_ndt_line(self, old_line, new_line, selected_only=False):
        old_line, new_line = old_line.strip(), new_line.strip()
        if not old_line or not new_line:
            raise ValueError('기존 라인번호와 새 라인번호를 입력하세요.')
        if old_line == new_line:
            raise ValueError('새 라인번호가 기존 번호와 같습니다.')
        targets = [row for index, row in enumerate(self.ndt_grid_entries)
                   if row['라인번호'].get().strip() == old_line
                   and (not selected_only or self.ndt_selection_vars[index].get())]
        if not targets:
            raise ValueError('변경할 행이 없습니다. 기존 라인번호와 체크한 행을 확인하세요.')
        for row in targets:
            row['라인번호'].set(new_line)
        if targets:
            choices = sorted({line for line in self.history_lines if line} | {new_line})
            self.history_lines = [''] + choices
            for row in self.ndt_grid_entries:
                row['라인번호']['values'] = self.history_lines
        return len(targets)

    def _replace_ndt_section(self, old_section, new_section, selected_only=False):
        old_section, new_section = old_section.strip(), new_section.strip()
        if not new_section or (not selected_only and not old_section):
            raise ValueError('기존 구간과 새 구간을 입력하세요.')
        if not selected_only and old_section == new_section:
            raise ValueError('새 구간이 기존 구간과 같습니다.')
        targets = [row for i, row in enumerate(self.ndt_grid_entries)
                   if (self.ndt_selection_vars[i].get() if selected_only
                       else row['구간'].get().strip() == old_section)]
        if not targets:
            raise ValueError('변경할 행이 없습니다. 구간 또는 체크한 행을 확인하세요.')
        for row in targets:
            row['구간'].set(new_section)
        self.history_sections = [''] + sorted(
            {s for s in self.history_sections if s} | {new_section})
        for row in self.ndt_grid_entries:
            row['구간']['values'] = self.history_sections
        return len(targets)

    def _edit_ndt_sections(self):
        sections = sorted({row['구간'].get().strip() for row in self.ndt_grid_entries
                           if row['구간'].get().strip()})
        selected = getattr(self, 'selected_ndt_row', None)
        initial = selected['구간'].get().strip() if selected is not None else ''
        dialog = tk.Toplevel(self)
        dialog.title('구간 일괄 변경')
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=15)
        body.pack(fill='both', expand=True)
        old = tk.StringVar(value=initial or (sections[0] if sections else ''))
        new = tk.StringVar()
        selected_only = tk.BooleanVar(value=True)
        preview = tk.StringVar()
        ttk.Label(body, text=f'현재 날짜: {self.date_entry.get()}').grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Radiobutton(body, text='기존 구간과 같은 모든 행', variable=selected_only,
                        value=False).grid(row=1, column=0, columnspan=2, sticky='w')
        ttk.Radiobutton(body, text='왼쪽 체크박스로 선택한 행만', variable=selected_only,
                        value=True).grid(row=2, column=0, columnspan=2, sticky='w')
        ttk.Label(body, text='기존 구간').grid(row=3, column=0, sticky='w', padx=(0, 10))
        old_entry = ttk.Combobox(body, textvariable=old, values=sections, state='readonly', width=28)
        old_entry.grid(row=3, column=1, pady=4)
        ttk.Label(body, text='새 구간').grid(row=4, column=0, sticky='w')
        entry = ttk.Combobox(body, textvariable=new, values=self.history_sections, width=28)
        entry.grid(row=4, column=1, pady=4)
        ttk.Label(body, textvariable=preview).grid(row=5, column=0, columnspan=2, sticky='w', pady=10)
        ttk.Label(body, text='변경 후 상단 저장 버튼을 눌러 주세요.').grid(
            row=6, column=0, columnspan=2, sticky='w')

        def update_preview(*args):
            only = selected_only.get()
            old_entry.configure(state='disabled' if only else 'readonly')
            count = (sum(bool(v.get()) for v in self.ndt_selection_vars) if only else
                     sum(row['구간'].get().strip() == old.get() for row in self.ndt_grid_entries)
                     if old.get() else 0)
            preview.set(f'변경 대상: {count}행')

        def apply():
            try:
                count = self._replace_ndt_section(old.get(), new.get(), selected_only.get())
            except ValueError as error:
                messagebox.showwarning('구간 일괄 변경', str(error), parent=dialog)
                return
            dialog.destroy()
            messagebox.showinfo('구간 일괄 변경', f'{count}행을 변경했습니다. 상단 저장 버튼을 눌러 주세요.')

        old.trace_add('write', update_preview)
        selected_only.trace_add('write', update_preview)
        update_preview()
        buttons = ttk.Frame(body)
        buttons.grid(row=7, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text='변경 적용', command=apply).pack(side='left', padx=4)
        ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left')
        dialog.grab_set()
        entry.focus_set()

    def _edit_ndt_lines(self):
        lines = sorted({row['라인번호'].get().strip() for row in self.ndt_grid_entries
                        if row['라인번호'].get().strip()})
        if not lines:
            messagebox.showinfo('라인번호 일괄 변경', '변경할 라인번호가 없습니다.')
            return
        selected = getattr(self, 'selected_ndt_row', None)
        initial = selected['라인번호'].get().strip() if selected is not None else lines[0]
        dialog = tk.Toplevel(self)
        dialog.title('라인번호 일괄 변경')
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=15)
        body.pack(fill='both', expand=True)
        old = tk.StringVar(value=initial if initial in lines else lines[0])
        new = tk.StringVar()
        selected_only = tk.BooleanVar(value=True)
        preview = tk.StringVar()
        ttk.Label(body, text=f'현재 날짜: {self.date_entry.get()}').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Radiobutton(body, text='체크한 행 중 기존 라인번호가 같은 행', variable=selected_only,
                        value=True).grid(row=1, column=0, columnspan=2, sticky='w')
        ttk.Radiobutton(body, text='기존 라인번호와 같은 모든 행', variable=selected_only,
                        value=False).grid(row=2, column=0, columnspan=2, sticky='w')
        ttk.Label(body, text='기존 라인번호').grid(row=3, column=0, sticky='w', padx=(0, 10))
        ttk.Combobox(body, textvariable=old, values=lines, state='readonly', width=42).grid(row=3, column=1, pady=4)
        ttk.Label(body, text='새 라인번호').grid(row=4, column=0, sticky='w')
        entry = ttk.Entry(body, textvariable=new, width=45)
        entry.grid(row=4, column=1, pady=4)
        ttk.Label(body, textvariable=preview).grid(row=5, column=0, columnspan=2, sticky='w', pady=10)
        ttk.Label(body, text='변경 후 상단 저장 버튼을 눌러 주세요.').grid(row=6, column=0, columnspan=2, sticky='w')

        def update_preview(*args):
            count = sum(row['라인번호'].get().strip() == old.get()
                        and (not selected_only.get() or self.ndt_selection_vars[i].get())
                        for i, row in enumerate(self.ndt_grid_entries))
            preview.set(f'변경 대상: {count}행')

        def apply():
            try:
                count = self._replace_ndt_line(old.get(), new.get(), selected_only.get())
            except ValueError as error:
                messagebox.showwarning('라인번호 일괄 변경', str(error), parent=dialog)
                return
            dialog.destroy()
            messagebox.showinfo('라인번호 일괄 변경', f'{count}행을 변경했습니다. 상단 저장 버튼을 눌러 주세요.')

        old.trace_add('write', update_preview)
        selected_only.trace_add('write', update_preview)
        update_preview()
        buttons = ttk.Frame(body)
        buttons.grid(row=7, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text='변경 적용', command=apply).pack(side='left', padx=4)
        ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left')
        dialog.grab_set()
        entry.focus_set()

    def _edit_ndt_work_categories(self):
        selected = getattr(self, 'selected_ndt_row', None)
        initial = selected['공사구분'].get().strip() if selected is not None and '공사구분' in selected else ''
        dialog = tk.Toplevel(self)
        dialog.title('공사구분 일괄 변경')
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=15)
        body.pack(fill='both', expand=True)

        new = tk.StringVar(value=initial)
        selected_only = tk.BooleanVar(value=True)
        preview = tk.StringVar()

        ttk.Label(body, text=f'현재 날짜: {self.date_entry.get()}').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Radiobutton(body, text='왼쪽 체크박스로 선택한 행만 변경', variable=selected_only, value=True).grid(row=1, column=0, columnspan=2, sticky='w')
        ttk.Radiobutton(body, text='현재 일보의 모든 유효 행 변경', variable=selected_only, value=False).grid(row=2, column=0, columnspan=2, sticky='w')

        ttk.Label(body, text='새 공사구분').grid(row=3, column=0, sticky='w', pady=(8, 4), padx=(0, 10))
        entry = ttk.Combobox(body, textvariable=new, values=ALL_WORK_CATEGORIES, state='readonly', width=26)
        entry.grid(row=3, column=1, pady=(8, 4))

        ttk.Label(body, textvariable=preview, foreground="#1d4ed8").grid(row=4, column=0, columnspan=2, sticky='w', pady=10)
        ttk.Label(body, text='※ 변경 후 상단 [💾 자동 집계/저장] 버튼을 눌러야 최종 저장됩니다.', foreground="#666666").grid(row=5, column=0, columnspan=2, sticky='w')

        def update_preview(*args):
            only = selected_only.get()
            if only:
                count = sum(bool(v.get()) for v in self.ndt_selection_vars)
            else:
                count = sum(bool(any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법') if k in row)) for row in self.ndt_grid_entries)
            preview.set(f'변경 대상: {count}행')

        def apply():
            val = new.get().strip()
            only = selected_only.get()
            count = 0
            for idx, row in enumerate(self.ndt_grid_entries):
                if '공사구분' not in row:
                    continue
                if only:
                    if idx < len(self.ndt_selection_vars) and self.ndt_selection_vars[idx].get():
                        row['공사구분'].set(val)
                        count += 1
                else:
                    if any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법') if k in row):
                        row['공사구분'].set(val)
                        count += 1
            if not count:
                messagebox.showwarning('공사구분 일괄 변경', '변경할 행이 없습니다. 체크한 행을 확인하세요.', parent=dialog)
                return
            dialog.destroy()
            messagebox.showinfo('공사구분 일괄 변경', f'{count}개 행의 공사구분을 [{val or "공란"}]으로 변경했습니다.\n상단 [💾 자동 집계/저장]을 눌러 저장해 주세요.', parent=self)

        selected_only.trace_add('write', update_preview)
        update_preview()

        buttons = ttk.Frame(body)
        buttons.grid(row=6, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text='변경 적용', command=apply).pack(side='left', padx=4)
        ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left')
        dialog.grab_set()
        entry.focus_set()

    def open_history_search_dialog(self):
        from site_apps.central.src.views.work_log_search_dialog import (
            open_history_search_dialog as _open_history_search_dialog,
        )
        return _open_history_search_dialog(self)

    def _sync_ndt_spec(self, row):
        import re
        method = row['검사방법'].get().strip().upper()
        spec = row['규격']
        spec['values'] = ['', '3 1/3 X 12"', '3 1/3 X 6"'] if method == 'RT' else ['']
        if method and method != 'RT' and re.fullmatch(r'3\s*1\s*/\s*3\s*[X×]\s*(?:12|6)\s*["”″]?', spec.get().strip(), re.IGNORECASE):
            spec.set('')

    def _edit_hot_tapping(self, row=None):
        from site_apps.central.src.hot_tapping import LENGTHS, inspection_length
        row = row if row is not None else getattr(self, 'selected_ndt_row', None)
        if row is None:
            messagebox.showwarning('핫태핑', '먼저 검사 행을 선택하세요.')
            return
        dialog = tk.Toplevel(self)
        dialog.title('핫태핑 검사 입력')
        dialog.transient(self.winfo_toplevel())
        variables = {}
        fields = [('검사방법', ['MT', 'PT'], 'MT'), ('핫태핑개소', None, ''),
                  ('핫태핑차수', ['1차', '2차'], '1차'),
                  ('관경', [f'{n}A' for n in LENGTHS], '600A'), ('핫태핑개소수', None, '1')]
        for index, (key, choices, default) in enumerate(fields):
            value = row[key].get().strip()
            if choices and value not in choices:
                value = default
            variables[key] = tk.StringVar(value=value or default)
            ttk.Label(dialog, text=key).grid(row=index, column=0, padx=12, pady=6)
            widget = (ttk.Combobox(dialog, textvariable=variables[key], values=choices, state='readonly')
                      if choices else ttk.Entry(dialog, textvariable=variables[key]))
            widget.grid(row=index, column=1, padx=12, pady=6)
        preview = tk.StringVar()
        def update_preview(*_):
            try:
                count = int(variables['핫태핑개소수'].get())
                if count <= 0:
                    raise ValueError('개소 수는 양의 정수여야 합니다.')
                length = inspection_length(variables['관경'].get(), variables['핫태핑차수'].get()) * count
                preview.set(f'검사 길이: {length:.4f} m')
                return length
            except ValueError as exc:
                preview.set(str(exc))
                return None
        for var in variables.values():
            var.trace_add('write', update_preview)
        ttk.Label(dialog, textvariable=preview).grid(row=5, column=0, columnspan=2, pady=8)
        def apply():
            length = update_preview()
            if length is None or not variables['핫태핑개소'].get().strip():
                messagebox.showwarning('핫태핑', '개소명/번호와 올바른 검사 조건을 입력하세요.', parent=dialog)
                return
            for key, var in variables.items():
                row[key].delete(0, tk.END)
                row[key].insert(0, var.get().strip()) if not isinstance(row[key], ttk.Combobox) else row[key].set(var.get())
            row['작업구분'].set('핫태핑')
            for key in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'):
                row[key].delete(0, tk.END)
            row[variables['검사방법'].get()].insert(0, f'{length:.4f}')
            row['구간정보'].apply_layout()
            self._refresh_hot_summary()
            dialog.destroy()
        ttk.Button(dialog, text='적용', command=apply).grid(row=6, column=0, columnspan=2, pady=10)
        update_preview()
        dialog.grab_set()

    def _refresh_hot_summary(self):
        totals = {'MT': 0.0, 'PT': 0.0}
        locations = set()
        for row in self.ndt_grid_entries:
            if row['작업구분'].get() != '핫태핑':
                continue
            locations.add(row['핫태핑개소'].get().strip())
            for method in totals:
                try:
                    totals[method] += float(row[method].get() or 0)
                except ValueError:
                    pass
        location_count = len(locations - {''})
        parts = [f"{location_count}개소명"] if location_count else []
        parts.extend(f"{method} {total:.4f} m" for method, total in totals.items() if round(total, 4))
        self.hot_summary_var.set("핫태핑: " + " / ".join(parts) if parts else "")

    def _validate_hot_rows(self):
        from site_apps.central.src.hot_tapping import inspection_length
        pending = []
        for index, row in enumerate(self.ndt_grid_entries, 1):
            if row['작업구분'].get() != '핫태핑':
                continue
            try:
                method = row['검사방법'].get().strip().upper()
                count = int(row['핫태핑개소수'].get())
                if method not in ('MT', 'PT') or count <= 0 or not row['핫태핑개소'].get().strip():
                    raise ValueError('MT/PT, 개소명과 양의 정수 개소 수를 입력하세요.')
                length = inspection_length(row['관경'].get(), row['핫태핑차수'].get()) * count
                pending.append((row, method, length))
            except ValueError as exc:
                messagebox.showwarning('핫태핑 입력 확인', f'{index}행: {exc}')
                return False
        for row, method, length in pending:
            for key in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'):
                row[key].delete(0, tk.END)
            row[method].insert(0, f'{length:.4f}')
        self._refresh_hot_summary()
        return True

    def _select_ndt_row(self, row_entries):
        """Focus an editing row for row-specific tools without changing batch checks."""
        prev = getattr(self, 'selected_ndt_row', None)
        self.selected_ndt_row = row_entries
        if prev is not None and prev is not row_entries and hasattr(prev, '_row_index'):
            self._update_ndt_row_tags(prev._row_index)
        if hasattr(row_entries, '_row_index'):
            self._update_ndt_row_tags(row_entries._row_index)
        method = row_entries['검사방법'].get().strip() or '-'
        section = row_entries['구간'].get().strip() or '-'
        joint = row_entries['Joint No.'].get().strip() or '-'
        welder = row_entries['용접사'].get().strip() or '-'
        self.photo_selection_var.set(
            f"선택: {method} / {section} / Joint {joint} / {welder}"
        )

    def _refresh_ndt_selection_text(self):
        count = sum(bool(var.get()) for var in self.ndt_selection_vars)
        self.ndt_selection_text.set(f'{count}개 행 선택됨' if count else '')
        if count:
            self.ndt_selection_label.grid()
        else:
            self.ndt_selection_label.grid_remove()

    def _select_filled_ndt_rows(self):
        self._clear_ndt_selection()
        first = None
        for index, row in enumerate(self.ndt_grid_entries):
            if any(str(widget.get()).strip() for widget in row.values()):
                self.ndt_selection_vars[index].set(True)
                if first is None:
                    first = row
        if first is not None:
            self._select_ndt_row(first)
        self._refresh_ndt_selection_text()

    def _select_ndt_range(self, index):
        anchor = getattr(self, '_ndt_selection_anchor', None)
        if anchor is None:
            anchor = index
        first, last = sorted((anchor, index))
        for selected in self.ndt_selection_vars[first:last + 1]:
            selected.set(True)
        self._select_ndt_row(self.ndt_grid_entries[index])
        self._ndt_selection_anchor = anchor
        self._refresh_ndt_selection_text()
        return 'break'

    def _toggle_ndt_row(self, row, selected):
        if selected.get():
            self._ndt_selection_anchor = next(i for i, var in enumerate(self.ndt_selection_vars) if var is selected)
            self._select_ndt_row(row)
        else:
            if self.selected_ndt_row is row:
                self.selected_ndt_row = None
                self.photo_selection_var.set('')
        self._refresh_ndt_selection_text()

    def _clear_ndt_selection(self, event=None):
        prev = getattr(self, 'selected_ndt_row', None)
        self.selected_ndt_row = None
        self._ndt_selection_anchor = None
        for var in self.ndt_selection_vars:
            var.set(False)
        if prev is not None and hasattr(prev, '_row_index'):
            self._update_ndt_row_tags(prev._row_index)
        self._refresh_ndt_selection_text()
        self.photo_selection_var.set("")

    def _default_photo_description(self, method):
        return {
            'PAUT': '위상배열 초음파탐상검사',
            'RT': '방사선투과검사',
            'MT': '자분탐상검사',
            'PT': '침투탐상검사',
        }.get(method, f'{method} 공정사진' if method else '공정사진')

    def add_process_photo(self):
        if self.selected_ndt_row is None:
            messagebox.showwarning('알림', '먼저 비파괴검사결과서의 행을 선택하세요.')
            return
        row = self.selected_ndt_row
        method = row['검사방법'].get().strip().upper()
        joint = row['Joint No.'].get().strip()
        if method not in {'PAUT', 'RT', 'MT', 'PT'}:
            messagebox.showwarning(
                '알림', '공정사진은 PAUT, RT, MT, PT 검사만 등록할 수 있습니다.'
            )
            return
        if not method or not joint:
            messagebox.showwarning('알림', '선택 행의 검사방법과 Joint No.를 입력하세요.')
            return

        files = filedialog.askopenfilenames(
            title='공정사진 선택',
            filetypes=[
                ('이미지 파일', '*.jpg *.jpeg *.png *.bmp *.tif *.tiff'),
                ('모든 파일', '*.*'),
            ],
        )
        if not files:
            return
        description = simpledialog.askstring(
            '사진 설명', '사진 설명:',
            initialvalue=self._default_photo_description(method), parent=self
        )
        if description is None:
            return

        current_date = self.date_entry.get()
        target_dir = os.path.join(self.photo_root, current_date)
        os.makedirs(target_dir, exist_ok=True)
        history = self.load_history()
        day_data = history.setdefault(current_date, {})
        photos = day_data.setdefault('process_photos', [])
        for source_path in files:
            photo_id = uuid.uuid4().hex
            extension = os.path.splitext(source_path)[1].lower() or '.jpg'
            filename = f'{method}_{photo_id}{extension}'
            target_path = os.path.join(target_dir, filename)
            shutil.copy2(source_path, target_path)
            photos.append({
                'id': photo_id,
                'process': method,
                'date': current_date,
                'section': row['구간'].get().strip(),
                'line_no': row['라인번호'].get().strip(),
                'joint_no': joint,
                'welder': self._normalize_welder_id(row['용접사'].get()),
                'location': (
                    row['구간'].get().strip()
                    or row['라인번호'].get().strip()
                    or joint
                ),
                'description': description.strip(),
                'file_path': os.path.relpath(target_path, os.path.dirname(self.history_path)),
            })
        self.save_history(history)
        messagebox.showinfo('완료', f'공정사진 {len(files)}장을 등록했습니다.')

    def _rt_context(self):
        return list(getattr(self, '_rt_parent_history', [])) + [widget_rt_values(row) for row in self.ndt_grid_entries]

    def _sync_rt_row(self, row):
        try:
            sync_rt_widgets(row, all_rows=self._rt_context())
        except ValueError:
            pass

    def _validate_rt_rows(self):
        try:
            context = self._rt_context()
            for row in self.ndt_grid_entries:
                if row['검사방법'].get().strip().upper() == 'RT':
                    rt_counts(resolve_rt_row(widget_rt_values(row), context))
            for row in self.ndt_grid_entries:
                sync_rt_widgets(row, all_rows=context)
            return True
        except ValueError as exc:
            messagebox.showwarning('RT 매수 확인', str(exc))
            return False

    def delete_current_date_data(self):
        """현재 선택된 날짜의 데이터를 삭제하고 화면을 초기화합니다."""
        date_str = self.date_entry.get()
        if not date_str: return

        history = self.load_history()
        day_data = history.get(date_str, {})
        saved_fields = {key: value for key, value in day_data.items() if key != 'process_photos'}
        if (getattr(self, '_loaded_log_date', None) == date_str
                and hasattr(self, '_saved_log_fields') and self._saved_log_fields != saved_fields):
            messagebox.showerror('작업일보 삭제 충돌',
                                 '다른 곳에서 이 날짜의 기록을 변경했습니다. 최신 내용을 다시 불러온 뒤 삭제해 주세요.',
                                 parent=self)
            return
        day_photos = day_data.get('process_photos', [])
        photo_count = len(day_photos)

        confirm_message = (
            f"정말로 {date_str}의 작업/감독일보 저장 내용을 모두 삭제하시겠습니까?\n"
            f"등록된 공정사진 {photo_count}장도 함께 삭제됩니다.\n\n"
            "※ 프로그램 관리 폴더의 복사본만 삭제되며 외부 원본은 유지됩니다.\n"
            "(이 작업은 되돌릴 수 없습니다)"
        )
        if not messagebox.askyesno("삭제 확인", confirm_message, parent=self):
            return

        # 다른 날짜에서 참조 중인 사진은 삭제하지 않는다.
        other_photo_paths = set()
        for other_date, other_data in history.items():
            if other_date == date_str:
                continue
            for photo in other_data.get('process_photos', []):
                file_path = photo.get('file_path', '')
                if file_path:
                    other_photo_paths.add(os.path.normcase(os.path.abspath(os.path.join(
                        os.path.dirname(self.history_path), file_path
                    ))))

        managed_root = os.path.normcase(os.path.abspath(self.photo_root))

        # 먼저 해당 날짜 기록을 저장소에서 제거한 뒤 관리용 사진을 정리한다.
        # 저장에 실패하면 사진 삭제를 시작하지 않는다.
        if date_str in history:
            self.save_history({day: data for day, data in history.items() if day != date_str})

        deleted_photo_count = 0
        failed_photos = []
        for photo in day_photos:
            file_path = photo.get('file_path', '')
            if not file_path:
                continue
            photo_path = os.path.normcase(os.path.abspath(os.path.join(
                os.path.dirname(self.history_path), file_path
            )))
            try:
                is_managed_photo = os.path.commonpath(
                    [managed_root, photo_path]
                ) == managed_root
            except ValueError:
                is_managed_photo = False

            if not is_managed_photo or photo_path in other_photo_paths:
                continue
            try:
                if os.path.isfile(photo_path):
                    os.remove(photo_path)
                    deleted_photo_count += 1
            except OSError as exc:
                failed_photos.append(f"{os.path.basename(photo_path)}: {exc}")

        # 화면 UI 초기화 (Clear all fields)
        self.weather_entry.delete(0, tk.END)
        
        for ents in self.qty_entries.values():
            for k in ['예상량', '전일누계', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                if ents.get(k):
                    ents[k].delete(0, tk.END)
                    
        for entries in self.equip_entries.values():
            for ent in entries.values():
                ent.delete(0, tk.END)

        for ent in self.personnel_entries.values():
            ent.delete(0, tk.END)

        self.remarks_text.delete('1.0', tk.END)

        for row in self.ndt_grid_entries:
            for ent in row.values():
                if isinstance(ent, ttk.Combobox):
                    ent.set('')
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)

        self._clear_ndt_selection()
        self._loaded_log_date = date_str
        self._saved_log_fields = {}
        self._log_baseline = self._capture_log_input()
            
        if hasattr(self, 'log'):
            self.log(f"[{date_str}] 작업일보 데이터가 삭제되었습니다.")
        result_message = (
            f"{date_str} 데이터가 삭제되었습니다.\n"
            f"관리용 공정사진 {deleted_photo_count}장을 삭제했습니다."
        )
        if failed_photos:
            result_message += (
                f"\n\n삭제하지 못한 사진 {len(failed_photos)}장:\n"
                + "\n".join(failed_photos)
            )
            messagebox.showwarning("일부 삭제 완료", result_message, parent=self)
        else:
            messagebox.showinfo("삭제 완료", result_message, parent=self)

    def manage_process_photos(self):
        current_date = self.date_entry.get()
        history = self.load_history()
        photos = history.get(current_date, {}).get('process_photos', [])
        if not photos:
            messagebox.showinfo('공정사진', '현재 날짜에 등록된 공정사진이 없습니다.')
            return

        window = tk.Toplevel(self)
        window.title(f'{current_date} 공정사진 관리')
        window.geometry('760x330')
        columns = ('공정', '위치', 'Joint', '용접사', '설명', '파일')
        tree = ttk.Treeview(window, columns=columns, show='headings', selectmode='browse')
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=90 if column != '설명' else 220)
        tree.pack(fill='both', expand=True, padx=8, pady=8)
        for index, photo in enumerate(photos):
            tree.insert('', 'end', iid=str(index), values=(
                photo.get('process', ''), photo.get('location', ''),
                photo.get('joint_no', ''), photo.get('welder', ''),
                photo.get('description', ''), os.path.basename(photo.get('file_path', '')),
            ))

        def selected_index():
            selection = tree.selection()
            return int(selection[0]) if selection else None

        def open_photo(event=None):
            index = selected_index()
            if index is None:
                return
            path = os.path.abspath(os.path.join(
                os.path.dirname(self.history_path), photos[index].get('file_path', '')
            ))
            if os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showerror('오류', f'사진 파일을 찾을 수 없습니다.\n{path}')

        def delete_photo(event=None):
            index = selected_index()
            if index is None or not messagebox.askyesno('삭제', '선택한 사진을 삭제할까요?', parent=window):
                return
            photo = photos.pop(index)
            path = os.path.abspath(os.path.join(
                os.path.dirname(self.history_path), photo.get('file_path', '')
            ))
            managed_root = os.path.abspath(self.photo_root)
            try:
                is_managed_photo = (
                    os.path.commonpath([managed_root, path]) == managed_root
                )
            except ValueError:
                is_managed_photo = False
            if is_managed_photo and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception as e:
                    messagebox.showwarning('경고', f'파일 삭제 중 오류가 발생했습니다: {e}', parent=window)
            history[current_date]['process_photos'] = photos
            self.save_history(history)
            window.destroy()
            self.manage_process_photos()

        # Context menu & bindings
        context_menu = tk.Menu(window, tearoff=0)
        context_menu.add_command(label="사진 열기", command=open_photo)
        context_menu.add_command(label="선택 삭제", command=delete_photo)

        def show_context_menu(event):
            item = tree.identify_row(event.y)
            if item:
                tree.selection_set(item)
                context_menu.tk_popup(event.x_root, event.y_root)

        tree.bind('<Double-1>', open_photo)
        tree.bind('<Delete>', delete_photo)
        tree.bind('<Button-3>', show_context_menu)

        button_bar = ttk.Frame(window)
        button_bar.pack(fill='x', padx=8, pady=(0, 8))
        ttk.Button(button_bar, text='사진 열기', command=open_photo).pack(side='left', padx=4)
        ttk.Button(button_bar, text='선택 삭제', command=delete_photo).pack(side='left', padx=4)
        ttk.Button(button_bar, text='닫기', command=window.destroy).pack(side='right', padx=4)
            
    def export_excel(self):
        if not self._validate_rt_rows():
            return
        if not self._validate_hot_rows():
            return
        # Save File Dialog
        default_name = f"{self.date_entry.get().replace('-', '')}_작업_감독일보.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            title="일보 저장",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return
        self.auto_calculate_and_save(show_completion=False)
        # Gather all data
        data = {
            'date': self.date_entry.get(),
            'weather': self.weather_entry.get(),
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': []
        }

        # Qty
        for comp_key, entries in self.qty_entries.items():
            data['qty_data'][comp_key] = {k: v.get() for k, v in entries.items()}

        # Equip
        for eq, entries in self.equip_entries.items():
            data['equip_data'][eq] = {k: v.get() for k, v in entries.items()}

        # Personnel
        for p_key, ent in self.personnel_entries.items():
            data['personnel_data'][p_key] = ent.get()

        # NDT - only gather rows that have at least one non-empty value
        for row_entries in self.ndt_grid_entries:
            row_dict = {col: ent.get() for col, ent in row_entries.items()}
            row_dict['용접사'] = self._normalize_welder_id(row_dict.get('용접사'))
            if any(val.strip() for val in row_dict.values()):
                data['ndt_results'].append(row_dict)

        self.save_current_history()
        if file_path:
            try:
                exporter = DailyWorkLogExporter()
                exporter.generate_excel(data, file_path)
                messagebox.showinfo("성공", f"엑셀 파일이 생성되었습니다!\n{file_path}")
            except Exception as e:
                messagebox.showerror("오류", f"엑셀 출력 중 오류가 발생했습니다:\n{str(e)}")

    def export_ndt_summary(self):
        try:
            from site_apps.central.src.ndt_summary_exporter import NDTSummaryExporter
        except ImportError:
            NDTSummaryExporter = None

        try:
            with open(self.history_path, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except Exception as e:
            messagebox.showerror("오류", f"데이터 파일을 읽는 데 실패했습니다:\n{str(e)}")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="NDT 누계 대장 저장",
            initialfile="NDT_누계대장.xlsx"
        )
        if not file_path:
            return
            
        try:
            if NDTSummaryExporter is None:
                raise ImportError("NDTSummaryExporter 모듈을 불러올 수 없습니다.")
            exporter = NDTSummaryExporter(history)
            exporter.generate(file_path)
            messagebox.showinfo("완료", "NDT 누계 대장 엑셀 파일이 성공적으로 생성되었습니다.")
            os.startfile(os.path.dirname(file_path))
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"NDT 누계 대장 생성 중 오류가 발생했습니다:\n{str(e)}")

    def export_monthly_report(self):
        # 1. 설정 다이얼로그
        top = tk.Toplevel(self)
        top.title("월간진도보고서 엑셀 출력")
        # Allow enough vertical space for all fields and the generate button.
        top.geometry("620x470")
        top.minsize(620, 470)
        top.transient(self)
        top.grab_set()
        
        ttk.Label(top, text="📊 월간진도보고서 자동 생성", font=('맑은 고딕', 12, 'bold')).pack(pady=10)
        
        # 년/월 선택
        f_date = ttk.Frame(top)
        f_date.pack(pady=10)
        
        now = datetime.now()
        ttk.Label(f_date, text="누적 종료 연도:").pack(side=tk.LEFT, padx=5)
        year_var = tk.StringVar(value="2027")
        ttk.Spinbox(f_date, from_=2020, to=2030, textvariable=year_var, width=6).pack(side=tk.LEFT)
        
        ttk.Label(f_date, text="종료 월:").pack(side=tk.LEFT, padx=5)
        month_var = tk.StringVar(value="08")
        ttk.Spinbox(f_date, from_=1, to=12, textvariable=month_var, width=4, format="%02.0f").pack(side=tk.LEFT)

        # 집계 범위 선택
        f_scope = ttk.LabelFrame(top, text="집계 범위", padding=(10, 5))
        f_scope.pack(pady=5)
        scope_var = tk.StringVar(value="month")
        ttk.Radiobutton(
            f_scope, text="해당 월만", variable=scope_var, value="month"
        ).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(
            f_scope, text="전체 누적", variable=scope_var, value="cumulative"
        ).pack(side=tk.LEFT, padx=10)
        
        # 문서번호 입력
        f_doc = ttk.Frame(top)
        f_doc.pack(pady=5)
        ttk.Label(f_doc, text="문서번호:").pack(side=tk.LEFT, padx=5)
        doc_var = tk.StringVar(value="01")
        ttk.Entry(f_doc, textvariable=doc_var, width=10).pack(side=tk.LEFT)
        
        # 작성일자 입력
        f_create_date = ttk.Frame(top)
        f_create_date.pack(pady=5)
        ttk.Label(f_create_date, text="작성일자:").pack(side=tk.LEFT, padx=5)
        create_date_var = tk.StringVar(value=now.strftime("%Y. %m. %d."))
        ttk.Entry(f_create_date, textvariable=create_date_var, width=15).pack(side=tk.LEFT)
        
        # 템플릿 파일 선택 (기본값)
        f_tmpl = ttk.Frame(top)
        f_tmpl.pack(pady=10, fill=tk.X, padx=10)
        ttk.Label(f_tmpl, text="템플릿:").pack(side=tk.LEFT)
        
        tmpl_var = tk.StringVar(value=os.path.join(
            os.path.expanduser("~"), "Desktop", "템플릿_최종완성본_V70.xlsx"
        ))
        ttk.Entry(f_tmpl, textvariable=tmpl_var, width=35).pack(side=tk.LEFT, padx=5)
        
        def browse_tmpl():
            path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")], title="템플릿 선택")
            if path: tmpl_var.set(path)
        ttk.Button(f_tmpl, text="...", width=3, command=browse_tmpl).pack(side=tk.LEFT)

        # 월별 발주자 보고자료 선택 (선택하지 않으면 16페이지는 빈 표 유지)
        f_owner = ttk.Frame(top)
        f_owner.pack(pady=5, fill=tk.X, padx=10)
        ttk.Label(f_owner, text="발주자 보고:").pack(side=tk.LEFT)
        owner_report_var = tk.StringVar()
        ttk.Entry(
            f_owner, textvariable=owner_report_var, width=43
        ).pack(side=tk.LEFT, padx=5)

        def browse_owner_report():
            path = filedialog.askopenfilename(
                filetypes=[
                    ("Excel", "*.xlsx *.xls"),
                    ("Excel 통합 문서", "*.xlsx"),
                    ("Excel 97-2003", "*.xls"),
                ],
                title="발주자 보고 파일 선택",
            )
            if path:
                owner_report_var.set(path)

        ttk.Button(
            f_owner, text="...", width=3, command=browse_owner_report
        ).pack(side=tk.LEFT)
        
        messages = Queue()
        generation_status = tk.StringVar(master=top, value="생성 대기")
        ttk.Label(top, textvariable=generation_status).pack(anchor="w", padx=14, pady=8)

        def set_generation_stage(stage):
            self.monthly_report_active_stage = stage
            generation_status.set(stage)

        def close_dialog():
            if not getattr(self, "monthly_report_active_stage", None):
                top.destroy()

        top.protocol("WM_DELETE_WINDOW", close_dialog)

        def do_generate():
            if getattr(self, "monthly_report_active_stage", None):
                return
            ym = f"{year_var.get()}-{month_var.get().zfill(2)}"
            tmpl_path = tmpl_var.get()
            owner_report_path = owner_report_var.get().strip()
            scope = scope_var.get()
            doc_num = doc_var.get().strip()
            create_date = create_date_var.get().strip()
            history_path = self.history_path
            
            if not os.path.exists(tmpl_path):
                messagebox.showerror("오류", f"템플릿 파일을 찾을 수 없습니다.\n{tmpl_path}")
                return
            if owner_report_path and not os.path.exists(owner_report_path):
                messagebox.showerror(
                    "오류",
                    f"발주자 보고 파일을 찾을 수 없습니다.\n{owner_report_path}",
                )
                return
                
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=(
                    f"월간진도보고서_{ym.replace('-', '')}_"
                    f"{'월간' if scope == 'month' else '전체누적'}.xlsx"
                ),
                title="저장 위치 선택",
                filetypes=[("Excel", "*.xlsx")]
            )
            if not save_path: return
            
            def generate():
                try:
                    result = MonthlyReportManager(tmpl_path).generate_report(
                        history_path, ym, save_path, doc_num=doc_num,
                        create_date=create_date, owner_report_path=owner_report_path or None,
                        report_scope=scope,
                        progress=lambda stage: messages.put(("stage", stage)),
                    )
                    messages.put(("result", result))
                except Exception as error:
                    messages.put(("error", str(error)))

            def poll_progress():
                while True:
                    try:
                        kind, value = messages.get_nowait()
                    except Empty:
                        top.after(50, poll_progress)
                        return
                    if kind == "stage":
                        set_generation_stage(value)
                        continue
                    self.monthly_report_active_stage = None
                    generate_button.config(state="normal")
                    if kind == "error":
                        generation_status.set("생성 실패")
                        messagebox.showerror("오류", f"생성 중 오류 발생:\n{value}", parent=top)
                    elif value:
                        generation_status.set("저장 완료")
                        messagebox.showinfo("성공", f"월간진도보고서가 생성되었습니다.\n{value}", parent=top)
                        try:
                            os.startfile(value)
                        except OSError as error:
                            messagebox.showwarning("파일 열기", f"저장은 완료됐지만 파일을 열지 못했습니다.\n{error}", parent=top)
                        top.destroy()
                    else:
                        generation_status.set("대상 작업일보 없음")
                        period_text = f"{ym}의" if scope == "month" else f"{ym}까지의"
                        messagebox.showwarning("알림", f"{period_text} 작업일보 데이터가 없습니다.", parent=top)
                    return

            generate_button.config(state="disabled")
            set_generation_stage("생성 준비 중…")
            try:
                Thread(target=generate, name="central-monthly-report", daemon=False).start()
            except RuntimeError as error:
                messages.put(("error", str(error)))
            top.after(50, poll_progress)
                
        generate_button = ttk.Button(top, text="보고서 생성", command=do_generate)
        generate_button.pack(side=tk.BOTTOM, pady=18, ipadx=24, ipady=6)

    def load_history(self):
        try:
            import os
            current_mtime = os.path.getmtime(self.history_path) if os.path.exists(self.history_path) else None
            if (getattr(self, '_history_cache', None) is not None and 
                getattr(self, '_history_cache_mtime', None) is not None and 
                current_mtime == self._history_cache_mtime):
                return self._history_cache

            with open(self.history_path, 'r', encoding='utf-8') as f:
                history = json.load(f)
            if not isinstance(history, dict) or any(not isinstance(day, dict) for day in history.values()):
                raise ValueError('일보 이력 형식이 올바르지 않습니다.')
            self._rt_parent_history = [row for day in sorted(history) for row in history[day].get('ndt_results', [])]
            self._history_cache = history
            self._history_cache_mtime = current_mtime
            return history
        except FileNotFoundError:
            self._history_cache = {}
            self._history_cache_mtime = None
            return {}
        except (OSError, ValueError) as exc:
            messagebox.showerror('일보 이력 읽기 실패', f'이력을 읽지 못해 작업을 중단합니다. 원본 파일은 유지됩니다.\n{exc}')
            raise

    def save_history(self, history):
        connection_error = get_data_connection_error()
        if connection_error:
            messagebox.showerror(
                "작업일보 저장 중단",
                connection_error,
                parent=self.winfo_toplevel(),
            )
            raise SharedDataUnavailableError(connection_error)

        expected_mtime = getattr(self, '_history_cache_mtime', None)
        temp_path = self.history_path + f'.{uuid.uuid4().hex}.tmp'

        def verify_version():
            try:
                current_mtime = os.path.getmtime(self.history_path)
            except FileNotFoundError:
                current_mtime = None
            if current_mtime != expected_mtime:
                self._history_cache = None
                self._history_cache_mtime = None
                message = '다른 곳에서 작업일보를 변경했습니다. 작성 내용은 화면에 남아 있습니다.'
                messagebox.showerror('작업일보 저장 충돌', message, parent=self)
                raise OSError(message)

        try:
            verify_version()
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=4)
            if expected_mtime is not None:
                shutil.copy2(self.history_path, self.history_path + '.backup.json')
            verify_version()
            os.replace(temp_path, self.history_path)
        finally:
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        self._history_cache = history
        self._history_cache_mtime = os.path.getmtime(self.history_path) if os.path.exists(self.history_path) else None
        self._corrected_history_source = None
        self._corrected_history_cache = None
        self._refresh_ndt_history_choices(history)
        self._refresh_saved_date_markers(history)
        self._update_date_status(history)
        # Saving photo metadata does not persist other editor fields.

    def _refresh_ndt_history_choices(self, history=None):
        """Refresh NDT combobox choices immediately after saving new rows."""
        if not hasattr(self, 'ndt_grid_entries'):
            return
        self._extract_history_choices(history)

        for row_entries in self.ndt_grid_entries:
            row_entries['구간'].configure(values=self.history_sections)
            row_entries['라인번호'].configure(values=self.history_lines)
            row_entries['업체'].configure(values=self.history_companies)
            row_entries['용접사'].configure(values=self.history_welders)

    def _refresh_saved_date_markers(self, history=None):
        """달력 팝업에서 작업일보가 저장된 날짜를 녹색으로 강조한다."""
        preview = getattr(self, 'calendar_work_summary', None)
        if preview is not None:
            preview.refresh(history=history)
            return
        if not hasattr(self, 'date_entry'):
            return
        top_cal = getattr(self.date_entry, '_top_cal', None)
        if top_cal is not None and not top_cal.winfo_ismapped():
            return
        summary = self._ensure_calendar_work_summary()
        if summary is not None:
            summary.refresh(history=history)

    def _update_date_status(self, history=None):
        """현재 선택한 날짜의 저장 여부를 날짜 입력칸 아래에 표시한다."""
        if history is None:
            history = self.load_history()
        current_date = self._selected_date_key()
        if current_date in history:
            self.date_status_label.configure(
                text='● 저장된 날짜', fg='#1B5E20'
            )
        else:
            self.date_status_label.configure(
                text='○ 저장되지 않은 날짜', fg='#777777'
            )

    def _selected_date_key(self):
        """DateEntry 값을 작업일보 저장 키(YYYY-MM-DD)로 통일한다."""
        try:
            return self.date_entry.get_date().strftime('%Y-%m-%d')
        except (AttributeError, TypeError, ValueError):
            value = str(self.date_entry.get() or '').strip()
            for date_format in ('%Y-%m-%d', '%Y.%m.%d', '%Y/%m/%d'):
                try:
                    return datetime.strptime(value, date_format).strftime('%Y-%m-%d')
                except ValueError:
                    continue
            return value

    def save_current_history(self):
        if not self._validate_rt_rows():
            return
        history = self.load_history()
        current_date = self._selected_date_key()
        saved_fields = {key: value for key, value in history.get(current_date, {}).items()
                        if key != 'process_photos'}
        if (getattr(self, '_loaded_log_date', None) == current_date
                and hasattr(self, '_saved_log_fields') and self._saved_log_fields != saved_fields):
            message = '작성 중인 날짜가 다른 곳에서 변경되었습니다. 작성 내용을 보존한 뒤 최신 기록을 다시 확인해 주세요.'
            messagebox.showerror('작업일보 저장 충돌', message, parent=self)
            raise OSError(message)
        if not self._validate_hot_rows():
            return
        existing_photos = history.get(current_date, {}).get('process_photos', [])
        
        data = {
            'weather': self.weather_entry.get(),
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': [],
            'process_photos': existing_photos,
        }
        for comp_key, entries in self.qty_entries.items():
            data['qty_data'][comp_key] = {k: v.get() for k, v in entries.items()}
        for eq, entries in self.equip_entries.items():
            data['equip_data'][eq] = {k: v.get() for k, v in entries.items()}
        for p_key, ent in self.personnel_entries.items():
            data['personnel_data'][p_key] = ent.get()
            
        for row_entries in self.ndt_grid_entries:
            row_dict = {}
            for col, ent in row_entries.items():
                if hasattr(ent, 'get'):
                    val = ent.get()
                    if callable(val): # In case get() returned a method? No, ent.get is the method.
                        pass
                    if col == '용접사':
                        val = self._normalize_welder_id(val)
                    row_dict[col] = str(val) if val else ""
            row_dict.update(getattr(row_entries.get('구간정보'), '_rt_retake_meta', {}))
            data['ndt_results'].append(row_dict)
            
        from copy import deepcopy
        self.save_history({**history, current_date: data})
        self._loaded_log_date = current_date
        self._saved_log_fields = deepcopy({key: value for key, value in data.items() if key != 'process_photos'})
        self._log_baseline = self._capture_log_input()

    def _capture_log_input(self):
        """Capture editable values without saving or including the date selector."""
        return {
            'weather': self.weather_entry.get(),
            'remarks': self.remarks_text.get('1.0', 'end-1c'),
            'qty': {key: {field: widget.get() for field, widget in entries.items()}
                    for key, entries in self.qty_entries.items()},
            'equipment': {key: {field: widget.get() for field, widget in entries.items()}
                          for key, entries in self.equip_entries.items()},
            'people': {key: widget.get() for key, widget in self.personnel_entries.items()},
            'ndt': [{key: widget.get() for key, widget in row.items() if hasattr(widget, 'get')}
                    for row in self.ndt_grid_entries],
        }

    def _history_for_date_view(self, history):
        if getattr(self, '_corrected_history_source', None) is not history:
            self._corrected_history_cache = corrected_rt_history(history)
            self._corrected_history_source = history
        return self._corrected_history_cache

    def on_date_change(self, event=None):
        if hasattr(self, '_commit_active_cell_editor'):
            self._commit_active_cell_editor()
            self._hide_active_cell_editor()
        current_date = self._selected_date_key()
        previous = getattr(self, '_loaded_log_date', None)
        baseline = getattr(self, '_log_baseline', None)
        if baseline is not None and self._capture_log_input() != baseline:
            if not messagebox.askyesno('일보 날짜 이동', '저장하지 않은 작성 내용을 버리고 선택한 날짜를 불러올까요?', parent=self):
                if previous:
                    self.date_entry.set_date(previous)
                return
        # Read before clearing the editor so failures preserve the draft.
        try:
            loaded_history = self.load_history()
        except (OSError, ValueError):
            if previous:
                self.date_entry.set_date(previous)
            return
        self._clear_ndt_selection()
        # 화면에서도 RT 불량과 불량률을 검사수량 누계와 같은 기준으로 표시한다.
        history = self._history_for_date_view(loaded_history)
        self._refresh_saved_date_markers(history)
        self._update_date_status(history)
        
        past_dates = [d for d in history.keys() if d < current_date]
        if past_dates:
            latest_past_date = max(past_dates)
            prev_data = history[latest_past_date]
        else:
            prev_data = {'qty_data': {}, 'equip_data': {}, 'personnel_data': {}}
            
        curr_data = history.get(current_date, {})
        has_current_data = current_date in history
            
        # Update Qty
        def load_format_val(ckey, field_name, v):
            if not str(v).strip(): return v
            if field_name == '예상량':
                try: return f"{int(float(str(v).replace(',','')))}"
                except ValueError: return v
            if field_name in ['전일누계', '금일작업', '총누계']:
                try:
                    fv = float(str(v).replace(',', ''))
                    if ckey.startswith(('PAUT', 'MT', 'PT')): return f"{fv:.4f}"
                    return f"{fv:.1f}" if fv % 1 else f"{int(fv)}"
                except ValueError: return v
            return v

        def is_zero_number(value):
            try:
                return float(str(value).replace(',', '').strip() or 0) == 0
            except (TypeError, ValueError):
                return False

        for comp_key, entries in self.qty_entries.items():
            # Load from today if exists
            curr_qty = curr_data.get('qty_data', {}).get(comp_key, {})
            prev_qty = prev_data.get('qty_data', {}).get(comp_key, {})
            for field in ['예상량', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                entries[field].delete(0, tk.END)
                current_value = curr_qty.get(field, '')
                
                # [FIX] Force expected volume (예상량) to ALWAYS load from updated defaults
                if field == '예상량':
                    method, spec = comp_key.split('_', 1)
                    current_value = self.default_qty.get((method, spec), '')
                    
                if not has_current_data:
                    if field == '금일작업':
                        current_value = ''
                    elif field == '총누계':
                        current_value = prev_qty.get('총누계', '0') or '0'
                    elif field in ['불량', '불량률', '비고']:
                        current_value = prev_qty.get(field, '')
                    elif field == '공정률':
                        previous_total = prev_qty.get('총누계', '0') or '0'
                        method, spec = comp_key.split('_', 1)
                        expected_value = self.default_qty.get((method, spec), '')
                        try:
                            total_num = float(str(previous_total).replace(',', ''))
                            expected_num = float(str(expected_value).replace(',', ''))
                            current_value = f"{(total_num / expected_num) * 100:.1f}" if expected_num > 0 else ''
                        except (TypeError, ValueError):
                            current_value = ''
                else:
                    if field == '공정률':
                        current_total = curr_qty.get('총누계', '0') or '0'
                        method, spec = comp_key.split('_', 1)
                        expected_value = self.default_qty.get((method, spec), '')
                        try:
                            total_num = float(str(current_total).replace(',', ''))
                            expected_num = float(str(expected_value).replace(',', ''))
                            current_value = f"{(total_num / expected_num) * 100:.1f}" if expected_num > 0 else ''
                        except (TypeError, ValueError):
                            current_value = ''
                effective_total = curr_qty.get('총누계', '') if has_current_data else prev_qty.get('총누계', '0')
                if field in ['금일작업', '총누계', '불량', '불량률'] and is_zero_number(current_value):
                    current_value = ''
                elif field == '공정률' and is_zero_number(effective_total):
                    current_value = ''
                if str(current_value).strip():
                    val = load_format_val(comp_key, field, current_value)
                    entries[field].insert(0, val)
                    
            # Always calculate 전일누계 from past
            prev_total = prev_qty.get('총누계', '0') or '0'
            entries['전일누계'].delete(0, tk.END)
            if not is_zero_number(prev_total):
                entries['전일누계'].insert(0, load_format_val(comp_key, '전일누계', prev_total))
            
        # Update Equip
        for eq, entries in self.equip_entries.items():
            curr_eq = curr_data.get('equip_data', {}).get(eq, {})
            prev_eq = prev_data.get('equip_data', {}).get(eq, {})
            for field in ['금일', '누계']:
                entries[field].delete(0, tk.END)
                if has_current_data:
                    value = curr_eq.get(field, '')
                elif field == '금일':
                    value = ''
                else:
                    value = prev_eq.get('누계', '0') or '0'
                if str(value).strip():
                    entries[field].insert(0, value)
                    
        # Update Personnel
        for p_key, ent in self.personnel_entries.items():
            ent.delete(0, tk.END)
            if has_current_data:
                value = curr_data.get('personnel_data', {}).get(p_key, '')
            elif p_key.endswith('_누계'):
                value = prev_data.get('personnel_data', {}).get(p_key, '0') or '0'
            else:
                value = ''
            if str(value).strip():
                ent.insert(0, value)
                
        # Update Weather
        self.weather_entry.delete(0, tk.END)
        if 'weather' in curr_data:
            self.weather_entry.insert(0, curr_data['weather'])
            
        # Update Remarks
        self.remarks_text.delete("1.0", tk.END)
        if 'remarks' in curr_data:
            self.remarks_text.insert("1.0", curr_data['remarks'])
            
        # Update NDT Results
        curr_ndt = curr_data.get('ndt_results', [])
        self._ensure_ndt_capacity(len(curr_ndt))

        def split_legacy_spec(row_data):
            spec = str(row_data.get('규격', '') or '').strip()
            shift = str(row_data.get('근무구분', '') or '').strip()
            if not shift:
                for legacy_shift in ('주간', '야간', '재검'):
                    if spec.endswith(legacy_shift):
                        spec = spec[:-len(legacy_shift)].strip()
                        shift = legacy_shift
                        break
            if spec.startswith('31/3'):
                spec = '3 1/3' + spec[len('31/3'):]
            spec = spec.replace(' X12"', ' X 12"').replace(' X6"', ' X 6"')
            return spec, shift

        for i, row_entries in enumerate(self.ndt_grid_entries):
            row_data = curr_ndt[i] if i < len(curr_ndt) else {}
            normalized_spec, work_shift = split_legacy_spec(row_data)
            if hasattr(row_entries, 'begin_batch'):
                row_entries.begin_batch()
            try:
                for col, ent in row_entries.items():
                    if col == '규격':
                        val = normalized_spec
                    elif col == '근무구분':
                        val = work_shift
                    elif col == '용접사':
                        val = self._format_welder_display(row_data.get(col, ''))
                    else:
                        val = row_data.get(col, '')
                    if hasattr(ent, 'set'):
                        ent.set(val)
                    elif hasattr(ent, 'delete'):
                        ent.delete(0, tk.END)
                        ent.insert(0, val)
                if '구간정보' in row_entries:
                    row_entries['구간정보']._rt_retake_meta = {
                        k: row_data[k] for k in ('원본JointNo', '재촬영구간') if k in row_data
                    }
            finally:
                if hasattr(row_entries, 'end_batch'):
                    row_entries.end_batch()
                    
        rt_context = None
        for widgets in self.ndt_grid_entries:
            self._sync_ndt_spec(widgets)
            if '공사구분' in widgets and '업체' in widgets:
                comp = widgets['업체'].get().strip()
                cat_cb = widgets['공사구분']
                matched_opts = None
                for k, opts in COMPANY_WORK_CATEGORIES.items():
                    if k in comp:
                        matched_opts = opts
                        break
                cat_cb['values'] = matched_opts if matched_opts else ALL_WORK_CATEGORIES
            if widgets['검사방법'].get().strip().upper() == 'RT':
                if rt_context is None:
                    rt_context = self._rt_context()
                try:
                    sync_rt_widgets(widgets, all_rows=rt_context)
                except ValueError:
                    pass  # Keep the saved inputs available for correction before saving.
        self._refresh_hot_summary()

        # [NEW] Recalculate Subtotals for Qty to heal missing '금일작업' fields in UI
        def format_val(ckey, v):
            if ckey.startswith(('PAUT', 'MT', 'PT')):
                return f"{v:.4f}"
            return f"{v:.1f}" if v % 1 else f"{int(v)}"
            
        paut_keys = ['300A이상', '300A이상-야간', '250A', '250A-야간', '200A', '200A-야간', '150A~125A', '150A~125A-야간', '100A이하', '100A이하-야간']
        paut_expected = sum(float(self.qty_entries[f"PAUT_{s}"]['예상량'].get() or 0) for s in paut_keys)
        paut_prev = sum(float(self.qty_entries[f"PAUT_{s}"]['전일누계'].get() or 0) for s in paut_keys)
        paut_today = sum(float(self.qty_entries[f"PAUT_{s}"]['금일작업'].get() or 0) for s in paut_keys)
        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in paut_keys)
        
        self.qty_entries['PAUT_소계']['예상량'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['전일누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['금일작업'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['공정률'].delete(0, tk.END)
        
        if paut_expected > 0:
            self.qty_entries['PAUT_소계']['예상량'].insert(0, f"{int(paut_expected)}")
        if paut_prev != 0:
            self.qty_entries['PAUT_소계']['전일누계'].insert(0, format_val('PAUT', paut_prev))
        if paut_today != 0:
            self.qty_entries['PAUT_소계']['금일작업'].insert(0, format_val('PAUT', paut_today))
        if paut_total != 0:
            self.qty_entries['PAUT_소계']['총누계'].insert(0, format_val('PAUT', paut_total))
            if paut_expected > 0:
                self.qty_entries['PAUT_소계']['공정률'].insert(0, f"{(paut_total / paut_expected) * 100:.1f}")
                
        rt_keys = ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간']
        rt_expected = sum(float(self.qty_entries[f"RT_{s}"]['예상량'].get() or 0) for s in rt_keys)
        rt_prev = sum(float(self.qty_entries[f"RT_{s}"]['전일누계'].get() or 0) for s in rt_keys)
        rt_today = sum(float(self.qty_entries[f"RT_{s}"]['금일작업'].get() or 0) for s in rt_keys)
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in rt_keys)
        
        self.qty_entries['RT_소계']['예상량'].delete(0, tk.END)
        self.qty_entries['RT_소계']['전일누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['금일작업'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['공정률'].delete(0, tk.END)
        
        if rt_expected > 0:
            self.qty_entries['RT_소계']['예상량'].insert(0, f"{int(rt_expected)}")
        if rt_prev != 0:
            self.qty_entries['RT_소계']['전일누계'].insert(0, format_val('RT', rt_prev))
        if rt_today != 0:
            self.qty_entries['RT_소계']['금일작업'].insert(0, format_val('RT', rt_today))
        if rt_total != 0:
            self.qty_entries['RT_소계']['총누계'].insert(0, format_val('RT', rt_total))
            if rt_expected > 0:
                self.qty_entries['RT_소계']['공정률'].insert(0, f"{(rt_total / rt_expected) * 100:.1f}")

        from copy import deepcopy
        self._loaded_log_date = current_date
        self._saved_log_fields = deepcopy({key: value for key, value in loaded_history.get(current_date, {}).items()
                                          if key != 'process_photos'})
        self._log_baseline = self._capture_log_input()

    def auto_calculate_and_save(self, show_completion=True):
        if hasattr(self, '_commit_active_cell_editor'):
            self._commit_active_cell_editor()
        if not self._validate_rt_rows():
            return
        if not self._validate_hot_rows():
            return
        self._refresh_hot_summary()
        current_date = self.date_entry.get()
        history = corrected_rt_history(self.load_history())
        
        # 1. Aggregate NDT Results -> 금일작업
        today_qty = {comp_key: 0.0 for comp_key in self.qty_entries.keys()}
        today_rejects = {comp_key: 0 for comp_key in self.qty_entries.keys()}
        
        for row_entries in self.ndt_grid_entries:
            if not hasattr(row_entries['검사방법'], 'get'): continue
            method = row_entries['검사방법'].get().upper().strip()
            if not method: continue
            
            size_str = row_entries['관경'].get().strip() if hasattr(row_entries['관경'], 'get') else ""
            shift_str = row_entries['근무구분'].get().strip() if hasattr(row_entries['근무구분'], 'get') else ""
            
            import re
            size_match = re.search(r'\d+', size_str)
            size_val = int(size_match.group()) if size_match else 0
            
            if method == 'RT':
                val, rejects = rt_counts({key: row_entries[key].get() for key in ('RT_OR', 'RT_RE', '구간정보')})
                if val > 0:
                    spec_key = '80A이하' if size_val <= 80 else '150A~100A'
                    if shift_str == '야간': spec_key += '-야간'
                    comp = f"RT_{spec_key}"
                    if comp in today_qty:
                        today_qty[comp] += val
                        today_rejects[comp] += rejects
            
            elif method == 'PAUT':
                val = float(row_entries['PAUT'].get() or 0)
                if val > 0:
                    spec_key = '200A'
                    if size_val >= 300: spec_key = '300A이상'
                    elif size_val == 250: spec_key = '250A'
                    elif size_val == 200: spec_key = '200A'
                    elif 125 <= size_val <= 150: spec_key = '150A~125A'
                    elif size_val <= 100: spec_key = '100A이하'
                    if shift_str == '야간':
                        spec_key += '-야간'
                    comp = f"PAUT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
                        
            elif method == 'MT':
                val = float(row_entries['MT'].get() or 0)
                if val > 0:
                    spec_key = '전체(야간)' if shift_str == '야간' else '전체(주간)'
                    comp = f"MT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
                    
            elif method == 'PT':
                val = float(row_entries['PT'].get() or 0)
                if val > 0:
                    spec_key = '전체(야간)' if shift_str == '야간' else '전체(주간)'
                    comp = f"PT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
                    
        # Update 금일작업 in UI
        def format_val(ckey, v):
            if ckey.startswith(('PAUT', 'MT', 'PT')):
                return f"{v:.4f}"
            return f"{v:.1f}" if v % 1 else f"{int(v)}"

        for comp_key, val in today_qty.items():
            if '소계' not in comp_key:
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                if val != 0:
                    ent.insert(0, format_val(comp_key, val))
                
        # 2. Update Totals based on previous date
        past_dates = [d for d in history.keys() if d < current_date]
        if past_dates:
            latest_past_date = max(past_dates)
            prev_data = history[latest_past_date]
        else:
            prev_data = {'qty_data': {}, 'equip_data': {}, 'personnel_data': {}}
            
        # Qty
        for comp_key, entries in self.qty_entries.items():
            if '소계' in comp_key: continue
            prev_total_str = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0') or '0'
            try: prev_total = float(prev_total_str.replace(',',''))
            except ValueError: prev_total = 0.0
            today_val = float(entries['금일작업'].get() or 0)
            
            entries['금일작업'].delete(0, tk.END)
            if today_val != 0:
                entries['금일작업'].insert(0, format_val(comp_key, today_val))
            
            entries['전일누계'].delete(0, tk.END)
            if prev_total != 0:
                entries['전일누계'].insert(0, format_val(comp_key, prev_total))
            
            total = prev_total + today_val
            entries['총누계'].delete(0, tk.END)
            if total != 0:
                entries['총누계'].insert(0, format_val(comp_key, total))
            
            expected = float(entries['예상량'].get() or 0)
            entries['공정률'].delete(0, tk.END)
            if expected > 0 and total != 0:
                entries['예상량'].delete(0, tk.END)
                entries['예상량'].insert(0, f"{int(expected)}")
                progress = (total / expected) * 100
                entries['공정률'].insert(0, f"{progress:.1f}")

        # Subtotals for Qty
        paut_keys = ['300A이상', '300A이상-야간', '250A', '250A-야간', '200A', '200A-야간', '150A~125A', '150A~125A-야간', '100A이하', '100A이하-야간']
        paut_expected = sum(float(self.qty_entries[f"PAUT_{s}"]['예상량'].get() or 0) for s in paut_keys)
        paut_prev = sum(float(self.qty_entries[f"PAUT_{s}"]['전일누계'].get() or 0) for s in paut_keys)
        paut_today = sum(float(self.qty_entries[f"PAUT_{s}"]['금일작업'].get() or 0) for s in paut_keys)
        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in paut_keys)
        
        self.qty_entries['PAUT_소계']['예상량'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['전일누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['금일작업'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['공정률'].delete(0, tk.END)
        
        if paut_expected > 0:
            self.qty_entries['PAUT_소계']['예상량'].insert(0, f"{int(paut_expected)}")
        if paut_prev != 0:
            self.qty_entries['PAUT_소계']['전일누계'].insert(0, format_val('PAUT', paut_prev))
        if paut_today != 0:
            self.qty_entries['PAUT_소계']['금일작업'].insert(0, format_val('PAUT', paut_today))
        if paut_total != 0:
            self.qty_entries['PAUT_소계']['총누계'].insert(0, format_val('PAUT', paut_total))
            if paut_expected > 0:
                self.qty_entries['PAUT_소계']['공정률'].insert(0, f"{(paut_total / paut_expected) * 100:.1f}")
        
        rt_keys = ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간']
        rt_expected = sum(float(self.qty_entries[f"RT_{s}"]['예상량'].get() or 0) for s in rt_keys)
        rt_prev = sum(float(self.qty_entries[f"RT_{s}"]['전일누계'].get() or 0) for s in rt_keys)
        rt_today = sum(float(self.qty_entries[f"RT_{s}"]['금일작업'].get() or 0) for s in rt_keys)
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in rt_keys)
        
        self.qty_entries['RT_소계']['예상량'].delete(0, tk.END)
        self.qty_entries['RT_소계']['전일누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['금일작업'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['공정률'].delete(0, tk.END)
        
        if rt_expected > 0:
            self.qty_entries['RT_소계']['예상량'].insert(0, f"{int(rt_expected)}")
        if rt_prev != 0:
            self.qty_entries['RT_소계']['전일누계'].insert(0, format_val('RT', rt_prev))
        if rt_today != 0:
            self.qty_entries['RT_소계']['금일작업'].insert(0, format_val('RT', rt_today))
        if rt_total != 0:
            self.qty_entries['RT_소계']['총누계'].insert(0, format_val('RT', rt_total))
            if rt_expected > 0:
                self.qty_entries['RT_소계']['공정률'].insert(0, f"{(rt_total / rt_expected) * 100:.1f}")
                
        # RT_RE is a subset of the inspected films. 불량과 불량률도 총누계 기준이다.
        today_rejects['RT_소계'] = sum(value for key, value in today_rejects.items() if key.startswith('RT_') and key != 'RT_소계')
        for key, entries in self.qty_entries.items():
            if not key.startswith('RT_'):
                continue
            total = float(entries['총누계'].get() or 0)
            previous_defects = prev_data.get('qty_data', {}).get(key, {}).get('불량', '0') or '0'
            try:
                previous_defects = int(float(str(previous_defects).replace(',', '')))
            except (TypeError, ValueError):
                previous_defects = 0
            defects = previous_defects + today_rejects.get(key, 0)
            if '불량' in entries:
                entries['불량'].delete(0, tk.END)
                if defects:
                    entries['불량'].insert(0, str(defects))
            if '불량률' in entries:
                entries['불량률'].delete(0, tk.END)
                entries['불량률'].insert(0, f'{defects / total * 100:.1f}' if total and defects else '')

        # Equip
        for eq, entries in self.equip_entries.items():
            prev_total_str = prev_data.get('equip_data', {}).get(eq, {}).get('누계', '0') or '0'
            try: prev_total = float(prev_total_str.replace(',',''))
            except ValueError: prev_total = 0.0
            today_val = float(entries['금일'].get() or 0)
            total = prev_total + today_val
            entries['누계'].delete(0, tk.END)
            entries['누계'].insert(0, f"{total:.1f}" if total % 1 else f"{int(total)}")
            
        # Personnel
        prev_p_total_str = prev_data.get('personnel_data', {}).get('검사원_누계', '0') or '0'
        try: prev_p_total = float(prev_p_total_str.replace(',',''))
        except ValueError: prev_p_total = 0.0
        
        today_p = float(self.personnel_entries['검사원_인원'].get() or 0) + float(self.personnel_entries['검사원_현장대리인'].get() or 0)
        total_p = prev_p_total + today_p
        
        self.personnel_entries['검사원_누계'].delete(0, tk.END)
        self.personnel_entries['검사원_누계'].insert(0, f"{total_p:.1f}" if total_p % 1 else f"{int(total_p)}")
        
        prev_s_total_str = prev_data.get('personnel_data', {}).get('안전_누계', '0') or '0'
        try: prev_s_total = float(prev_s_total_str.replace(',',''))
        except ValueError: prev_s_total = 0.0
        
        today_s = float(self.personnel_entries['안전_인원'].get() or 0) + float(self.personnel_entries['안전_현장대리인'].get() or 0)
        total_s = prev_s_total + today_s
        
        self.personnel_entries['안전_누계'].delete(0, tk.END)
        self.personnel_entries['안전_누계'].insert(0, f"{total_s:.1f}" if total_s % 1 else f"{int(total_s)}")
        
        # Save to history
        self.save_current_history()
        
        if show_completion:
            messagebox.showinfo("완료", "자동 집계 및 데이터 저장이 완료되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test Tab")
    root.geometry("1400x800")
    tab = DailyWorkLogTab(root)
    tab.pack(fill="both", expand=True)
    root.mainloop()

