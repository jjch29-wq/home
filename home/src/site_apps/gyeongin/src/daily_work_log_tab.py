import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from site_apps.gyeongin.src.company_master import DEFAULT_COMPANIES
from site_apps.gyeongin.src.ndt_summary_exporter import NDTSummaryExporter
from site_apps.gyeongin.src.services.monthly_report_manager import MonthlyReportManager
from tkcalendar import DateEntry
from datetime import datetime
import os
from site_apps.gyeongin.src.data_sync import (
    SharedDataUnavailableError,
    get_data_connection_error,
    get_history_path,
    get_process_photos_dir,
)
import sys
import json
import re
import shutil
import uuid
from site_apps.gyeongin.src.rt_quantities import rt_counts, sync_rt_widgets, widget_rt_values, resolve_rt_row, retest_parent, default_rt_copy, corrected_rt_history

# Import the exporter
try:
    from site_apps.gyeongin.src.daily_work_log_exporter import DailyWorkLogExporter
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

NDT_TARGET_TYPES = ('배관', '판재/시험편', '기계부품', '구조물/설비', 'PMI')

NDT_ITEM_CHOICES = {
    '배관': (
        'PIPE', '강관', '배관 용접부', '튜브(TUBE)', '라이저 파이프',
        '매니폴드', '배관 피팅', '플랜지',
    ),
    '판재/시험편': (
        'PLATE', '판재 용접부', 'WQT 시험편', '맞대기 용접 시험편',
        '필릿 용접 시험편', '용접사 자격 시험편',
    ),
    '기계부품': (
        'SHAFT', 'PINION', 'GEAR', 'ROLL', 'UNIVERSAL JOINT',
        'CASING', 'BODY', 'IMPELLER', 'COUPLING', 'BOLT/NUT',
    ),
    '구조물/설비': (
        '구조물 용접부', '중앙말뚝', 'H-PILE', 'GUIDE RAIL', 'TANK',
        'PRESSURE VESSEL', 'HEAT EXCHANGER', 'PUMP', 'CRANE', 'FORKLIFT',
    ),
    'PMI': (
        'PIPE', 'FITTING', 'FLANGE', 'VALVE', 'PUMP', 'CASING/BODY',
        'BOLT/NUT', 'WELD', '기타 구성부품',
    ),
}

NDT_DIMENSION_TEMPLATES = {
    '배관': ('관경A × 두께T', '외경Ø × 두께T × 길이L'),
    '판재/시험편': ('폭 × 길이 × 두께T / 수량EA', '폭 × 폭 × 두께T × 두께T'),
    '기계부품': ('Ø직경 × 길이L / 수량EA', 'Ø직경1/직경2 × 길이L / 수량EA', '수량EA (검사 Point)'),
    '구조물/설비': ('폭 × 길이 × 높이 × 두께T', '용량 CAPA / 수량EA'),
    'PMI': ('N/A',),
}

NDT_COMMON_COLUMNS = (
    '업체', '검사방법', '대상유형', '검사품명', '식별번호', '재질',
    '치수원문', '수량', '단위', '결과',
)

NDT_VIEW_COLUMNS = {
    '배관': (
        '구간', '작업구분', '핫태핑개소', '핫태핑차수', '핫태핑개소수',
        '라인번호', 'Joint No.', '관경', '두께', '용접자세', '용접사',
        '구간정보', '규격', '근무구분', 'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT',
    ),
    '판재/시험편': ('시험편번호', '폭', '길이', '두께', '용접자세', '검사위치', '검사범위'),
    '기계부품': ('부품번호', '직경', '길이', '폭', '높이', '검사위치', '검사범위'),
    '구조물/설비': ('설비번호', '모델', '용량', '폭', '길이', '높이', '두께', '검사위치', '검사범위'),
    'PMI': ('태그번호', '구성부품', '요구재질', '판정재질', '검사위치'),
}

NDT_DETAIL_FIELDS = (
    '성적서번호', '발주처', '공사명', '의뢰자', '표면상태',
    '적용규격', '검사절차서', '개정번호', '도면번호', '검사위치', '검사범위', '비고'
)


def ndt_visible_columns(view):
    """Return the stable column order for one bulk-entry view."""
    view = view if view in NDT_VIEW_COLUMNS else '배관'
    allowed = set(NDT_COMMON_COLUMNS) | set(NDT_VIEW_COLUMNS[view])
    all_columns = tuple(dict.fromkeys(
        NDT_COMMON_COLUMNS
        + tuple(column for columns in NDT_VIEW_COLUMNS.values() for column in columns)
        + NDT_DETAIL_FIELDS
    ))
    return tuple(column for column in all_columns if column in allowed)


def infer_ndt_target(row):
    """Keep legacy work-log rows in the pipe workflow without rewriting history."""
    saved = str(row.get('대상유형', '') or '').strip()
    if saved in NDT_TARGET_TYPES:
        return saved
    pipe_keys = ('라인번호', 'Joint No.', '관경', '핫태핑개소', '구간정보')
    if any(str(row.get(key, '') or '').strip() for key in pipe_keys):
        return '배관'
    return ''


def ndt_item_choices(target_type, saved_items=()):
    """Combine reviewed report items with site history while preserving order."""
    choices = list(NDT_ITEM_CHOICES.get(target_type, ()))
    for item in saved_items:
        item = str(item or '').strip()
        if item and item not in choices:
            choices.append(item)
    return tuple(choices)


def ndt_dimension_choices(target_type, saved_dimensions=()):
    """Offer editable templates first and previously saved source text after them."""
    choices = list(NDT_DIMENSION_TEMPLATES.get(target_type, ()))
    for value in saved_dimensions:
        value = str(value or '').strip()
        if value and value not in choices:
            choices.append(value)
    return tuple(choices)


def parse_machine_dimension(value):
    """Parse common report text such as Ø90/220×120L / 112EA."""
    text = str(value or '').strip()
    if not text:
        return {}
    match = re.search(
        r'[ØΦφ⌀]\s*'
        r'(?P<diameter>\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)*)\s*'
        r'[×xX*]\s*(?P<length>\d+(?:\.\d+)?)\s*(?:L|ℓ)?',
        text,
    )
    if not match:
        return {}
    result = {
        '직경': re.sub(r'\s+', '', match.group('diameter')),
        '길이': match.group('length'),
    }
    quantity = re.search(r'/\s*(\d+(?:\.\d+)?)\s*(EA|PCS|개)\b', text[match.end():], re.IGNORECASE)
    if quantity:
        result['수량'] = quantity.group(1)
        unit = quantity.group(2).upper()
        result['단위'] = '개' if unit == '개' else unit
    return result

# 용접사 관리대장에서 갱신된 최신 발급번호.
# 과거 작업일보 및 공정사진에 저장된 번호도 같은 기준으로 자동 이관한다.
LEGACY_WELDER_IDS = {
    'W-2023-A-12': 'W-2026-A-08',
    'W-2023-A-13': 'W-2026-A-05',
}

class DailyWorkLogTab(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.history_path = get_history_path()
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
        # Do not call modal messagebox inside after_idle during __init__,
        # because subsequent tab setup calls update_idletasks() and would
        # block create_widgets() mid-initialization with grab_set().
        if self.data_connection_error:
            print(f"[경인 작업일보] 공유 경로 안내: {self.data_connection_error}")

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
            with open(self.history_path, 'r', encoding='utf-8') as history_file:
                history = json.load(history_file)

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
            if not changed_count:
                return

            temp_path = self.history_path + '.tmp'
            with open(temp_path, 'w', encoding='utf-8') as history_file:
                json.dump(migrated_history, history_file, ensure_ascii=False, indent=4)
            os.replace(temp_path, self.history_path)
            print(f"[DailyWorkLog] 용접사 번호 {changed_count}건 최신화 완료")
        except Exception as exc:
            print(f"[DailyWorkLog] 용접사 번호 최신화 실패: {exc}")

    def _format_welder_display(self, value):
        """저장된 용접사 번호를 '이름 | 번호' 화면 표시 형식으로 변환한다."""
        welder_id = self._normalize_welder_id(value)
        name = self.welder_names.get(welder_id)
        return f"{name} | {welder_id}" if name else welder_id
        
    def setup_ui(self):
        # Create PanedWindow for Left/Right split
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill="both", expand=True)
        
        # --- LEFT PANE ---
        self.left_container = ttk.Frame(self.paned)
        self.paned.add(self.left_container, weight=4) # 40% width
        
        self.left_canvas = tk.Canvas(self.left_container)
        self.left_scrollbar = ttk.Scrollbar(self.left_container, orient="vertical", command=self.left_canvas.yview)
        self.left_frame = ttk.Frame(self.left_canvas)
        
        def _update_left_scroll(e=None):
            bbox = self.left_canvas.bbox("all")
            if bbox:
                self.left_canvas.configure(scrollregion=(0, 0, bbox[2], max(bbox[3], self.left_canvas.winfo_height())))
        self.left_frame.bind("<Configure>", _update_left_scroll)
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
        
        def _on_left_canvas_configure(event):
            new_width = event.width
            new_height = max(event.height, self.left_frame.winfo_reqheight())
            self.left_canvas.itemconfig(self.left_window, width=new_width, height=new_height)
            _update_left_scroll()
            
        self.left_canvas.bind("<Configure>", _on_left_canvas_configure)
        
        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_scrollbar.pack(side="right", fill="y")
        
        self._build_left_pane(self.left_frame)
        
        # --- RIGHT PANE ---
        self.right_container = ttk.Frame(self.paned)
        self.paned.add(self.right_container, weight=6) # 60% width
        
        self.right_canvas = tk.Canvas(self.right_container)
        self.right_xscroll = ttk.Scrollbar(self.right_container, orient="horizontal", command=self.right_canvas.xview)
        self.right_yscroll = ttk.Scrollbar(self.right_container, orient="vertical", command=self.right_canvas.yview)
        self.right_frame = ttk.Frame(self.right_canvas)
        
        def _update_right_scroll(e=None):
            bbox = self.right_canvas.bbox("all")
            if bbox:
                self.right_canvas.configure(scrollregion=(0, 0, bbox[2], max(bbox[3], self.right_canvas.winfo_height())))
        self.right_frame.bind("<Configure>", _update_right_scroll)
        self.right_window = self.right_canvas.create_window((0, 0), window=self.right_frame, anchor="nw")
        self.right_canvas.configure(xscrollcommand=self.right_xscroll.set, yscrollcommand=self.right_yscroll.set)
        
        def _on_right_canvas_configure(event):
            # Stretch the embedded pane to the visible canvas so no unused strip remains.
            if event.width > self.right_frame.winfo_reqwidth():
                self.right_canvas.itemconfig(self.right_window, width=event.width)
            if event.height > self.right_frame.winfo_reqheight():
                self.right_canvas.itemconfig(self.right_window, height=event.height)
            _update_right_scroll()
        self.right_canvas.bind("<Configure>", _on_right_canvas_configure)
        
        self.right_canvas.grid(row=0, column=0, sticky="nsew")
        self.right_yscroll.grid(row=0, column=1, sticky="ns")
        self.right_xscroll.grid(row=1, column=0, sticky="ew")
        self.right_container.grid_rowconfigure(0, weight=1)
        self.right_container.grid_columnconfigure(0, weight=1)
        
        self._build_right_pane(self.right_frame)

        # 재시작할 때 좌우 작업영역을 화면 중앙에서 동일한 폭으로 시작한다.
        self.after(200, self._apply_initial_pane_layout)
        
        # Initial load for today's date
        self.after(100, self.on_date_change)

    def _apply_initial_pane_layout(self):
        try:
            self.update_idletasks()
            total_width = self.paned.winfo_width()
            if total_width > 200:
                self.paned.sashpos(0, total_width // 2)
        except (tk.TclError, AttributeError):
            pass

    def _build_left_pane(self, parent):
        # --- Top Section: General Info ---
        top_frame = ttk.LabelFrame(parent, text="기본 정보", padding=10)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(top_frame, text="검사일자:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.date_entry = DateEntry(top_frame, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.date_entry.bind("<<DateEntrySelected>>", self.on_date_change)
        self.date_entry.bind("<Return>", self.on_date_change)
        from site_apps.gyeongin.src.views.calendar_work_summary import CalendarWorkSummary
        self.calendar_work_summary = CalendarWorkSummary(
            self.date_entry, self,
            site_getter=lambda: '경인사무소',
            frame_getter=lambda: getattr(getattr(self, 'main_app', None), 'daily_usage_df', None),
            history_getter=self.load_history)
        # self.date_entry.bind("<FocusOut>", self.on_date_change) # Removed to prevent accidental UI wipes

        self.date_status_label = tk.Label(
            top_frame, text="○ 저장되지 않은 날짜", fg="#777777",
            font=('맑은 고딕', 9, 'bold')
        )
        self.date_status_label.grid(row=1, column=0, columnspan=2, padx=5, sticky="e")
        
        ttk.Label(top_frame, text="날씨:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.weather_entry = ttk.Entry(top_frame, width=15)
        self.weather_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        btn_calc = ttk.Button(top_frame, text="자동 집계 및 저장", command=self.auto_calculate_and_save)
        btn_calc.grid(row=0, column=4, padx=5, pady=5)
        
        # [NEW] 삭제 버튼 추가
        btn_delete = ttk.Button(top_frame, text="🗑️ 데이터 삭제", command=self.delete_current_date_data)
        btn_delete.grid(row=0, column=5, padx=5, pady=5)
        
        btn_export = ttk.Button(top_frame, text="엑셀 출력 (일보 생성)", command=self.export_excel)
        btn_export.grid(row=1, column=4, columnspan=2, padx=5, pady=(0, 3), sticky="e")
        
        # Saved inspection quantities are grouped by client, method and unit.
        mid_frame = ttk.LabelFrame(parent, text="1. 업체별 검사 실적", padding=10)
        mid_frame.pack(fill='both', expand=True, padx=5, pady=5)
        period = ttk.Frame(mid_frame)
        period.pack(fill='x', pady=(0, 5))
        ttk.Label(period, text='조회 시작일:').pack(side='left')
        self.company_start_date = DateEntry(period, width=12, date_pattern='yyyy-mm-dd', state='readonly')
        self.company_start_date.pack(side='left', padx=5)
        ttk.Label(period, text='~ 상단 검사일자').pack(side='left')
        ttk.Button(period, text='조회', command=self.refresh_company_summary).pack(side='right')
        self.company_start_date.bind('<<DateEntrySelected>>', self._company_period_changed)
        self.company_summary_rows = []
        self._summary_start_initialized = False
        columns = ('company', 'method', 'unit', 'previous', 'today', 'total')
        table = ttk.Frame(mid_frame)
        table.pack(fill='both', expand=True)
        self.company_summary_tree = ttk.Treeview(table, columns=columns, show='headings', height=12)
        for key, title, width in zip(columns,
                ('업체명', '검사방법', '단위', '전일까지', '선택일 실적', '누계'),
                (170, 75, 80, 95, 95, 95)):
            self.company_summary_tree.heading(key, text=title, anchor='center')
            self.company_summary_tree.column(key, width=width, minwidth=60,
                                             stretch=False, anchor='center')
        from site_apps.gyeongin.src.views.tree_column_resize import enable_column_resize
        enable_column_resize(self.company_summary_tree)
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.company_summary_tree.xview)
        horizontal.pack(side='bottom', fill='x')
        self.company_summary_tree.configure(xscrollcommand=horizontal.set)
        scrollbar = ttk.Scrollbar(table, orient='vertical', command=self.company_summary_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.company_summary_tree.configure(yscrollcommand=scrollbar.set)
        self.company_summary_tree.pack(side='left', fill='both', expand=True)
        self.company_summary_status = ttk.Label(mid_frame, text='저장된 일일 사용량 기준 · 미저장 입력 제외', wraplength=650)
        self.company_summary_status.pack(fill='x', pady=(5, 0))
        # Keep the old history field available for readers of archived daily logs.
        self.qty_entries = {}
        self.qty_rows = []
        self.default_qty = {}
        self.bind('<Map>', lambda event: self.refresh_company_summary(), add='+')

        # Archived equipment/personnel fields remain in history, without duplicate inputs.
        self.equip_rows = []
        self.equip_entries = {}
        self.personnel_entries = {}
        # Common and company notes use the remaining width.
        rm_frame = ttk.LabelFrame(parent, text='특이사항 및 계획', padding=10)
        rm_frame.pack(fill='x', padx=5, pady=5)
        memo_toolbar = ttk.Frame(rm_frame)
        memo_toolbar.pack(fill='x', pady=(0, 5))
        ttk.Label(memo_toolbar, text='저장: 상단 [자동 집계 및 저장]').pack(side='left')
        ttk.Button(memo_toolbar, text='메모 이력', command=self.open_memo_history).pack(side='right')
        self.memo_notebook = ttk.Notebook(rm_frame)
        self.memo_notebook.pack(fill='both', expand=True)
        common_page = ttk.Frame(self.memo_notebook, padding=5)
        self.company_memo_page = ttk.Frame(self.memo_notebook, padding=5)
        self.memo_notebook.add(common_page, text='공통 메모')
        self.memo_notebook.add(self.company_memo_page, text='업체별 메모')
        self.remarks_text = tk.Text(common_page, height=6, width=40, wrap='word')
        self.remarks_text.pack(fill='both', expand=True)
        company_toolbar = ttk.Frame(self.company_memo_page)
        company_toolbar.pack(fill='x', pady=(0, 4))
        ttk.Label(company_toolbar, text='업체명:').pack(side='left')
        self.company_memo_selector = ttk.Combobox(company_toolbar, state='readonly', width=28)
        self.company_memo_selector.pack(side='left', padx=5)
        self.company_memo_selector.bind('<<ComboboxSelected>>', self._on_company_memo_selected)
        self.company_memo_hint = ttk.Label(company_toolbar, text='실적 표에서 업체를 선택하세요.')
        self.company_memo_hint.pack(side='left', padx=5)
        self.company_memo_text = tk.Text(self.company_memo_page, height=5, width=40,
                                         wrap='word', state='disabled')
        self.company_memo_text.pack(fill='both', expand=True)
        self.company_memos = {}
        self._company_memo_company = None
        self.company_summary_tree.bind('<<TreeviewSelect>>', self._on_company_summary_selected)

    def _company_memo_snapshot(self):
        from site_apps.gyeongin.src.company_memos import memo_snapshot
        company = getattr(self, '_company_memo_company', None)
        text = self.company_memo_text.get('1.0', 'end-1c') if company else ''
        return memo_snapshot(getattr(self, 'company_memos', {}), company, text)

    def _refresh_company_memo_choices(self):
        names = {row['company'] for row in self.company_summary_rows
                 if row['company'] != '(업체명 미입력)'}
        names.update(self._company_memo_snapshot())
        if self._company_memo_company:
            names.add(self._company_memo_company)
        self.company_memo_selector.configure(values=sorted(names))

    def _select_company_memo(self, company):
        if not company or company == '(업체명 미입력)':
            return
        if company != self._company_memo_company:
            self.company_memos = self._company_memo_snapshot()
            self._company_memo_company = company
            self.company_memo_text.configure(state='normal')
            self.company_memo_text.delete('1.0', tk.END)
            self.company_memo_text.insert('1.0', self.company_memos.get(company, ''))
        self.company_memo_selector.set(company)
        self.company_memo_hint.configure(text=f'{self.date_entry.get()} · 같은 업체의 검사방법 간 공유')
        self.memo_notebook.select(self.company_memo_page)

    def _on_company_memo_selected(self, event=None):
        self._select_company_memo(self.company_memo_selector.get())

    def _on_company_summary_selected(self, event=None):
        selected = self.company_summary_tree.selection()
        if selected:
            values = self.company_summary_tree.item(selected[0], 'values')
            if values:
                self._select_company_memo(str(values[0]))

    def _load_company_memos(self, memos):
        from site_apps.gyeongin.src.company_memos import memo_snapshot
        self.company_memos = memo_snapshot(memos)
        self._company_memo_company = None
        self.company_memo_selector.set('')
        self.company_memo_text.configure(state='normal')
        self.company_memo_text.delete('1.0', tk.END)
        self.company_memo_text.configure(state='disabled')
        self.company_memo_hint.configure(text='실적 표에서 업체를 선택하세요.')
        self._refresh_company_memo_choices()

    def _company_period_changed(self, event=None):
        self._summary_start_initialized = True
        self.refresh_company_summary()

    def refresh_company_summary(self):
        from site_apps.gyeongin.src.company_usage_summary import first_usage_date, summarize_company_usage
        frame = getattr(getattr(self, 'main_app', None), 'daily_usage_df', None)
        selected = self.date_entry.get_date()
        if not self._summary_start_initialized and frame is not None:
            self.company_start_date.set_date(first_usage_date(frame, selected))
            self._summary_start_initialized = True
        try:
            rows, skipped = summarize_company_usage(frame, self.company_start_date.get_date(), selected)
        except ValueError as error:
            self.company_summary_rows = []
            self.company_summary_tree.delete(*self.company_summary_tree.get_children())
            self.company_summary_status.configure(text=str(error), foreground='#9A5B00')
            return False
        self.company_summary_rows = rows
        if hasattr(self, 'company_memo_selector'):
            self._refresh_company_memo_choices()
        self.company_summary_tree.delete(*self.company_summary_tree.get_children())
        for row in rows:
            self.company_summary_tree.insert('', 'end', values=(
                row['company'], row['method'], row['unit'],
                *(f"{row[key]:,.4f}".rstrip('0').rstrip('.') for key in ('previous', 'today', 'total'))))
        text = f'저장된 일일 사용량 기준 · {len(rows)}개 그룹 · 미저장 입력 제외'
        if not rows:
            text += ' · 조회 기간에 검사 실적이 없습니다.'
        if skipped:
            text += f' · 날짜/수량 오류 {skipped}건 제외'
        self.company_summary_status.configure(text=text, foreground='#9A5B00' if skipped else '#555555')
        return True

    def open_memo_history(self):
        """저장된 일보 메모를 날짜/키워드로 찾아 해당 날짜로 이동한다."""
        history = self.load_history()
        from site_apps.gyeongin.src.company_memos import memo_history_rows
        memo_rows = memo_history_rows(history)

        window = tk.Toplevel(self)
        window.title("일보 메모 이력")
        window.geometry("760x480")
        window.minsize(560, 320)
        window.transient(self.winfo_toplevel())

        search_frame = ttk.Frame(window, padding=10)
        search_frame.pack(fill="x")
        ttk.Label(search_frame, text="날짜·업체·내용 검색:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        table_frame = ttk.Frame(window, padding=(10, 0, 10, 10))
        table_frame.pack(fill="both", expand=True)
        columns = ('date', 'company', 'memo')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        tree.heading('date', text='날짜')
        tree.heading('company', text='업체 / 공통')
        tree.heading('memo', text='메모 내용')
        tree.column('company', width=140, minwidth=100, stretch=False)
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
            matched = [row for row in memo_rows if not query or query in ' '.join(row).lower()]
            for date_key, company, memo in matched:
                tree.insert('', 'end', values=(date_key, company or '공통', ' '.join(memo.split())))
            status_var.set(f"메모 {len(matched)}건  ·  더블클릭하면 해당 날짜로 이동합니다.")

        def move_to_selected(_event=None):
            selected = tree.selection()
            if not selected:
                return
            date_key = str(tree.item(selected[0], 'values')[0])
            try:
                self.date_entry.set_date(datetime.strptime(date_key, '%Y-%m-%d').date())
                self.on_date_change()
                if self._loaded_log_date == date_key:
                    company = str(tree.item(selected[0], 'values')[1])
                    if company != '공통':
                        self._select_company_memo(company)
                    else:
                        self.memo_notebook.select(0)
                    window.destroy()
            except ValueError:
                messagebox.showerror("날짜 오류", f"저장된 날짜 형식을 읽을 수 없습니다: {date_key}", parent=window)

        search_var.trace_add('write', refresh_rows)
        search_entry.bind('<Return>', lambda _event: move_to_selected())
        tree.bind('<Double-1>', move_to_selected)
        refresh_rows()
        search_entry.focus_set()

    def _build_right_pane(self, parent):
        # --- NDT Results Grid ---
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", pady=5)
        ttk.Label(title_frame, text="2. 비파괴검사결과서 (행 추가 가능)", font=("맑은 고딕", 12, "bold")).pack(side="left", padx=10)
        
        ttk.Button(title_frame, text="+ 10행 추가", command=lambda: self._append_ndt_rows(10, focus=True)).pack(side="left", padx=5)

        btn_export_ndt = ttk.Button(title_frame, text="NDT 누계 대장 엑셀 출력", command=self.export_ndt_summary)
        btn_export_ndt.pack(side="left", padx=10)
        
        btn_export_monthly = ttk.Button(title_frame, text="월간진도보고서 출력", command=self.export_monthly_report)
        btn_export_monthly.pack(side="left", padx=10)
        
        btn_weekly_report = ttk.Button(title_frame, text="🗓️ 주간업무보고", command=lambda: self.main_app.export_weekly_report() if hasattr(self, 'main_app') else None)
        btn_weekly_report.pack(side="left", padx=10)

        line_bar = ttk.Frame(parent)
        line_bar.pack(fill='x', padx=5, pady=(0, 4))
        ttk.Button(line_bar, text='선택 행 삭제', command=self._delete_selected_ndt_row).pack(side='left', padx=(0, 8))
        self.ndt_selection_vars = []
        self.ndt_selection_text = tk.StringVar(value='왼쪽 체크박스로 여러 행 선택 가능')
        ttk.Label(line_bar, textvariable=self.ndt_selection_text).pack(side='left', padx=(0, 8))
        ttk.Button(line_bar, text='라인번호 일괄 변경', command=self._edit_ndt_lines).pack(side='left')
        ttk.Button(line_bar, text='구간 일괄 변경', command=self._edit_ndt_sections).pack(side='left', padx=(8, 0))
        ttk.Button(line_bar, text='선택 행 상세정보', command=self._edit_ndt_details).pack(side='left', padx=(8, 0))
        ttk.Button(line_bar, text='📄 시험성적서 출력', command=self.export_certificates).pack(side='left', padx=(12, 0))
        ttk.Button(line_bar, text='📒 성적서 발행대장', command=self.export_certificate_ledger).pack(side='left', padx=(8, 0))

        view_bar = ttk.LabelFrame(parent, text='검사 대상 보기')
        view_bar.pack(fill='x', padx=5, pady=(0, 4))
        self.ndt_view_var = tk.StringVar(value='배관')
        for target_type in NDT_TARGET_TYPES:
            ttk.Radiobutton(
                view_bar, text=target_type, value=target_type,
                variable=self.ndt_view_var, command=self._set_ndt_view,
                style='Toolbutton',
            ).pack(side='left', padx=6, pady=3)
        ttk.Label(
            view_bar,
            text='공통정보는 유지하고 선택한 대상에 필요한 열만 표시합니다.',
        ).pack(side='left', padx=12)
        
        # Grid Container
        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        all_type_columns = tuple(dict.fromkeys(
            column for columns in NDT_VIEW_COLUMNS.values() for column in columns
        ))
        self.ndt_cols = tuple(dict.fromkeys(
            NDT_COMMON_COLUMNS + all_type_columns + NDT_DETAIL_FIELDS
        ))
        history = self.load_history()
        sections = set()
        lines = set()
        companies = set(self._configured_companies())
        history_items = {target: set() for target in NDT_TARGET_TYPES}
        history_dimensions = {target: set() for target in NDT_TARGET_TYPES}
        # 월간 보고서와 동일한 용접사 번호/성명 기준정보를 사용한다.
        welders = set(self.welder_names)
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
                if r.get('업체'): companies.add(r['업체'].strip())
                if r.get('용접사'): welders.add(self._normalize_welder_id(r['용접사']))
                target_type = infer_ndt_target(r)
                item_name = str(r.get('검사품명', '') or '').strip()
                if target_type in history_items and item_name:
                    history_items[target_type].add(item_name)
                dimension = str(r.get('치수원문', '') or '').strip()
                if target_type in history_dimensions and dimension:
                    history_dimensions[target_type].add(dimension)
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
        self.history_companies = self._format_company_choices(companies)
        self.history_ndt_items = {
            target: sorted(items) for target, items in history_items.items()
        }
        self.history_ndt_dimensions = {
            target: sorted(values) for target, values in history_dimensions.items()
        }
        self.history_welders = [''] + [
            self._format_welder_display(welder_id)
            for welder_id in sorted(welders)
        ]

        
        # Draw Headers
        ttk.Label(grid_frame, text='선택').grid(row=0, column=0, padx=4)
        self._ndt_header_widgets = {}
        self._ndt_column_indices = {}
        for col_idx, c in enumerate(self.ndt_cols, start=1):
            label = {
                'RT_OR': 'RT_OR 전체', 'RT_RE': 'RT_RE 불량',
                '규격': '필름규격', '대상유형': '대상 유형',
            }.get(c, c)
            header = ttk.Label(grid_frame, text=label, font=("맑은 고딕", 9, "bold"))
            header.grid(row=0, column=col_idx, padx=1, pady=2)
            self._ndt_header_widgets[c] = header
            self._ndt_column_indices[c] = col_idx
            if c == '구간정보':
                grid_frame.grid_columnconfigure(col_idx, weight=5, minsize=260)
            elif c == '라인번호':
                grid_frame.grid_columnconfigure(col_idx, weight=3)
            elif c == '업체':
                grid_frame.grid_columnconfigure(col_idx, weight=2, minsize=145)
            elif c == '규격':
                grid_frame.grid_columnconfigure(col_idx, weight=2, minsize=105)
            else:
                grid_frame.grid_columnconfigure(col_idx, weight=1)
            
        self._ndt_grid_frame = grid_frame
        self.ndt_grid_entries = []
        self._append_ndt_rows(30)
        self._set_ndt_view()

        hot_bar = ttk.Frame(parent)
        hot_bar.pack(fill='x', padx=5, pady=4)
        ttk.Button(hot_bar, text='선택 행 핫태핑 입력/수정', command=self._edit_hot_tapping).pack(side='left')
        self.hot_summary_var = tk.StringVar()
        ttk.Label(hot_bar, textvariable=self.hot_summary_var).pack(side='left', padx=10)
        photo_bar = ttk.LabelFrame(parent, text="선택 행 공정사진")
        photo_bar.pack(fill='x', padx=5, pady=(8, 0))
        self.photo_selection_var = tk.StringVar(value="NDT 행을 선택하세요.")
        ttk.Label(photo_bar, textvariable=self.photo_selection_var).pack(
            side='left', padx=8, pady=6
        )
        ttk.Button(
            photo_bar, text="사진 추가", command=self.add_process_photo
        ).pack(side='left', padx=4)
        ttk.Button(
            photo_bar, text="사진 관리", command=self.manage_process_photos
        ).pack(side='left', padx=4)

    def _append_ndt_rows(self, count=10, focus=False):
        first = len(self.ndt_grid_entries)
        for row_idx in range(first + 1, first + count + 1):
            self._create_ndt_row(row_idx)
        if focus and count:
            self._focus_ndt_joint(first)

    def _ensure_ndt_capacity(self, required):
        missing = required - len(self.ndt_grid_entries)
        if missing > 0:
            self._append_ndt_rows(((missing + 9) // 10) * 10)

    def _focus_ndt_joint(self, index):
        entry = self.ndt_grid_entries[index]['Joint No.']
        entry.focus_set()
        self.update_idletasks()
        y = self._ndt_grid_frame.winfo_y() + entry.winfo_y()
        top = self.right_canvas.canvasy(0)
        if y < top or y + entry.winfo_height() > top + self.right_canvas.winfo_height():
            self.right_canvas.yview_moveto(y / max(1, self.right_frame.winfo_height()))

    def _create_ndt_row(self, row_idx):
        grid_frame = self._ndt_grid_frame
        row_entries = {}
        selected = tk.BooleanVar(value=False)
        self.ndt_selection_vars.append(selected)
        ttk.Checkbutton(
            grid_frame, text=str(row_idx), variable=selected,
            command=lambda: self._toggle_ndt_row(row_entries, selected)
        ).grid(row=row_idx, column=0, sticky='w', padx=4)

        for col_idx, c in enumerate(self.ndt_cols, start=1):
            # Adjust width for some columns
            w = 8
            if c == '규격': w = 13
            elif c == '근무구분': w = 7
            elif c in ('검사방법', '결과', '관경', '두께'): w = 6
            elif c == '업체': w = 16
            elif c == '대상유형': w = 12
            elif c in ('검사품명', '치수원문', '식별번호'): w = 18
            elif c in ('검사위치', '검사범위', '구성부품'): w = 16
            elif c == '구간': w = 10
            elif c == '용접사': w = 23
            elif c == '라인번호': w = 25
            elif c == 'Joint No.': w = 12
            elif c == '구간정보': w = 35
            else: w = 8

            if c in ('작업구분', '핫태핑차수'):
                values = ['', '일반', '핫태핑'] if c == '작업구분' else ['', '1차', '2차']
                ent = ttk.Combobox(grid_frame, width=7, values=values, state='readonly')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky='ew')
                row_entries[c] = ent
            elif c == '검사방법':
                ent = ttk.Combobox(grid_frame, width=w, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'], justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '대상유형':
                ent = ttk.Combobox(
                    grid_frame, width=w, values=[''] + list(NDT_TARGET_TYPES),
                    state='readonly', justify='center'
                )
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky='ew')
                row_entries[c] = ent
            elif c == '검사품명':
                ent = ttk.Combobox(grid_frame, width=w, values=[], justify='left')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky='ew')
                row_entries[c] = ent
            elif c == '치수원문':
                ent = ttk.Combobox(grid_frame, width=w, values=[], justify='left')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky='ew')
                row_entries[c] = ent
            elif c == '구간정보':
                frame = ttk.Frame(grid_frame)
                frame.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                entries = []
                for i in range(4):
                    e = ttk.Entry(frame, width=5, justify='center')
                    e.pack(side='left', fill='x', expand=True, padx=(0, 1 if i<3 else 0))
                    entries.append(e)
                frame.entries = entries
                def update_rt_films(event=None, r=row_entries):
                    try:
                        sync_rt_widgets(r, from_details=True, all_rows=self._rt_context())
                    except ValueError:
                        pass  # Allow an incomplete value while typing; validate before save.
                for entry in entries:
                    entry.bind('<KeyRelease>', update_rt_films, add='+')
                    entry.bind('<FocusOut>', update_rt_films, add='+')
                def apply_layout(ents=entries, r_dict=row_entries):
                    method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                    if method in ['PAUT', 'PT', 'MT']:
                        ents[0].configure(justify='left')
                        for i in range(1, 4):
                            ents[i].pack_forget()
                    else:
                        ents[0].configure(justify='center')
                        for i in range(1, 4):
                            if not ents[i].winfo_manager():
                                ents[i].pack(
                                    side='left', fill='x', expand=True,
                                    padx=(0, 1 if i < 3 else 0)
                                )
                def get_val(ents=entries, r_dict=row_entries):
                    method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                    if method in ['PAUT', 'PT', 'MT']:
                        return ents[0].get()
                    parts = [e.get().strip() for e in ents]
                    return ','.join(parts).rstrip(',') if method == 'RT' else ','.join([p for p in parts if p])
                def set_val(val, ents=entries, r_dict=row_entries, layout_fn=apply_layout):
                    for e in ents: e.delete(0, tk.END)
                    method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                    if method in ['PAUT', 'PT', 'MT']:
                        ents[0].insert(0, val)
                    else:
                        parts = val.split(',') if val else []
                        for i, p in enumerate(parts):
                            if i < len(ents): ents[i].insert(0, p)
                    layout_fn()
                    try:
                        sync_rt_widgets(r_dict, all_rows=self._rt_context())
                    except ValueError:
                        pass
                def delete_val(first, last, ents=entries):
                    for e in ents: e.delete(first, last)
                def insert_val(idx, val, ents=entries):
                    ents[0].insert(idx, val)
                frame.get = get_val
                frame.set = set_val
                frame.delete = delete_val
                frame.insert = insert_val
                frame.apply_layout = apply_layout
                row_entries[c] = frame
            elif c == '관경':
                ent = ttk.Combobox(grid_frame, width=w, values=[''] + list(SIZE_LENGTH.keys()), justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '결과':
                ent = ttk.Combobox(grid_frame, width=w, values=['', '합격', '불합격', '재촬영', '보류'], justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '규격':
                ent = ttk.Combobox(grid_frame, width=w, values=[''], justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '근무구분':
                ent = ttk.Combobox(
                    grid_frame, width=w, values=['', '주간', '야간', '재검'],
                    justify='center'
                )
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '용접사':
                ent = ttk.Combobox(
                    grid_frame, width=w, values=self.history_welders,
                    justify='center'
                )
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '업체':
                ent = ttk.Combobox(grid_frame, width=w, values=self.history_companies, justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '구간':
                ent = ttk.Combobox(grid_frame, width=w, values=self.history_sections, justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            elif c == '라인번호':
                ent = ttk.Combobox(grid_frame, width=w, values=self.history_lines, justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent
            else:
                ent = ttk.Entry(grid_frame, width=w, justify='center')
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                row_entries[c] = ent

        def on_ndt_change(event, r=row_entries):
            method = r['검사방법'].get().strip().upper()
            pipe = r['관경'].get().strip().upper()
            self._sync_ndt_spec(r)

            # --- 관경 선택 시 두께 자동 입력 ---
            thk_mapping = {
                "1100": "11.1", "1000": "11.1", "900": "10.3", "850": "10.3", "800": "9.5",
                "750": "8.7", "700": "8.7", "650": "8.7", "600": "9.5", "550": "9.5",
                "500": "6.4", "450": "6.4", "400": "6.4", "350": "6.4", "300": "6.4", "250": "6.4", "200": "6.4",
                "150": "5.5", "125": "5.1", "100": "4.9", "80": "4.5", "65": "4.5", "50": "3.9",
                "40": "3.7", "32": "3.6", "25": "3.4", "20": "2.9"
            }
            lookup_size = str(pipe).replace("A", "").strip()
            if lookup_size in thk_mapping and r['작업구분'].get() != '핫태핑':
                r['두께'].delete(0, tk.END)
                r['두께'].insert(0, thk_mapping[lookup_size])
            # -----------------------------------

            if pipe and pipe.isdigit():
                pipe += 'A'

            if r['작업구분'].get() == '핫태핑':
                from site_apps.gyeongin.src.hot_tapping import inspection_length
                try:
                    count = int(r['핫태핑개소수'].get())
                    if method in ('MT', 'PT') and count > 0:
                        length = inspection_length(pipe, r['핫태핑차수'].get()) * count
                        for col in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT'):
                            r[col].delete(0, tk.END)
                        r[method].insert(0, f'{length:.4f}')
                except ValueError:
                    pass  # 미입력 조건은 저장 시 검증한다.
                r['구간정보'].apply_layout()
                self._refresh_hot_summary()
                return

            if pipe in SIZE_LENGTH:
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

                target_col = None
                if method == 'PAUT': target_col = 'PAUT'
                elif method == 'RT': target_col = 'RT_OR'
                elif method == 'MT': target_col = 'MT'
                elif method == 'PT': target_col = 'PT'

                if target_col and not r[target_col].get().strip():
                    # Clear other length columns if we are auto-filling
                    for col in ['RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT']:
                        r[col].delete(0, tk.END)
                    r[target_col].insert(0, length_str)

            # Update '구간정보' display based on method
            if '구간정보' in r and hasattr(r['구간정보'], 'entries'):
                r['구간정보'].apply_layout()
            self._apply_company_method_preset(r, source='method')

        def on_target_change(event=None, r=row_entries):
            target_type = r['대상유형'].get().strip()
            if target_type in NDT_TARGET_TYPES:
                self._sync_ndt_item_choices(r)
                self._sync_ndt_dimension_choices(r)
                self._sync_machine_dimensions(r)
                self.ndt_view_var.set(target_type)
                self._set_ndt_view()


        def on_hot_change(event=None, r=row_entries):
            if r['작업구분'].get() == '핫태핑':
                self._edit_hot_tapping(r)
        row_entries['업체'].bind(
            '<<ComboboxSelected>>',
            lambda event, r=row_entries: self._apply_company_method_preset(r, source='company'),
            add='+',
        )
        row_entries['업체'].bind(
            '<FocusOut>',
            lambda event, r=row_entries: self._apply_company_method_preset(r, source='company'),
            add='+',
        )
        row_entries['작업구분'].bind('<<ComboboxSelected>>', on_hot_change)
        row_entries['검사방법'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['검사방법'].bind('<FocusOut>', on_ndt_change)
        row_entries['대상유형'].bind('<<ComboboxSelected>>', on_target_change)
        row_entries['관경'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['관경'].bind('<FocusOut>', on_ndt_change)
        row_entries['치수원문'].bind(
            '<FocusOut>', lambda event, r=row_entries: self._sync_machine_dimensions(r), add='+'
        )
        row_entries['치수원문'].bind(
            '<<ComboboxSelected>>', lambda event, r=row_entries: self._sync_machine_dimensions(r), add='+'
        )
        row_entries['Joint No.'].bind('<FocusOut>', lambda event, r=row_entries: self._sync_rt_row(r), add='+')

        def on_joint_enter(event, r_idx=row_idx-1, r=row_entries):
            import re
            if r_idx == 0:
                if len(self.ndt_grid_entries) > 1:
                    self.ndt_grid_entries[1]['Joint No.'].focus_set()
                return 'break'

            prev_row = self.ndt_grid_entries[r_idx - 1]

            # Auto-increment Joint No.
            curr_joint = r['Joint No.'].get().strip()
            if not curr_joint:
                prev_joint = prev_row['Joint No.'].get().strip()
                if prev_joint:
                    match = re.search(r'(\d+)$', prev_joint)
                    if match:
                        num_str = match.group(1)
                        prefix = prev_joint[:-len(num_str)]
                        next_num = str(int(num_str) + 1).zfill(len(num_str))
                        r['Joint No.'].insert(0, prefix + next_num)

            for col in ['업체', '검사방법', '대상유형', '검사품명', '재질', '구간', '작업구분', '핫태핑개소', '핫태핑차수', '핫태핑개소수', '라인번호', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분', '적용규격', '검사절차서', '개정번호']:
                if col == '규격' and r['검사방법'].get().strip().upper() != prev_row['검사방법'].get().strip().upper():
                    continue
                if col in ('구간정보', '결과') and (r['검사방법'].get().strip().upper() == 'RT' or prev_row['검사방법'].get().strip().upper() == 'RT'):
                    continue
                if not r[col].get().strip():
                    prev_val = prev_row[col].get().strip()
                    if prev_val:
                        if col == '구간정보':
                            r[col].set(prev_val)
                        elif isinstance(r[col], ttk.Combobox):
                            r[col].set(prev_val)
                        else:
                            r[col].delete(0, tk.END)
                            r[col].insert(0, prev_val)

            on_ndt_change(None, r)
            defaults = default_rt_copy(widget_rt_values(r), self._rt_context())
            if defaults.get('구간정보', '') != r['구간정보'].get():
                r['구간정보'].set(defaults['구간정보'])
            self._sync_rt_row(r)

            if r_idx + 1 == len(self.ndt_grid_entries):
                self._append_ndt_rows(10)
            if r_idx + 1 < len(self.ndt_grid_entries):
                self._focus_ndt_joint(r_idx + 1)
            return 'break'

        for col_name, widget in row_entries.items():
            if hasattr(widget, 'entries'):
                for e in widget.entries:
                    e.bind('<Return>', on_joint_enter)
                    e.bind('<Button-1>', lambda event, r=row_entries: self._select_ndt_row(r), add='+')
                    e.bind('<Escape>', self._clear_ndt_selection, add='+')
            else:
                widget.bind('<Return>', on_joint_enter)
                widget.bind('<Button-1>', lambda event, r=row_entries: self._select_ndt_row(r), add='+')
                widget.bind('<Escape>', self._clear_ndt_selection, add='+')

        self.ndt_grid_entries.append(row_entries)

        if hasattr(self, 'ndt_view_var'):
            self._set_ndt_view()

    def _apply_company_method_preset(self, row, source='company'):
        """Apply (company, method) target type, item dropdown ordering, and default code."""
        from site_apps.gyeongin.src.company_master import suggest_company_method_preset

        company = row['업체'].get().strip()
        method = row['검사방법'].get().strip().upper()
        preset = suggest_company_method_preset(company, method)
        suggested_target = preset.get('target_type', '')
        current_target = row['대상유형'].get().strip()
        has_pipe_id = any(row[k].get().strip() for k in ('라인번호', 'Joint No.', '관경') if k in row)

        if (
            suggested_target in NDT_TARGET_TYPES
            and not has_pipe_id
            and (not current_target or (source == 'company' and not row['검사품명'].get().strip()))
        ):
            if current_target != suggested_target:
                row['대상유형'].set(suggested_target)
            self._sync_ndt_dimension_choices(row)
            if hasattr(self, 'ndt_view_var') and self.ndt_view_var.get() != suggested_target:
                self.ndt_view_var.set(suggested_target)
                self._set_ndt_view()

        self._sync_ndt_item_choices(row)

        default_code = preset.get('default_code', '')
        if default_code and '적용규격' in row and not row['적용규격'].get().strip():
            row['적용규격'].insert(0, default_code)

    def _sync_ndt_item_choices(self, row):
        from site_apps.gyeongin.src.company_master import merge_company_items

        target_type = row['대상유형'].get().strip()
        company = row['업체'].get().strip()
        method = row['검사방법'].get().strip().upper()
        saved = getattr(self, 'history_ndt_items', {}).get(target_type, ())
        current = row['검사품명'].get().strip()
        target_defaults = ndt_item_choices(target_type, tuple(saved) + ((current,) if current else ()))
        choices = merge_company_items(
            company,
            method,
            history_items=((current,) if current else ()) + tuple(saved),
            fallback_items=target_defaults,
        )
        row['검사품명'].configure(values=choices)

    def _sync_ndt_dimension_choices(self, row):
        target_type = row['대상유형'].get().strip()
        saved = getattr(self, 'history_ndt_dimensions', {}).get(target_type, ())
        current = row['치수원문'].get().strip()
        choices = ndt_dimension_choices(target_type, tuple(saved) + ((current,) if current else ()))
        row['치수원문'].configure(values=choices)

    def _sync_machine_dimensions(self, row):
        """Fill only empty structured fields; never replace the entered source text."""
        if row['대상유형'].get().strip() != '기계부품':
            return
        parsed = parse_machine_dimension(row['치수원문'].get())
        for field in ('직경', '길이', '수량', '단위'):
            if parsed.get(field) and not row[field].get().strip():
                row[field].insert(0, parsed[field])

    def _set_ndt_view(self):
        """Show the compact common grid plus fields relevant to one target type."""
        if not hasattr(self, '_ndt_grid_frame'):
            return
        view = self.ndt_view_var.get() if hasattr(self, 'ndt_view_var') else '배관'
        if view not in NDT_VIEW_COLUMNS:
            view = '배관'
        visible = set(ndt_visible_columns(view))
        for column, col_idx in self._ndt_column_indices.items():
            is_visible = column in visible
            header = self._ndt_header_widgets.get(column)
            widgets = [row.get(column) for row in getattr(self, 'ndt_grid_entries', [])]
            if is_visible:
                if header is not None:
                    header.grid()
                for widget in widgets:
                    if widget is not None:
                        widget.grid()
                weight = 3 if column in ('라인번호', '검사품명', '치수원문') else 1
                minsize = 130 if column in ('업체', '검사품명', '치수원문') else 0
                self._ndt_grid_frame.grid_columnconfigure(col_idx, weight=weight, minsize=minsize)
            else:
                if header is not None:
                    header.grid_remove()
                for widget in widgets:
                    if widget is not None:
                        widget.grid_remove()
                self._ndt_grid_frame.grid_columnconfigure(col_idx, weight=0, minsize=0)

    def _edit_ndt_details(self):
        """Edit report metadata that should not permanently widen the bulk grid."""
        row = getattr(self, 'selected_ndt_row', None)
        if row is None:
            messagebox.showwarning('상세정보', '먼저 검사 행을 선택하세요.')
            return

        dialog = tk.Toplevel(self)
        dialog.title('선택 행 상세정보')
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=15)
        body.pack(fill='both', expand=True)

        target_type = row['대상유형'].get().strip() or self.ndt_view_var.get()
        summary = ' / '.join(filter(None, (
            row['업체'].get().strip(), row['검사방법'].get().strip(),
            row['검사품명'].get().strip(), row['식별번호'].get().strip(),
        ))) or '새 검사 행'
        ttk.Label(body, text=f'{target_type} · {summary}', font=('맑은 고딕', 10, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0, 10)
        )

        variables = {}
        for index, field in enumerate(NDT_DETAIL_FIELDS, start=1):
            variables[field] = tk.StringVar(value=row[field].get())
            ttk.Label(body, text=field).grid(row=index, column=0, sticky='w', padx=(0, 12), pady=4)
            ttk.Entry(body, textvariable=variables[field], width=56).grid(
                row=index, column=1, sticky='ew', pady=4
            )

        ttk.Label(
            body,
            text='치수 원문과 수량은 표에서 입력하고, 규격·절차·위치·범위는 여기에서 보완합니다.',
            foreground='#555555',
        ).grid(row=len(NDT_DETAIL_FIELDS) + 1, column=0, columnspan=2, sticky='w', pady=(8, 0))

        def apply():
            if not row['대상유형'].get().strip():
                row['대상유형'].set(target_type)
            for field, variable in variables.items():
                row[field].delete(0, tk.END)
                row[field].insert(0, variable.get().strip())
            dialog.destroy()

        buttons = ttk.Frame(body)
        buttons.grid(row=len(NDT_DETAIL_FIELDS) + 2, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text='적용', command=apply).pack(side='left', padx=4)
        ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left')
        dialog.grab_set()


    def _delete_selected_ndt_row(self):
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

    def _replace_ndt_line(self, old_line, new_line):
        old_line, new_line = old_line.strip(), new_line.strip()
        if not old_line or not new_line:
            raise ValueError('기존 라인번호와 새 라인번호를 입력하세요.')
        if old_line == new_line:
            raise ValueError('새 라인번호가 기존 번호와 같습니다.')
        targets = [row for row in self.ndt_grid_entries if row['라인번호'].get().strip() == old_line]
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
        selected_only = tk.BooleanVar(value=False)
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
        preview = tk.StringVar()
        ttk.Label(body, text=f'현재 날짜: {self.date_entry.get()}').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Label(body, text='기존 라인번호').grid(row=1, column=0, sticky='w', padx=(0, 10))
        ttk.Combobox(body, textvariable=old, values=lines, state='readonly', width=42).grid(row=1, column=1, pady=4)
        ttk.Label(body, text='새 라인번호').grid(row=2, column=0, sticky='w')
        entry = ttk.Entry(body, textvariable=new, width=45)
        entry.grid(row=2, column=1, pady=4)
        ttk.Label(body, textvariable=preview).grid(row=3, column=0, columnspan=2, sticky='w', pady=10)
        ttk.Label(body, text='변경 후 상단 저장 버튼을 눌러 주세요.').grid(row=4, column=0, columnspan=2, sticky='w')

        def update_preview(*args):
            count = sum(row['라인번호'].get().strip() == old.get() for row in self.ndt_grid_entries)
            preview.set(f'변경 대상: {count}행 (기존 라인번호가 일치하는 행)')

        def apply():
            try:
                count = self._replace_ndt_line(old.get(), new.get())
            except ValueError as error:
                messagebox.showwarning('라인번호 일괄 변경', str(error), parent=dialog)
                return
            dialog.destroy()
            messagebox.showinfo('라인번호 일괄 변경', f'{count}행을 변경했습니다. 상단 저장 버튼을 눌러 주세요.')

        old.trace_add('write', update_preview)
        update_preview()
        buttons = ttk.Frame(body)
        buttons.grid(row=5, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text='변경 적용', command=apply).pack(side='left', padx=4)
        ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left')
        dialog.grab_set()
        entry.focus_set()

    def _sync_ndt_spec(self, row):
        import re
        method = row['검사방법'].get().strip().upper()
        spec = row['규격']
        spec['values'] = ['', '3 1/3 X 12"', '3 1/3 X 6"'] if method == 'RT' else ['']
        if method and method != 'RT' and re.fullmatch(r'3\s*1\s*/\s*3\s*[X×]\s*(?:12|6)\s*["”″]?', spec.get().strip(), re.IGNORECASE):
            spec.set('')

    def _edit_hot_tapping(self, row=None):
        from site_apps.gyeongin.src.hot_tapping import LENGTHS, inspection_length
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
        self.hot_summary_var.set(f"핫태핑: {len(locations - {''})}개소명 / MT {totals['MT']:.4f} m / PT {totals['PT']:.4f} m")

    def _validate_hot_rows(self):
        from site_apps.gyeongin.src.hot_tapping import inspection_length
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
        self.selected_ndt_row = row_entries
        index = next(i for i, row in enumerate(self.ndt_grid_entries, 1) if row is row_entries)
        self.ndt_selection_vars[index - 1].set(True)
        self._refresh_ndt_selection_text()
        method = row_entries['검사방법'].get().strip() or '-'
        section = row_entries['구간'].get().strip() or '-'
        joint = row_entries['Joint No.'].get().strip() or '-'
        welder = row_entries['용접사'].get().strip() or '-'
        self.photo_selection_var.set(
            f"선택: {method} / {section} / Joint {joint} / {welder}"
        )

    def _refresh_ndt_selection_text(self):
        count = sum(bool(var.get()) for var in self.ndt_selection_vars)
        self.ndt_selection_text.set(f'{count}개 행 선택됨' if count else '왼쪽 체크박스로 여러 행 선택 가능')

    def _toggle_ndt_row(self, row, selected):
        if selected.get():
            self._select_ndt_row(row)
        else:
            if self.selected_ndt_row is row:
                self.selected_ndt_row = None
                self.photo_selection_var.set('NDT 행을 선택하세요.')
            self._refresh_ndt_selection_text()

    def _clear_ndt_selection(self, event=None):
        self.selected_ndt_row = None
        for var in self.ndt_selection_vars:
            var.set(False)
        self._refresh_ndt_selection_text()
        self.photo_selection_var.set("NDT 행을 선택하세요.")

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
        day_photos = day_data.get('process_photos', [])
        photo_count = len(day_photos)

        confirm_message = (
            f"정말로 {date_str}의 작업일보 저장 내용을 모두 삭제하시겠습니까?\n"
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
            del history[date_str]
            self.save_history(history)

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
        self._load_company_memos({})

        for row in self.ndt_grid_entries:
            for ent in row.values():
                if isinstance(ent, ttk.Combobox):
                    ent.set('')
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
            
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
        default_name = f"{self.date_entry.get().replace('-', '')}_작업일보.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            title="일보 저장",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if not file_path:
            return
        if self.auto_calculate_and_save(show_completion=False) is False:
            return
        archived = self.load_history().get(self._selected_date_key(), {})
        # Gather all data
        data = {
            'date': self.date_entry.get(),
            'weather': self.weather_entry.get(),
            'qty_data': {},
            'summary_start_date': self.company_start_date.get(),
            'company_summary': self.company_summary_rows,
            'equip_data': archived.get('equip_data', {}),
            'personnel_data': archived.get('personnel_data', {}),
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'company_memos': DailyWorkLogTab._company_memo_snapshot(self),
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
            exporter = NDTSummaryExporter(history)
            exporter.generate(file_path)
            messagebox.showinfo("완료", "NDT 누계 대장 엑셀 파일이 성공적으로 생성되었습니다.")
            os.startfile(os.path.dirname(file_path))
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"NDT 누계 대장 생성 중 오류가 발생했습니다:\n{str(e)}")

    def _current_ndt_rows_for_export(self):
        checked = []
        all_rows = []
        for idx, row_entries in enumerate(self.ndt_grid_entries):
            row_dict = {
                col: ent.get() for col, ent in row_entries.items() if hasattr(ent, 'get')
            }
            if not any(str(v).strip() for v in row_dict.values()):
                continue
            all_rows.append(row_dict)
            if idx < len(self.ndt_selection_vars) and self.ndt_selection_vars[idx].get():
                checked.append(row_dict)
        return checked if checked else all_rows

    def _daily_usage_frame(self):
        app_obj = getattr(self, 'main_app', None) or getattr(self, 'app', None)
        return getattr(app_obj, 'daily_usage_df', None)

    def export_certificates(self):
        """선택 행(또는 당일 전체)을 업체·검사방법별 시험성적서(갑/을)로 출력합니다."""
        from site_apps.gyeongin.src.cert_exporter import (
            CertificateExporter,
            CertificateLedgerExporter,
            build_cert_batches,
        )

        self.save_current_history()
        rows = self._current_ndt_rows_for_export()
        date_str = self._selected_date_key()
        batches = build_cert_batches(
            rows,
            target_date=date_str,
            default_site='경인사무소',
            daily_usage_df=self._daily_usage_frame(),
        )
        if not batches:
            messagebox.showwarning(
                "시험성적서",
                "출력할 검사 행이 없습니다.\n업체·검사방법(MT/PT/UT/RT/PMI)을 입력한 뒤 다시 시도하세요.",
                parent=self,
            )
            return

        top = tk.Toplevel(self)
        top.title("시험성적서(갑/을) 출력")
        top.geometry("660x470")
        top.transient(self.winfo_toplevel())
        top.grab_set()

        body = ttk.Frame(top, padding=12)
        body.pack(fill='both', expand=True)

        ttk.Label(
            body,
            text=f"📄 {date_str} 시험성적서 생성 ({len(batches)}건 묶음)",
            font=('맑은 고딕', 11, 'bold'),
        ).pack(anchor='w', pady=(0, 8))

        tree = ttk.Treeview(
            body,
            columns=('company', 'method', 'count', 'report_no', 'customer', 'project'),
            show='headings',
            height=6,
        )
        for col, text, w in (
            ('company', '업체', 120),
            ('method', '방법', 55),
            ('count', '품목수', 55),
            ('report_no', '성적서번호', 180),
            ('customer', '발주처', 110),
            ('project', '공사명/현장', 110),
        ):
            tree.heading(col, text=text)
            tree.column(col, width=w, anchor='center' if col in ('method', 'count') else 'w')
        for idx, b in enumerate(batches):
            tree.insert(
                '',
                'end',
                iid=str(idx),
                values=(b['company'], b['method'], len(b['rows']), b['report_no'], b['customer'], b['project_name']),
            )
        tree.pack(fill='x', pady=(0, 10))

        edit_box = ttk.LabelFrame(body, text="선택 성적서 정보 보완 (클릭하여 수정)", padding=8)
        edit_box.pack(fill='x', pady=(0, 8))

        vars_map = {
            'report_no': tk.StringVar(),
            'customer': tk.StringVar(),
            'project_name': tk.StringVar(),
            'code_procedure': tk.StringVar(),
            'inspector': tk.StringVar(),
        }
        fields_meta = [
            ('성적서번호', 'report_no', 0, 0),
            ('발주처', 'customer', 0, 2),
            ('공사명/현장', 'project_name', 1, 0),
            ('적용규격/절차', 'code_procedure', 1, 2),
            ('검사자', 'inspector', 2, 0),
        ]
        for label_txt, key, r_i, c_i in fields_meta:
            ttk.Label(edit_box, text=label_txt).grid(row=r_i, column=c_i, sticky='w', padx=(4, 6), pady=3)
            ttk.Entry(edit_box, textvariable=vars_map[key], width=24).grid(
                row=r_i, column=c_i + 1, sticky='ew', padx=(0, 10), pady=3
            )
        edit_box.columnconfigure(1, weight=1)
        edit_box.columnconfigure(3, weight=1)

        selected_idx = [0]

        def load_form(idx):
            selected_idx[0] = idx
            b = batches[idx]
            for k, v in vars_map.items():
                v.set(b.get(k, ''))

        def save_form():
            b = batches[selected_idx[0]]
            for k, v in vars_map.items():
                b[k] = v.get().strip()
            tree.item(
                str(selected_idx[0]),
                values=(b['company'], b['method'], len(b['rows']), b['report_no'], b['customer'], b['project_name']),
            )

        ttk.Button(edit_box, text="목록에 적용", command=save_form).grid(row=2, column=3, sticky='e', pady=3)
        tree.selection_set('0')
        load_form(0)
        tree.bind('<<TreeviewSelect>>', lambda _e: tree.selection() and load_form(int(tree.selection()[0])))

        ledger_frame = ttk.Frame(body)
        ledger_frame.pack(fill='x', pady=(4, 8))
        append_ledger_var = tk.BooleanVar(value=False)
        ledger_path_var = tk.StringVar()
        ttk.Checkbutton(
            ledger_frame,
            text="생성 시 기존 발행대장(.xlsx)에도 자동 누적:",
            variable=append_ledger_var,
        ).pack(side='left')
        ttk.Entry(ledger_frame, textvariable=ledger_path_var, width=28).pack(side='left', padx=6, fill='x', expand=True)

        def browse_ledger():
            p = filedialog.askopenfilename(title="기존 성적서 발행대장 선택", filetypes=[("Excel", "*.xlsx")], parent=top)
            if p:
                ledger_path_var.set(p)
                append_ledger_var.set(True)

        ttk.Button(ledger_frame, text="...", width=3, command=browse_ledger).pack(side='left')

        def do_export():
            save_form()
            exporter = CertificateExporter()
            created = []
            try:
                if len(batches) == 1:
                    b = batches[0]
                    safe_no = b['report_no'].replace('/', '').replace('\\', '-').replace(':', '-')
                    out_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        initialfile=f"{safe_no}.xlsx",
                        title="시험성적서 저장",
                        filetypes=[("Excel files", "*.xlsx")],
                        parent=top,
                    )
                    if not out_path:
                        return
                    created.append(exporter.export_certificate(b, out_path))
                else:
                    folder = filedialog.askdirectory(title="성적서 일괄 저장 폴더 선택", parent=top)
                    if not folder:
                        return
                    for b in batches:
                        safe_no = b['report_no'].replace('/', '').replace('\\', '-').replace(':', '-')
                        target_file = os.path.join(folder, f"{safe_no}_{b['company']}.xlsx")
                        created.append(exporter.export_certificate(b, target_file))

                ledger_note = ""
                if append_ledger_var.get() and ledger_path_var.get().strip():
                    lpath = ledger_path_var.get().strip()
                    ledger_recs = []
                    for b in batches:
                        proj = b['project_name']
                        itm = b['item_name']
                        p_item = f"{proj} / {itm}" if proj not in ('', 'N/A') and itm not in ('', 'N/A') else (itm if itm not in ('', 'N/A') else proj)
                        ledger_recs.append({
                            'date': b['date'],
                            'company': b['company'],
                            'report_no': b['report_no'],
                            'customer': b['customer'],
                            'project_item': p_item,
                            'requester': b['requester'],
                            'issuer': b['inspector'],
                            'remark': f"{b['standard_short']} ({b['method']})",
                        })
                    CertificateLedgerExporter().export_ledger(
                        ledger_recs, lpath, existing_ledger_path=lpath
                    )
                    ledger_note = f"\n발행대장 누적 완료: {os.path.basename(lpath)}"

                messagebox.showinfo(
                    "성적서 생성 완료",
                    f"총 {len(created)}건의 시험성적서가 생성되었습니다.{ledger_note}",
                    parent=top,
                )
                os.startfile(str(created[0]) if len(created) == 1 else os.path.dirname(str(created[0])))
                top.destroy()
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("오류", f"시험성적서 생성 중 오류가 발생했습니다:\n{e}", parent=top)

        btn_bar = ttk.Frame(body)
        btn_bar.pack(fill='x', pady=(6, 0))
        ttk.Button(btn_bar, text="시험성적서 생성", command=do_export).pack(side='right', padx=4)
        ttk.Button(btn_bar, text="닫기", command=top.destroy).pack(side='right')

    def export_certificate_ledger(self):
        """작업일보 전체 이력을 바탕으로 업체별 성적서 발행대장을 생성/누적합니다."""
        from site_apps.gyeongin.src.cert_exporter import (
            CertificateLedgerExporter,
            collect_ledger_records_from_history,
        )

        self.save_current_history()
        history = self.load_history()
        all_records = collect_ledger_records_from_history(history, daily_usage_df=self._daily_usage_frame())
        if not all_records:
            messagebox.showwarning("발행대장", "발행대장에 반영할 검사 이력이 없습니다.", parent=self)
            return

        companies = sorted({r['company'] for r in all_records if r.get('company')})
        top = tk.Toplevel(self)
        top.title("성적서 발행대장 출력")
        top.geometry("480x240")
        top.transient(self.winfo_toplevel())
        top.grab_set()

        body = ttk.Frame(top, padding=14)
        body.pack(fill='both', expand=True)

        ttk.Label(body, text="📒 성적서 발행대장 생성 / 누적", font=('맑은 고딕', 11, 'bold')).pack(anchor='w', pady=(0, 10))

        row1 = ttk.Frame(body)
        row1.pack(fill='x', pady=4)
        ttk.Label(row1, text="업체 선택:", width=12).pack(side='left')
        comp_var = tk.StringVar(value='전체 업체')
        ttk.Combobox(row1, textvariable=comp_var, values=['전체 업체'] + companies, state='readonly', width=28).pack(side='left', fill='x', expand=True)

        row2 = ttk.Frame(body)
        row2.pack(fill='x', pady=4)
        ttk.Label(row2, text="기존 대장:", width=12).pack(side='left')
        existing_var = tk.StringVar()
        ttk.Entry(row2, textvariable=existing_var, width=28).pack(side='left', fill='x', expand=True)

        def pick_existing():
            p = filedialog.askopenfilename(title="기존 발행대장 선택 (미선택 시 새 파일 생성)", filetypes=[("Excel", "*.xlsx")], parent=top)
            if p:
                existing_var.set(p)

        ttk.Button(row2, text="...", width=3, command=pick_existing).pack(side='left', padx=(4, 0))

        ttk.Label(
            body,
            text="기존 파일을 선택하면 중복 발행번호를 건너뛰고 새 이력만 이어서 기록합니다.",
            foreground='#555555',
        ).pack(anchor='w', pady=(6, 12))

        def do_ledger():
            chosen = comp_var.get().strip()
            comp_filter = '' if chosen == '전체 업체' else chosen
            recs = [r for r in all_records if not comp_filter or r['company'] == comp_filter]
            if not recs:
                messagebox.showwarning("발행대장", f"'{chosen}'의 검사 기록이 없습니다.", parent=top)
                return

            existing = existing_var.get().strip() or None
            if existing:
                save_path = existing
            else:
                year_txt = str(recs[-1]['date'])[:4] or '2026'
                label = comp_filter or '경인'
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"{year_txt}년 {label} 성적서 발행대장.xlsx",
                    title="발행대장 저장",
                    filetypes=[("Excel", "*.xlsx")],
                    parent=top,
                )
                if not save_path:
                    return
            try:
                year_txt = str(recs[-1]['date'])[:4] or '2026'
                out = CertificateLedgerExporter().export_ledger(
                    recs,
                    save_path,
                    company_name=comp_filter,
                    year=year_txt,
                    existing_ledger_path=existing,
                )
                messagebox.showinfo("완료", f"성적서 발행대장이 저장되었습니다 ({len(recs)}건 반영).\n{out}", parent=top)
                os.startfile(str(out))
                top.destroy()
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("오류", f"발행대장 생성 중 오류가 발생했습니다:\n{e}", parent=top)

        btns = ttk.Frame(body)
        btns.pack(fill='x')
        ttk.Button(btns, text="발행대장 출력", command=do_ledger).pack(side='right', padx=4)
        ttk.Button(btns, text="닫기", command=top.destroy).pack(side='right')

    def export_monthly_report(self):
        # 1. 설정 다이얼로그
        top = tk.Toplevel(self)
        top.title("월간진도보고서 엑셀 출력 (년간계약)")
        # Allow enough vertical space for all fields and the generate button.
        top.geometry("620x480")
        top.minsize(620, 480)
        top.transient(self)
        top.grab_set()
        
        ttk.Label(top, text="📊 월간진도보고서 자동 생성 (년간계약 전용)", font=('맑은 고딕', 12, 'bold')).pack(pady=10)
        
        # 년간계약 업체 필터
        f_company = ttk.Frame(top)
        f_company.pack(pady=4, fill=tk.X, padx=10)
        ttk.Label(f_company, text="년간계약 업체:").pack(side=tk.LEFT, padx=5)
        history_companies = sorted({
            str(r.get('업체', '')).strip()
            for day in self.load_history().values()
            if isinstance(day, dict)
            for r in day.get('ndt_results', [])
            if str(r.get('업체', '')).strip()
        })
        comp_choices = ['전체 (미지정)'] + list(dict.fromkeys(self._configured_companies() + history_companies))
        company_var = tk.StringVar(value='전체 (미지정)')
        ttk.Combobox(f_company, textvariable=company_var, values=comp_choices, state='readonly', width=28).pack(side=tk.LEFT, padx=5)

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
                filetypes=[("Excel", "*.xlsx")],
                title="발주자 보고 파일 선택",
            )
            if path:
                owner_report_var.set(path)

        ttk.Button(
            f_owner, text="...", width=3, command=browse_owner_report
        ).pack(side=tk.LEFT)
        
        def do_generate():
            ym = f"{year_var.get()}-{month_var.get().zfill(2)}"
            tmpl_path = tmpl_var.get()
            owner_report_path = owner_report_var.get().strip()
            
            if not os.path.exists(tmpl_path):
                messagebox.showerror("오류", f"템플릿 파일을 찾을 수 없습니다.\n{tmpl_path}")
                return
            if owner_report_path and not os.path.exists(owner_report_path):
                messagebox.showerror(
                    "오류",
                    f"발주자 보고 파일을 찾을 수 없습니다.\n{owner_report_path}",
                )
                return
                
            chosen_comp = company_var.get().strip()
            comp_filter = '' if chosen_comp in ('', '전체 (미지정)') else chosen_comp
            file_tag = f"_{comp_filter}" if comp_filter else ""
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=(
                    f"월간진도보고서{file_tag}_{ym.replace('-', '')}_"
                    f"{'월간' if scope_var.get() == 'month' else '전체누적'}.xlsx"
                ),
                title="저장 위치 선택",
                filetypes=[("Excel", "*.xlsx")]
            )
            if not save_path: return
            
            try:
                # Assuming MonthlyReportManager is in the python path
                manager = MonthlyReportManager(tmpl_path)
                result_path = manager.generate_report(
                    self.history_path,
                    ym,
                    save_path,
                    doc_num=doc_var.get().strip(),
                    create_date=create_date_var.get().strip(),
                    owner_report_path=owner_report_path or None,
                    report_scope=scope_var.get(),
                    company_filter=comp_filter or None,
                )
                if result_path:
                    messagebox.showinfo("성공", f"월간진도보고서가 생성되었습니다.\n{result_path}")
                    os.startfile(result_path)
                    top.destroy()
                else:
                    period_text = (
                        f"{ym}의" if scope_var.get() == "month"
                        else f"{ym}까지의"
                    )
                    messagebox.showwarning("알림", f"{period_text} 작업일보 데이터가 없습니다.")
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("오류", f"생성 중 오류 발생:\n{e}")
                
        ttk.Button(top, text="보고서 생성", command=do_generate).pack(
            side=tk.BOTTOM, pady=18, ipadx=24, ipady=6
        )

    def load_history(self):
        try:
            with open(self.history_path, 'r', encoding='utf-8') as f:
                history = json.load(f)
            if not isinstance(history, dict) or any(not isinstance(day, dict) for day in history.values()):
                raise ValueError('일보 이력 형식이 올바르지 않습니다.')
            self._rt_parent_history = [row for day in sorted(history) for row in history[day].get('ndt_results', [])]
            return history
        except FileNotFoundError:
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

        if os.path.isfile(self.history_path):
            self.load_history() # Refuse to overwrite an unreadable existing history.
            shutil.copy2(self.history_path, self.history_path + '.backup.json')
        temp_path = self.history_path + '.tmp'
        with open(temp_path, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=4)
        os.replace(temp_path, self.history_path)
        self._refresh_ndt_history_choices(history)
        self._refresh_saved_date_markers(history)
        self._update_date_status(history)
        # Saving photo metadata does not persist other editor fields.

    def _config_file_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "Material_Manager_Config.json")

    def _configured_companies(self):
        names = []
        app_obj = getattr(self, "main_app", None) or getattr(self, "app", None)
        app_companies = getattr(app_obj, "companies", None)
        if isinstance(app_companies, list):
            names.extend(str(c).strip() for c in app_companies if str(c).strip())
        cfg_path = self._config_file_path()
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as cfg_file:
                    cfg_data = json.load(cfg_file)
                for item in cfg_data.get("companies", []):
                    val = str(item or "").strip()
                    if val and val.lower() != "nan":
                        names.append(val)
            except Exception:
                pass
        deduped = list(dict.fromkeys(names))
        if len(deduped) <= 1:
            deduped = list(dict.fromkeys(list(DEFAULT_COMPANIES) + deduped))
        return deduped

    @staticmethod
    def _format_company_choices(companies):
        cleaned = {str(c).strip() for c in companies if str(c or "").strip() and str(c).strip().lower() != "nan"}
        if "경인사무소" in cleaned:
            cleaned.remove("경인사무소")
            return ["", "경인사무소"] + sorted(cleaned)
        return [""] + sorted(cleaned)

    def _register_configured_companies(self, new_names):
        cleaned = [str(n or "").strip() for n in new_names if str(n or "").strip() and str(n or "").strip().lower() != "nan"]
        if not cleaned:
            return
        cfg_path = self._config_file_path()
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as cfg_file:
                    cfg_data = json.load(cfg_file)
                existing = [str(c).strip() for c in cfg_data.get("companies", []) if str(c).strip()]
                added = False
                for name in cleaned:
                    if name not in existing:
                        existing.append(name)
                        added = True
                if added:
                    cfg_data["companies"] = existing
                    with open(cfg_path, "w", encoding="utf-8") as cfg_file:
                        json.dump(cfg_data, cfg_file, ensure_ascii=False, indent=2)
            except Exception:
                pass
        app_obj = getattr(self, "main_app", None) or getattr(self, "app", None)
        if app_obj is not None and isinstance(getattr(app_obj, "companies", None), list):
            changed = False
            for name in cleaned:
                if name not in app_obj.companies:
                    app_obj.companies.append(name)
                    changed = True
            if changed and hasattr(app_obj, "refresh_ui_for_list_change"):
                app_obj.refresh_ui_for_list_change("companies")

    def _refresh_ndt_history_choices(self, history=None):
        """Refresh NDT combobox choices immediately after saving new rows."""
        if not hasattr(self, 'ndt_grid_entries'):
            return
        if history is None:
            history = self.load_history()

        sections = set()
        lines = set()
        companies = set(self._configured_companies())
        welders = set(self.welder_names)
        history_items = {target: set() for target in NDT_TARGET_TYPES}
        history_dimensions = {target: set() for target in NDT_TARGET_TYPES}
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
                target_type = infer_ndt_target(row)
                item_name = str(row.get('검사품명', '') or '').strip()
                if target_type in history_items and item_name:
                    history_items[target_type].add(item_name)
                dimension = str(row.get('치수원문', '') or '').strip()
                if target_type in history_dimensions and dimension:
                    history_dimensions[target_type].add(dimension)

        self.history_sections = [''] + sorted(sections)
        self.history_lines = [''] + sorted(lines)
        self.history_companies = self._format_company_choices(companies)
        self.history_welders = [''] + [
            self._format_welder_display(welder_id)
            for welder_id in sorted(welders)
        ]
        self.history_ndt_items = {
            target: sorted(items) for target, items in history_items.items()
        }
        self.history_ndt_dimensions = {
            target: sorted(values) for target, values in history_dimensions.items()
        }

        for row_entries in self.ndt_grid_entries:
            row_entries['구간'].configure(values=self.history_sections)
            row_entries['라인번호'].configure(values=self.history_lines)
            row_entries['업체'].configure(values=self.history_companies)
            row_entries['용접사'].configure(values=self.history_welders)
            self._sync_ndt_item_choices(row_entries)
            self._sync_ndt_dimension_choices(row_entries)

    def _refresh_saved_date_markers(self, history=None):
        """달력 팝업에서 작업일보가 저장된 날짜를 녹색으로 강조한다."""
        preview = getattr(self, 'calendar_work_summary', None)
        if preview is not None:
            preview.refresh(history=history)
            return
        if history is None:
            history = self.load_history()

        calendar = getattr(self.date_entry, '_calendar', None)
        if calendar is None:
            return

        try:
            calendar.calevent_remove('all')
            calendar.tag_config(
                'saved_work_log', background='#2E7D32', foreground='white'
            )
            for date_str in history:
                try:
                    saved_date = datetime.strptime(date_str, '%Y-%m-%d').date()
                except (TypeError, ValueError):
                    continue
                calendar.calevent_create(
                    saved_date, '작업일보 저장됨', 'saved_work_log'
                )
        except (AttributeError, tk.TclError):
            # tkcalendar 버전 차이가 있어도 날짜 선택 기능은 그대로 유지한다.
            pass

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
        if not self.refresh_company_summary():
            return False
        history = self.load_history()
        current_date = self._selected_date_key()
        if not self._validate_hot_rows():
            return
        existing_photos = history.get(current_date, {}).get('process_photos', [])
        
        data = {
            'weather': self.weather_entry.get(),
            'qty_data': history.get(current_date, {}).get('qty_data', {}),
            'summary_start_date': self.company_start_date.get(),
            'company_summary': self.company_summary_rows,
            'equip_data': history.get(current_date, {}).get('equip_data', {}),
            'personnel_data': history.get(current_date, {}).get('personnel_data', {}),
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'company_memos': DailyWorkLogTab._company_memo_snapshot(self),
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
            
        with open('debug_save.txt', 'w', encoding='utf-8') as debug_f:
            import json
            json.dump(data['ndt_results'], debug_f, ensure_ascii=False, indent=2)
            
        history[current_date] = data
        self.save_history(history)
        if hasattr(self, '_register_configured_companies'):
            self._register_configured_companies(row.get('업체', '') for row in data['ndt_results'])
        self._loaded_log_date = current_date
        self._log_baseline = self._capture_log_input()

    def _capture_log_input(self):
        """Capture editable values without saving or including the date selector."""
        return {
            'weather': self.weather_entry.get(),
            'remarks': self.remarks_text.get('1.0', 'end-1c'),
            'company_memos': DailyWorkLogTab._company_memo_snapshot(self),
            'qty': {key: {field: widget.get() for field, widget in entries.items()}
                    for key, entries in self.qty_entries.items()},
            'equipment': {key: {field: widget.get() for field, widget in entries.items()}
                          for key, entries in self.equip_entries.items()},
            'people': {key: widget.get() for key, widget in self.personnel_entries.items()},
            'ndt': [{key: widget.get() for key, widget in row.items() if hasattr(widget, 'get')}
                    for row in self.ndt_grid_entries],
        }

    def on_date_change(self, event=None):
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
        history = corrected_rt_history(loaded_history)
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
            
        if curr_data.get('summary_start_date'):
            self.company_start_date.set_date(curr_data['summary_start_date'])
            self._summary_start_initialized = True

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
            
        self._load_company_memos(curr_data.get('company_memos', {}))

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
            for col, ent in row_entries.items():
                if col == '규격':
                    val = normalized_spec
                elif col == '근무구분':
                    val = work_shift
                elif col == '대상유형':
                    val = infer_ndt_target(row_data)
                elif col == '용접사':
                    val = self._format_welder_display(row_data.get(col, ''))
                else:
                    val = row_data.get(col, '')
                if col == '구간정보':
                    ent.set(val)
                elif isinstance(ent, tk.ttk.Combobox):
                    ent.set(val)
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
                    ent.insert(0, val)
                    
        for widgets in self.ndt_grid_entries:
            self._sync_ndt_item_choices(widgets)
            self._sync_ndt_dimension_choices(widgets)
            self._sync_ndt_spec(widgets)
            try:
                sync_rt_widgets(widgets, all_rows=self._rt_context())
            except ValueError:
                pass  # Keep the saved inputs available for correction before saving.
        self._refresh_hot_summary()

        self.refresh_company_summary()

        self._loaded_log_date = current_date
        self._log_baseline = self._capture_log_input()

    def auto_calculate_and_save(self, show_completion=True):
        if not self._validate_rt_rows():
            return
        if not self._validate_hot_rows():
            return
        self._refresh_hot_summary()
        if not self.refresh_company_summary():
            return False

        # Save to history
        if self.save_current_history() is False:
            return False
        
        if show_completion:
            messagebox.showinfo("완료", "자동 집계 및 데이터 저장이 완료되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test Tab")
    root.geometry("1400x800")
    tab = DailyWorkLogTab(root)
    tab.pack(fill="both", expand=True)
    root.mainloop()

