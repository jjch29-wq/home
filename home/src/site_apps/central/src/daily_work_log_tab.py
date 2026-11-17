import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from site_apps.central.src.ndt_summary_exporter import NDTSummaryExporter
from site_apps.central.src.services.monthly_report_manager import MonthlyReportManager
from tkcalendar import DateEntry
from datetime import datetime
import os
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
from site_apps.central.src.rt_quantities import rt_counts, sync_rt_widgets, widget_rt_values, resolve_rt_row, retest_parent, default_rt_copy, corrected_rt_history

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
        self._left_pane_visible = True
        self._left_pane_width = None

        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=5, pady=(4, 0))
        self.left_pane_toggle = ttk.Button(
            toolbar, text='◀ 좌측 숨기기', command=self._toggle_left_pane,
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
        
        def _update_left_scroll(e=None):
            bbox = self.left_canvas.bbox("all")
            if bbox:
                self.left_canvas.configure(scrollregion=(0, 0, bbox[2], max(bbox[3], self.left_canvas.winfo_height())))
        self.left_frame.bind("<Configure>", _update_left_scroll)
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
        
        def _on_left_canvas_configure(event):
            new_width = max(event.width, self.left_frame.winfo_reqwidth())
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
            if total_width > 200 and self._left_pane_visible:
                self.paned.sashpos(0, total_width // 2)
        except (tk.TclError, AttributeError):
            pass

    def _toggle_left_pane(self):
        if self._left_pane_visible:
            self._left_pane_width = self.paned.sashpos(0)
            self.paned.forget(self.left_container)
            self._left_pane_visible = False
            self.left_pane_toggle.configure(text='▶ 좌측 보이기')
        else:
            self.paned.insert(0, self.left_container, weight=4)
            self._left_pane_visible = True
            self.paned.update_idletasks()
            width = self.paned.winfo_width()
            position = self._left_pane_width if self._left_pane_width is not None else width // 2
            self.paned.sashpos(0, max(0, min(position, width - 1)))
            self.left_pane_toggle.configure(text='◀ 좌측 숨기기')

    def _build_basic_info(self, parent):
        # --- Top Section: General Info & Unified Work Toolbar ---
        top_frame = ttk.LabelFrame(parent, text="기본 정보 및 통합 작업 도구", padding=6)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        # 1. 좌측: 기본 정보 입력 및 저장
        left_box = ttk.LabelFrame(top_frame, text="기본 정보", padding=(6, 3))
        left_box.pack(side="left", fill="y", padx=(0, 6))
        
        ttk.Label(left_box, text="검사일자:").grid(row=0, column=0, padx=(0, 4), pady=2, sticky="e")
        self.date_entry = DateEntry(left_box, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=(0, 6), pady=2, sticky="w")
        self.date_entry.bind("<<DateEntrySelected>>", self.on_date_change)
        self.date_entry.bind("<Return>", self.on_date_change)
        from site_apps.central.src.views.calendar_work_summary import CalendarWorkSummary
        self.calendar_work_summary = CalendarWorkSummary(
            self.date_entry, self,
            site_getter=lambda: '중앙지사',
            frame_getter=lambda: getattr(getattr(self, 'main_app', None), 'daily_usage_df', None),
            history_getter=self.load_history)

        self.date_status_label = tk.Label(
            left_box, text="○ 저장되지 않은 날짜", fg="#777777",
            font=('맑은 고딕', 9, 'bold')
        )
        self.date_status_label.grid(row=1, column=0, columnspan=2, padx=(0, 6), sticky="w")
        
        ttk.Label(left_box, text="날씨:").grid(row=0, column=2, padx=(0, 4), pady=2, sticky="e")
        self.weather_entry = ttk.Entry(left_box, width=8)
        self.weather_entry.grid(row=0, column=3, padx=(0, 6), pady=2, sticky="w")
        
        btn_calc = ttk.Button(left_box, text="💾 자동 집계/저장", command=self.auto_calculate_and_save)
        btn_calc.grid(row=0, column=4, padx=(0, 3), pady=2)
        
        btn_delete = ttk.Button(left_box, text="🗑️ 삭제", command=self.delete_current_date_data)
        btn_delete.grid(row=0, column=5, padx=0, pady=2)
        
        btn_export = ttk.Button(left_box, text="📊 엑셀 출력 (일보)", command=self.export_excel)
        btn_export.grid(row=1, column=4, columnspan=2, padx=0, pady=(1, 0), sticky="ew")
        
        # 2. 중앙: 보고서 발행
        mid_box = ttk.LabelFrame(top_frame, text="보고서 발행", padding=(6, 3))
        mid_box.pack(side="left", fill="y", padx=(0, 6))
        
        btn_export_ndt = ttk.Button(mid_box, text="📗 NDT 누계대장", command=self.export_ndt_summary)
        btn_export_ndt.grid(row=0, column=0, padx=3, pady=2, sticky="ew")
        
        btn_export_monthly = ttk.Button(mid_box, text="📈 월간진도보고서", command=self.export_monthly_report)
        btn_export_monthly.grid(row=0, column=1, padx=3, pady=2, sticky="ew")
        
        btn_weekly_report = ttk.Button(mid_box, text="🗓️ 주간업무보고", command=lambda: self.main_app.export_weekly_report() if hasattr(self, 'main_app') else None)
        btn_weekly_report.grid(row=1, column=0, columnspan=2, padx=3, pady=(1, 0), sticky="ew")

        # 3. 우측: 비파괴검사결과서 편집 도구 (우측 빈 공간 100% 활용)
        right_box = ttk.LabelFrame(top_frame, text="비파괴검사결과서 편집", padding=(6, 3))
        right_box.pack(side="left", fill="both", expand=True)

        ttk.Button(right_box, text="➕ 10행 추가", command=lambda: self._append_ndt_rows(10, focus=True)).grid(row=0, column=0, padx=3, pady=2, sticky="w")
        ttk.Button(right_box, text="🗑️ 선택 행 삭제", command=self._delete_selected_ndt_row).grid(row=0, column=1, padx=3, pady=2, sticky="w")
        ttk.Button(right_box, text="✏️ 라인번호 일괄", command=self._edit_ndt_lines).grid(row=0, column=2, padx=3, pady=2, sticky="w")
        ttk.Button(right_box, text="✏️ 구간 일괄", command=self._edit_ndt_sections).grid(row=0, column=3, padx=3, pady=2, sticky="w")
        ttk.Button(right_box, text="🔍 작업내역 검색·수정", command=self.open_history_search_dialog).grid(row=0, column=4, padx=3, pady=2, sticky="w")

        self.ndt_selection_vars = []
        self.ndt_selection_text = tk.StringVar(value='왼쪽 체크박스로 여러 행 선택 가능')
        ttk.Label(right_box, textvariable=self.ndt_selection_text, foreground="#666666").grid(row=1, column=0, columnspan=5, padx=4, pady=2, sticky="w")

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
        eq_frame = ttk.LabelFrame(bot_frame, text="장비투입 현황", padding=10)
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
        pe_frame = ttk.LabelFrame(bot_frame, text="금일 투입 인원", padding=10)
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
        rm_frame = ttk.LabelFrame(bot_frame, text="특이사항 및 계획", padding=10)
        rm_frame.pack(side="left", fill="both", expand=True, padx=5)
        memo_toolbar = ttk.Frame(rm_frame)
        memo_toolbar.pack(fill="x", pady=(0, 5))
        ttk.Label(memo_toolbar, text="선택한 검사일자의 메모").pack(side="left")
        ttk.Button(memo_toolbar, text="메모 이력", command=self.open_memo_history).pack(side="right")
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

    def _build_right_pane(self, parent):
        # --- NDT Results Grid ---
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", padx=5, pady=(4, 2))
        ttk.Label(title_frame, text="2. 비파괴검사결과서", font=("맑은 고딕", 11, "bold")).pack(side="left", padx=(5, 10))
        ttk.Label(
            title_frame,
            text="※ 행 추가/삭제/일괄변경 및 보고서 발행 도구는 상단 툴바에 통합되어 있습니다.",
            foreground="#666666"
        ).pack(side="left", padx=5)
        
        # Grid Container
        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.ndt_cols = ('업체', '검사방법', '구간', '작업구분', '핫태핑개소', '핫태핑차수', '핫태핑개소수', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분',
                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
        history = self.load_history()
        sections = set()
        lines = set()
        companies = set()
        # 월간 보고서와 동일한 용접사 번호/성명 기준정보를 사용한다.
        welders = set(self.welder_names)
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
                if r.get('업체'): companies.add(r['업체'].strip())
                if r.get('용접사'): welders.add(self._normalize_welder_id(r['용접사']))
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
        self.history_companies = [''] + sorted(list(companies))
        self.history_welders = [''] + [
            self._format_welder_display(welder_id)
            for welder_id in sorted(welders)
        ]

        
        # Draw Headers
        ttk.Label(grid_frame, text='선택').grid(row=0, column=0, padx=4)
        for col_idx, c in enumerate(self.ndt_cols, start=1):
            label = {'RT_OR': 'RT_OR 전체', 'RT_RE': 'RT_RE 불량'}.get(c, c)
            ttk.Label(grid_frame, text=label, font=("맑은 고딕", 9, "bold")).grid(row=0, column=col_idx, padx=1, pady=2)
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
                from site_apps.central.src.hot_tapping import inspection_length
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


        def on_hot_change(event=None, r=row_entries):
            if r['작업구분'].get() == '핫태핑':
                self._edit_hot_tapping(r)
        row_entries['작업구분'].bind('<<ComboboxSelected>>', on_hot_change)
        row_entries['검사방법'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['검사방법'].bind('<FocusOut>', on_ndt_change)
        row_entries['관경'].bind('<<ComboboxSelected>>', on_ndt_change)
        row_entries['관경'].bind('<FocusOut>', on_ndt_change)
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

            for col in ['업체', '검사방법', '구간', '작업구분', '핫태핑개소', '핫태핑차수', '핫태핑개소수', '라인번호', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분']:
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

        def _show_ndt_cell_context_menu(event, r=row_entries):
            self._select_ndt_row(r)
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="➕ 10행 추가", command=lambda: self._append_ndt_rows(10, focus=True))
            menu.add_command(label="🗑️ 선택 행 삭제", command=self._delete_selected_ndt_row)
            menu.add_separator()
            menu.add_command(label="✏️ 라인번호 일괄 변경...", command=self._edit_ndt_lines)
            menu.add_command(label="✏️ 구간 일괄 변경...", command=self._edit_ndt_sections)
            menu.add_command(label="🔍 작업내역 검색·수정...", command=self.open_history_search_dialog)
            menu.add_separator()
            menu.add_command(label="💾 자동 집계 및 저장", command=self.auto_calculate_and_save)
            menu.add_command(label="📊 엑셀 출력 (일보)", command=self.export_excel)
            menu.add_command(label="📗 NDT 누계대장 엑셀 출력", command=self.export_ndt_summary)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        for col_name, widget in row_entries.items():
            if hasattr(widget, 'entries'):
                for e in widget.entries:
                    e.bind('<Return>', on_joint_enter)
                    e.bind('<Button-1>', lambda event, r=row_entries: self._select_ndt_row(r), add='+')
                    e.bind('<Button-3>', lambda event, r=row_entries: _show_ndt_cell_context_menu(event, r), add='+')
                    e.bind('<Escape>', self._clear_ndt_selection, add='+')
            else:
                widget.bind('<Return>', on_joint_enter)
                widget.bind('<Button-1>', lambda event, r=row_entries: self._select_ndt_row(r), add='+')
                widget.bind('<Button-3>', lambda event, r=row_entries: _show_ndt_cell_context_menu(event, r), add='+')
                widget.bind('<Escape>', self._clear_ndt_selection, add='+')

        self.ndt_grid_entries.append(row_entries)


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

    def open_history_search_dialog(self):
        """저장된 전체 작업내역에서 일자·업체·라인번호별로 검색하고,
        원하는 행을 선택하여 해당 일자의 결과서로 이동해 수정하거나 현재 일보로 복사해옵니다."""
        if hasattr(self, '_history_search_window') and self._history_search_window is not None:
            try:
                if self._history_search_window.winfo_exists():
                    self._history_search_window.deiconify()
                    self._history_search_window.lift()
                    self._history_search_window.focus_force()
                    return
            except Exception:
                pass

        history = self.load_history()
        all_records = []
        company_set = set()
        line_set = set()
        method_set = set()

        for date_key in sorted(history.keys(), reverse=True):
            day_data = history[date_key]
            for row_idx, r in enumerate(day_data.get('ndt_results', [])):
                if not any(str(r.get(k, '')).strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법', '구간')):
                    continue
                co = str(r.get('업체', '')).strip()
                ln = str(r.get('라인번호', '')).strip()
                mth = str(r.get('검사방법', '')).strip().upper()
                if co: company_set.add(co)
                if ln: line_set.add(ln)
                if mth: method_set.add(mth)
                all_records.append({
                    'date': str(date_key),
                    'row_idx': row_idx,
                    'company': co,
                    'method': mth,
                    'section': str(r.get('구간', '')).strip(),
                    'work_type': str(r.get('작업구분', '')).strip(),
                    'hot_loc': str(r.get('핫태핑개소', '')).strip(),
                    'hot_deg': str(r.get('핫태핑차수', '')).strip(),
                    'hot_cnt': str(r.get('핫태핑개소수', '')).strip(),
                    'line_no': ln,
                    'joint_no': str(r.get('Joint No.', '')).strip(),
                    'pipe': str(r.get('관경', '')).strip(),
                    'thick': str(r.get('두께', '')).strip(),
                    'welder': str(r.get('용접사', '')).strip(),
                    'joint_info': str(r.get('구간정보', '')).strip(),
                    'result': str(r.get('결과', '')).strip(),
                    'spec': str(r.get('규격', '')).strip(),
                    'shift': str(r.get('근무구분', '')).strip(),
                    'raw': dict(r),
                })

        window = tk.Toplevel(self)
        self._history_search_window = window

        def _on_window_close():
            self._history_search_window = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", _on_window_close)
        window.title("작업내역 검색 및 결과서 불러오기 (일자 · 업체 · 라인번호)")
        window.geometry("1120x660")
        window.minsize(880, 500)
        window.resizable(True, True)
        window.transient(self.winfo_toplevel())

        # 화면 최대화 / 원래 크기 복원 토글 함수
        def toggle_maximize(_event=None):
            try:
                is_zoomed = (window.state() == 'zoomed')
                if is_zoomed:
                    window.state('normal')
                    btn_text = "⛶ 화면 최대화"
                else:
                    window.state('zoomed')
                    btn_text = "🗗 원래 크기로"
                if 'top_max_btn' in locals() and top_max_btn.winfo_exists():
                    top_max_btn.config(text=btn_text)
                if 'bottom_max_btn' in locals() and bottom_max_btn.winfo_exists():
                    bottom_max_btn.config(text=btn_text)
            except Exception:
                pass

        window.bind("<F11>", toggle_maximize)

        # 1. 상단 필터 바
        filter_box = ttk.LabelFrame(window, text="검색 및 필터 조건", padding=8)
        filter_box.pack(fill="x", padx=10, pady=(10, 6))

        # 일자 필터
        row1 = ttk.Frame(filter_box)
        row1.pack(fill="x", pady=2)

        ttk.Label(row1, text="검사일자:").pack(side="left")
        all_dates_var = tk.BooleanVar(value=True)

        try:
            from datetime import timedelta
            today_dt = datetime.strptime(self._selected_date_key(), '%Y-%m-%d').date()
        except Exception:
            today_dt = datetime.now().date()
        month_ago_dt = today_dt - timedelta(days=30)

        start_date_entry = DateEntry(row1, width=11, background='darkblue', foreground='white',
                                     date_pattern='yyyy-mm-dd', year=month_ago_dt.year,
                                     month=month_ago_dt.month, day=month_ago_dt.day)
        start_date_entry.pack(side="left", padx=4)

        ttk.Label(row1, text="~").pack(side="left")
        end_date_entry = DateEntry(row1, width=11, background='darkblue', foreground='white',
                                   date_pattern='yyyy-mm-dd', year=today_dt.year,
                                   month=today_dt.month, day=today_dt.day)
        end_date_entry.pack(side="left", padx=4)

        ttk.Checkbutton(row1, text="전체 기간(제한 없음)", variable=all_dates_var, command=lambda: refresh_rows()).pack(side="left", padx=(6, 16))

        ttk.Label(row1, text="업체:").pack(side="left")
        company_var = tk.StringVar(value="(전체)")
        company_combo = ttk.Combobox(row1, textvariable=company_var, values=['(전체)'] + sorted(company_set), width=14, state='readonly')
        company_combo.pack(side="left", padx=(4, 16))

        ttk.Label(row1, text="방법:").pack(side="left")
        method_var = tk.StringVar(value="(전체)")
        method_combo = ttk.Combobox(row1, textvariable=method_var, values=['(전체)'] + sorted(method_set), width=8, state='readonly')
        method_combo.pack(side="left", padx=(4, 16))

        # 라인번호 및 키워드 검색
        row2 = ttk.Frame(filter_box)
        row2.pack(fill="x", pady=(4, 2))

        ttk.Label(row2, text="라인번호:").pack(side="left")
        line_var = tk.StringVar(value="(전체)")
        line_combo = ttk.Combobox(row2, textvariable=line_var, values=['(전체)'] + sorted(line_set), width=26)
        line_combo.pack(side="left", padx=(4, 16))

        ttk.Label(row2, text="통합 검색어:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(row2, textvariable=search_var, width=18)
        search_entry.pack(side="left", padx=(4, 8))
        ttk.Label(row2, text="(Joint, 구간, 용접사 등)", foreground="#666666").pack(side="left", padx=(0, 16))

        ttk.Button(row2, text="🔍 조회", command=lambda: refresh_rows()).pack(side="left", padx=4)
        
        def reset_filters():
            all_dates_var.set(True)
            company_var.set("(전체)")
            method_var.set("(전체)")
            line_var.set("(전체)")
            line_combo['values'] = ['(전체)'] + sorted(line_set)
            search_var.set("")
            refresh_rows()

        ttk.Button(row2, text="초기화", command=reset_filters).pack(side="left", padx=2)

        top_max_btn = ttk.Button(row2, text="⛶ 화면 최대화", command=toggle_maximize)
        top_max_btn.pack(side="right", padx=2)

        # 업체 선택 변경 시 해당 업체의 라인번호로 동적 연동
        def on_company_change(*_args):
            sel_co = company_var.get().strip()
            if not sel_co or sel_co == '(전체)':
                line_combo['values'] = ['(전체)'] + sorted(line_set)
            else:
                co_lines = sorted({r['line_no'] for r in all_records if r['company'] == sel_co and r['line_no']})
                line_combo['values'] = ['(전체)'] + co_lines
            line_var.set('(전체)')
            refresh_rows()

        company_combo.bind('<<ComboboxSelected>>', on_company_change)
        method_combo.bind('<<ComboboxSelected>>', lambda _e: refresh_rows())
        line_combo.bind('<<ComboboxSelected>>', lambda _e: refresh_rows())
        line_combo.bind('<Return>', lambda _e: refresh_rows())
        search_entry.bind('<Return>', lambda _e: refresh_rows())

        # 2. 중앙 테이블
        table_frame = ttk.Frame(window, padding=(10, 4, 10, 4))
        table_frame.pack(fill="both", expand=True)

        cols = ('date', 'company', 'method', 'section', 'line_no', 'joint_no', 'pipe', 'thick', 'welder', 'result', 'spec')
        tree = ttk.Treeview(table_frame, columns=cols, show='headings', selectmode='extended')
        
        headers = {
            'date': ('검사일자', 90, 'center'),
            'company': ('업체', 110, 'w'),
            'method': ('방법', 60, 'center'),
            'section': ('구간', 80, 'center'),
            'line_no': ('라인번호', 220, 'w'),
            'joint_no': ('Joint No.', 85, 'center'),
            'pipe': ('관경', 65, 'center'),
            'thick': ('두께', 55, 'center'),
            'welder': ('용접사', 85, 'center'),
            'result': ('결과', 60, 'center'),
            'spec': ('규격/근무', 85, 'center'),
        }
        for col_id, (header_text, col_width, align) in headers.items():
            tree.heading(col_id, text=header_text)
            tree.column(col_id, width=col_width, minwidth=40, anchor=align)

        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        # 3. 하단 상태 및 액션 버튼 바
        status_var = tk.StringVar()
        ttk.Label(window, textvariable=status_var, padding=(12, 2, 12, 4), foreground="#1d4ed8").pack(fill="x")

        bottom_frame = ttk.Frame(window, padding=(10, 4, 10, 8))
        bottom_frame.pack(fill="x")

        # 1단: 가이드 문구 + 이동/복사 시 창 닫기 체크박스
        guide_row = ttk.Frame(bottom_frame)
        guide_row.pack(fill="x", pady=(0, 6))

        ttk.Label(
            guide_row,
            text="💡 항목 더블클릭 또는 [이동하여 수정]을 누르면 해당 일자의 결과서로 이동해 편집할 수 있습니다.",
            foreground="#555555"
        ).pack(side="left")

        close_on_move_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(guide_row, text="이동/복사 시 창 닫기", variable=close_on_move_var).pack(side="right")

        # 2단: 화면 최대화 버튼(좌측) + 액션 버튼들(우측)
        action_row = ttk.Frame(bottom_frame)
        action_row.pack(fill="x")

        bottom_max_btn = ttk.Button(action_row, text="⛶ 화면 최대화", command=toggle_maximize)
        bottom_max_btn.pack(side="left")

        filtered_records = []

        def refresh_rows(*_args):
            nonlocal filtered_records
            tree.delete(*tree.get_children())
            
            sel_co = company_var.get().strip()
            sel_mth = method_var.get().strip()
            sel_line = line_var.get().strip()
            query = search_var.get().strip().lower()
            all_dates = all_dates_var.get()

            try:
                start_str = start_date_entry.get_date().strftime('%Y-%m-%d')
                end_str = end_date_entry.get_date().strftime('%Y-%m-%d')
            except Exception:
                start_str, end_str = "0000-00-00", "9999-99-99"

            filtered = []
            for r in all_records:
                if not all_dates:
                    if not (start_str <= r['date'] <= end_str):
                        continue
                if sel_co and sel_co != '(전체)' and r['company'] != sel_co:
                    continue
                if sel_mth and sel_mth != '(전체)' and r['method'] != sel_mth:
                    continue
                if sel_line and sel_line != '(전체)' and sel_line.lower() not in r['line_no'].lower():
                    continue
                if query:
                    search_corpus = f"{r['date']} {r['company']} {r['line_no']} {r['joint_no']} {r['section']} {r['welder']} {r['joint_info']}".lower()
                    if query not in search_corpus:
                        continue
                filtered.append(r)

            filtered_records = filtered
            for idx, r in enumerate(filtered):
                spec_display = f"{r['spec']} ({r['shift']})" if r['shift'] else r['spec']
                tree.insert(
                    '', 'end', iid=str(idx),
                    values=(r['date'], r['company'], r['method'], r['section'], r['line_no'],
                            r['joint_no'], r['pipe'], r['thick'], r['welder'], r['result'], spec_display)
                )

            status_var.set(f"총 {len(filtered)}건의 작업내역이 조회되었습니다. (선택 후 버튼 클릭)")

        def move_to_selected(_event=None):
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("선택 확인", "이동할 작업내역을 목록에서 먼저 선택하세요.", parent=window)
                return
            idx = int(sel[0])
            record = filtered_records[idx]
            target_date = record['date']
            target_line = record['line_no']
            target_joint = record['joint_no']

            current_date = self._selected_date_key()
            if target_date != current_date:
                baseline = getattr(self, '_log_baseline', None)
                if baseline is not None and self._capture_log_input() != baseline:
                    if not messagebox.askyesno(
                        '일보 날짜 이동',
                        f'현재 {current_date} 일보에 저장하지 않은 작성 내용이 있습니다.\n'
                        f'작성 내용을 버리고 {target_date} 일보로 이동할까요?',
                        parent=window
                    ):
                        return
                try:
                    self.date_entry.set_date(datetime.strptime(target_date, '%Y-%m-%d').date())
                except Exception:
                    pass
                self.on_date_change()

            target_row_idx = None
            for r_i, row in enumerate(self.ndt_grid_entries):
                if (row['라인번호'].get().strip() == target_line and
                    row['Joint No.'].get().strip() == target_joint and
                    row['업체'].get().strip() == record['company']):
                    target_row_idx = r_i
                    break
            if target_row_idx is None:
                for r_i, row in enumerate(self.ndt_grid_entries):
                    if (row['라인번호'].get().strip() == target_line and
                        row['Joint No.'].get().strip() == target_joint):
                        target_row_idx = r_i
                        break

            if target_row_idx is not None:
                self._clear_ndt_selection()
                self._select_ndt_row(self.ndt_grid_entries[target_row_idx])
                self._focus_ndt_joint(target_row_idx)

            if close_on_move_var.get():
                _on_window_close()
                if target_row_idx is not None:
                    messagebox.showinfo(
                        "작업내역 이동 완료",
                        f"[{target_date}] {record['company']} - {target_line} (Joint {target_joint})\n"
                        f"해당 행({target_row_idx + 1}행)을 선택했습니다. 수정 후 상단 [💾 자동 집계/저장]을 눌러 주세요.",
                        parent=self
                    )
            else:
                if target_row_idx is not None:
                    status_var.set(
                        f"✅ [{target_date}] {record['company']} · {target_line} · Joint {target_joint} "
                        f"-> 결과서 {target_row_idx + 1}행 선택 및 포커스 완료! 메인 화면에서 바로 수정 후 [💾 자동 집계/저장]하세요."
                    )
                else:
                    status_var.set(
                        f"⚠️ [{target_date}] 일보로 이동했으나 결과서에서 일치하는 행을 찾지 못했습니다."
                    )

        def copy_to_current():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("선택 확인", "가져올 작업내역을 1개 이상 선택하세요.", parent=window)
                return
            items_to_copy = [filtered_records[int(i)] for i in sel]

            empty_indices = [
                idx for idx, row in enumerate(self.ndt_grid_entries)
                if not any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법'))
            ]

            needed = len(items_to_copy) - len(empty_indices)
            if needed > 0:
                self._append_ndt_rows(((needed + 9) // 10) * 10)
                empty_indices = [
                    idx for idx, row in enumerate(self.ndt_grid_entries)
                    if not any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법'))
                ]

            first_filled_idx = None
            for item, row_idx in zip(items_to_copy, empty_indices):
                row = self.ndt_grid_entries[row_idx]
                if first_filled_idx is None:
                    first_filled_idx = row_idx
                field_map = {
                    '업체': item['company'],
                    '검사방법': item['method'],
                    '구간': item['section'],
                    '작업구분': item['work_type'],
                    '핫태핑개소': item['hot_loc'],
                    '핫태핑차수': item['hot_deg'],
                    '핫태핑개소수': item['hot_cnt'],
                    '라인번호': item['line_no'],
                    'Joint No.': item['joint_no'],
                    '관경': item['pipe'],
                    '두께': item['thick'],
                    '용접사': item['welder'],
                    '구간정보': item['joint_info'],
                    '결과': item['result'] or '합격',
                    '규격': item['spec'],
                    '근무구분': item['shift'] or '주간',
                }
                for k, v in field_map.items():
                    if k in row:
                        if hasattr(row[k], 'set'):
                            row[k].set(v)
                        elif hasattr(row[k], 'delete'):
                            row[k].delete(0, tk.END)
                            row[k].insert(0, v)
                self._sync_ndt_spec(row)

            if first_filled_idx is not None:
                self._clear_ndt_selection()
                self._select_ndt_row(self.ndt_grid_entries[first_filled_idx])
                self._focus_ndt_joint(first_filled_idx)

            if close_on_move_var.get():
                _on_window_close()
                messagebox.showinfo(
                    "가져오기 완료",
                    f"{len(items_to_copy)}건의 작업내역을 현재 일보({self._selected_date_key()}) 결과서로 가져왔습니다.\n"
                    "수정 또는 확인 후 상단 [💾 자동 집계/저장]을 눌러 주세요.",
                    parent=self
                )
            else:
                status_var.set(
                    f"✅ {len(items_to_copy)}건의 작업내역을 현재 일보({self._selected_date_key()}) 결과서로 복사했습니다. "
                    f"메인 화면에서 확인 후 [💾 자동 집계/저장]하세요."
                )

        tree.bind('<Double-1>', move_to_selected)
        tree.bind('<Return>', move_to_selected)

        ttk.Button(action_row, text="닫기", command=_on_window_close, width=10).pack(side="right", padx=(4, 0))
        ttk.Button(action_row, text="📋 현재 일보로 행 복사하여 가져오기", command=copy_to_current).pack(side="right", padx=4)
        ttk.Button(action_row, text="📌 해당 일자로 이동하여 수정", command=move_to_selected).pack(side="right", padx=4)

        refresh_rows()
        search_entry.focus_set()

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
        self.hot_summary_var.set(f"핫태핑: {len(locations - {''})}개소명 / MT {totals['MT']:.4f} m / PT {totals['PT']:.4f} m")

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

    def save_current_history(self):
        date_str = self.date_entry.get()
        if not date_str: return
        history = self.load_history()
        history[date_str] = self.get_current_data()
        self.save_history(history)
        
    def delete_current_date_data(self):
        """현재 선택된 날짜의 데이터를 삭제하고 화면을 초기화합니다."""
        date_str = self.date_entry.get()
        if not date_str: return

        history = self.load_history()
        day_data = history.get(date_str, {})
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

    def export_monthly_report(self):
        # 1. 설정 다이얼로그
        top = tk.Toplevel(self)
        top.title("월간진도보고서 엑셀 출력")
        # Allow enough vertical space for all fields and the generate button.
        top.geometry("620x430")
        top.minsize(620, 430)
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
                
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=(
                    f"월간진도보고서_{ym.replace('-', '')}_"
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

    def _refresh_ndt_history_choices(self, history=None):
        """Refresh NDT combobox choices immediately after saving new rows."""
        if not hasattr(self, 'ndt_grid_entries'):
            return
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

        self.history_sections = [''] + sorted(sections)
        self.history_lines = [''] + sorted(lines)
        self.history_companies = [''] + sorted(companies)
        self.history_welders = [''] + [
            self._format_welder_display(welder_id)
            for welder_id in sorted(welders)
        ]

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
        history = self.load_history()
        current_date = self._selected_date_key()
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
            
        with open('debug_save.txt', 'w', encoding='utf-8') as debug_f:
            import json
            json.dump(data['ndt_results'], debug_f, ensure_ascii=False, indent=2)
            
        history[current_date] = data
        self.save_history(history)
        self._loaded_log_date = current_date
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
            for col, ent in row_entries.items():
                if col == '규격':
                    val = normalized_spec
                elif col == '근무구분':
                    val = work_shift
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
            self._sync_ndt_spec(widgets)
            try:
                sync_rt_widgets(widgets, all_rows=self._rt_context())
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

        self._loaded_log_date = current_date
        self._log_baseline = self._capture_log_input()

    def auto_calculate_and_save(self, show_completion=True):
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

