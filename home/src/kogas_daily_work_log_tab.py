import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from ndt_summary_exporter import NDTSummaryExporter
from services.kogas_monthly_report_manager import KogasMonthlyReportManager
from tkcalendar import DateEntry
from datetime import datetime
import os
import sys
import json
import shutil
import uuid

# Import the exporter
try:
    from kogas_daily_work_log_exporter import KogasDailyWorkLogExporter
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

class KogasDailyWorkLogTab(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.history_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'kogas_daily_work_history.json')
        self.photo_root = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'data', 'process_photos'
        )
        self.selected_ndt_row = None
        self.setup_ui()
        
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
            new_width = max(event.width, self.left_frame.winfo_reqwidth())
            self.left_canvas.itemconfig(self.left_window, width=new_width)
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
            # Only stretch if the canvas is wider than the required width
            if event.width > self.right_frame.winfo_reqwidth():
                self.right_canvas.itemconfig(self.right_window, width=event.width)
            _update_right_scroll()
        self.right_canvas.bind("<Configure>", _on_right_canvas_configure)
        
        self.right_canvas.grid(row=0, column=0, sticky="nsew")
        self.right_yscroll.grid(row=0, column=1, sticky="ns")
        self.right_xscroll.grid(row=1, column=0, sticky="ew")
        self.right_container.grid_rowconfigure(0, weight=1)
        self.right_container.grid_columnconfigure(0, weight=1)
        
        self._build_right_pane(self.right_frame)
        
        # Initial load for today's date
        self.after(100, self.on_date_change)

    def _build_left_pane(self, parent):
        # --- Top Section: General Info ---
        top_frame = ttk.LabelFrame(parent, text="기본 정보", padding=10)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(top_frame, text="검사일자:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.date_entry = DateEntry(top_frame, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.date_entry.bind("<<DateEntrySelected>>", self.on_date_change)
        # self.date_entry.bind("<FocusOut>", self.on_date_change) # Removed to prevent accidental UI wipes
        
        ttk.Label(top_frame, text="날씨:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.weather_entry = ttk.Entry(top_frame, width=15)
        self.weather_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        btn_calc = ttk.Button(top_frame, text="자동 집계 및 저장", command=self.auto_calculate_and_save)
        btn_calc.grid(row=0, column=4, padx=5, pady=5)
        
        # [NEW] 삭제 버튼 추가
        btn_delete = ttk.Button(top_frame, text="🗑️ 데이터 삭제", command=self.delete_current_date_data)
        btn_delete.grid(row=0, column=5, padx=5, pady=5)
        
        btn_export = ttk.Button(top_frame, text="엑셀 출력 (일보 생성)", command=self.export_excel)
        btn_export.grid(row=0, column=6, padx=10, pady=5)
        
        # --- Middle Section: Work Quantities ---
        mid_frame = ttk.LabelFrame(parent, text="1. 작업 물량 및 누계 현황", padding=10)
        mid_frame.pack(fill="x", padx=5, pady=5)
        
        headers = ['방법', '규격', '예상량', '전일누계', '금일작업', '총누계', '공정률(%)', '불량', '불량률(%)', '비고']
        for col, h in enumerate(headers):
            ttk.Label(mid_frame, text=h, font=('맑은 고딕', 9, 'bold')).grid(row=0, column=col, padx=1, pady=2)
            
        self.qty_rows = [
            ('RT', 'B필름: 3⅓"x17"'),
            ('RT', 'A필름: 3⅓"x12"'),
            ('RT', 'A/2필름: 3⅓"x6"'),
            ('RT', '소계'),
            ('UT', '초음파탐상'),
            ('PT', '침투탐상'),
            ('MT', '자분탐상')
        ]
        
        self.qty_entries = {}
        self.default_qty = {
            ('RT', 'B필름: 3⅓"x17"'): '20368',
            ('RT', 'A필름: 3⅓"x12"'): '2464',
            ('RT', 'A/2필름: 3⅓"x6"'): '1704',
            ('RT', '소계'): '24536',
            ('UT', '초음파탐상'): '319.02',
            ('PT', '침투탐상'): '338.63',
            ('MT', '자분탐상'): '0'
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
        
        self.equip_rows = ['RT장비(선원)', 'UT장비', 'MT/PT장비']
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
        self.remarks_text = tk.Text(rm_frame, width=20, height=8)
        self.remarks_text.pack(fill="both", expand=True)

    def _build_right_pane(self, parent):
        # --- NDT Results Grid ---
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", pady=5)
        ttk.Label(title_frame, text="2. 비파괴검사결과서 (연속 30줄 입력 가능)", font=("맑은 고딕", 12, "bold")).pack(side="left", padx=10)
        
        btn_export_ndt = ttk.Button(title_frame, text="NDT 누계 대장 엑셀 출력", command=self.export_ndt_summary)
        btn_export_ndt.pack(side="left", padx=10)
        
        btn_export_monthly = ttk.Button(title_frame, text="누적진도보고서 출력", command=self.export_monthly_report)
        btn_export_monthly.pack(side="left", padx=10)
        
        btn_weekly_report = ttk.Button(title_frame, text="🗓️ 주간업무보고", command=lambda: self.main_app.export_weekly_report() if hasattr(self, 'main_app') else None)
        btn_weekly_report.pack(side="left", padx=10)
        
        # Grid Container
        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '두께', '용접사', '구간정보', '결과', '규격', '근무구분',
                         'RT_OR', 'RT_RE', 'UT', 'MT', 'PT')
        history = self.load_history()
        sections = set()
        lines = set()
        companies = set()
        welders = {'W-2023-A-10', 'W-2023-A-13', 'W-2023-A-25'}
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
                if r.get('업체'): companies.add(r['업체'].strip())
                if r.get('용접사'): welders.add(r['용접사'].strip())
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
        self.history_companies = [''] + sorted(list(companies))
        self.history_welders = [''] + sorted(welders)

        
        # Draw Headers
        for col_idx, c in enumerate(self.ndt_cols):
            ttk.Label(grid_frame, text=c, font=("맑은 고딕", 9, "bold")).grid(row=0, column=col_idx, padx=1, pady=2)
            if c in ('구간정보', '라인번호'):
                grid_frame.grid_columnconfigure(col_idx, weight=3)
            elif c == '규격':
                grid_frame.grid_columnconfigure(col_idx, weight=2, minsize=105)
            else:
                grid_frame.grid_columnconfigure(col_idx, weight=1)
            
        # Draw 30 Rows of Entries
        self.ndt_grid_entries = []
        for row_idx in range(1, 31):
            row_entries = {}
            # Row number label
            ttk.Label(grid_frame, text=f"{row_idx}").grid(row=row_idx, column=0, sticky="w", padx=(0, 2))
            
            for col_idx, c in enumerate(self.ndt_cols):
                # Adjust width for some columns
                w = 8
                if c == '규격': w = 13
                elif c == '근무구분': w = 7
                elif c in ('검사방법', '결과', '관경', '두께'): w = 6
                elif c in ('구간', '업체'): w = 10
                elif c == '용접사': w = 15
                elif c == '라인번호': w = 25
                elif c == 'Joint No.': w = 12
                elif c == '구간정보': w = 20
                else: w = 8
                
                if c == '검사방법':
                    ent = ttk.Combobox(grid_frame, width=w, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'], justify='center')
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                elif c == '구간정보':
                    frame = ttk.Frame(grid_frame)
                    frame.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    entries = []
                    for i in range(4):
                        e = ttk.Entry(frame, width=3, justify='center')
                        e.pack(side='left', fill='x', expand=True, padx=(0, 1 if i<3 else 0))
                        entries.append(e)
                    frame.entries = entries
                    def get_val(ents=entries, r_dict=row_entries):
                        method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                        if method in ['PAUT', 'PT', 'MT']:
                            return ents[0].get()
                        parts = [e.get().strip() for e in ents]
                        return ','.join([p for p in parts if p])
                    def set_val(val, ents=entries, r_dict=row_entries):
                        for e in ents: e.delete(0, tk.END)
                        method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                        if method in ['PAUT', 'PT', 'MT']:
                            ents[0].insert(0, val)
                        else:
                            parts = val.split(',') if val else []
                            for i, p in enumerate(parts):
                                if i < len(ents): ents[i].insert(0, p)
                    def delete_val(first, last, ents=entries):
                        for e in ents: e.delete(first, last)
                    def insert_val(idx, val, ents=entries):
                        ents[0].insert(idx, val)
                    frame.get = get_val
                    frame.set = set_val
                    frame.delete = delete_val
                    frame.insert = insert_val
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
                
                # --- 관경 선택 시 두께 자동 입력 ---
                thk_mapping = {
                    "1100": "11.1", "1000": "11.1", "900": "10.3", "850": "10.3", "800": "9.5",
                    "750": "8.7", "700": "8.7", "650": "8.7", "600": "9.5", "550": "9.5",
                    "500": "6.4", "450": "6.4", "400": "6.4", "350": "6.4", "300": "6.4", "250": "6.4", "200": "6.4", "150": "4.5", "100": "4.5", "80": "4.5"
                }
                lookup_size = str(pipe).replace("A", "").strip()
                if lookup_size in thk_mapping:
                    r['두께'].delete(0, tk.END)
                    r['두께'].insert(0, thk_mapping[lookup_size])
                # -----------------------------------
                
                if pipe and pipe.isdigit():
                    pipe += 'A'
                
                if pipe in SIZE_LENGTH:
                    length = SIZE_LENGTH[pipe]
                    length_str = f"{length:.4f}"
                    
                    if method == 'RT':
                        try:
                            size_val = int(''.join(filter(str.isdigit, pipe)))
                            if size_val <= 150:
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
                    if method in ['PAUT', 'PT', 'MT']:
                        r['구간정보'].entries[0].configure(justify='left')
                        for i in range(1, 4):
                            r['구간정보'].entries[i].pack_forget()
                    else:
                        r['구간정보'].entries[0].configure(justify='center')
                        for i in range(1, 4):
                            r['구간정보'].entries[i].pack(side='left', fill='x', expand=True, padx=(0, 1 if i<3 else 0))
                            
                # Update '규격' values based on method
                if '규격' in r and hasattr(r['규격'], 'configure'):
                    if method == 'RT':
                        r['규격']['values'] = ['', '3 1/3 X 12"', '3 1/3 X 6"']
                    elif method in ['PAUT', 'MT', 'PT']:
                        r['규격']['values'] = ['']
                    else:
                        r['규격']['values'] = ['']
                        
            row_entries['검사방법'].bind('<<ComboboxSelected>>', on_ndt_change)
            row_entries['검사방법'].bind('<FocusOut>', on_ndt_change)
            row_entries['관경'].bind('<<ComboboxSelected>>', on_ndt_change)
            row_entries['관경'].bind('<FocusOut>', on_ndt_change)
            
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
                
                for col in ['업체', '검사방법', '구간', '라인번호', '관경', '두께', '용접사', '결과', '규격']:
                    if not r[col].get().strip():
                        prev_val = prev_row[col].get().strip()
                        if prev_val:
                            if isinstance(r[col], ttk.Combobox):
                                r[col].set(prev_val)
                            else:
                                r[col].delete(0, tk.END)
                                r[col].insert(0, prev_val)
                
                on_ndt_change(None, r)
                
                if r_idx + 1 < len(self.ndt_grid_entries):
                    self.ndt_grid_entries[r_idx + 1]['Joint No.'].focus_set()
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

        photo_bar = ttk.LabelFrame(parent, text="선택 행 공정사진")
        photo_bar.pack(fill='x', padx=5, pady=(8, 5))
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

    def _select_ndt_row(self, row_entries):
        self.selected_ndt_row = row_entries
        method = row_entries['검사방법'].get().strip() or '-'
        section = row_entries['구간'].get().strip() or '-'
        joint = row_entries['Joint No.'].get().strip() or '-'
        welder = row_entries['용접사'].get().strip() or '-'
        self.photo_selection_var.set(
            f"선택: {method} / {section} / Joint {joint} / {welder}"
        )

    def _clear_ndt_selection(self, event=None):
        self.selected_ndt_row = None
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
                'welder': row['용접사'].get().strip(),
                'location': row['구간'].get().strip() or joint,
                'description': description.strip(),
                'file_path': os.path.relpath(target_path, os.path.dirname(self.history_path)),
            })
        self.save_history(history)
        messagebox.showinfo('완료', f'공정사진 {len(files)}장을 등록했습니다.')

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
            if any(val.strip() for val in row_dict.values()):
                data['ndt_results'].append(row_dict)
            
        self.save_current_history()
        # Save File Dialog
        default_name = f"{self.date_entry.get().replace('-', '')}_작업일보.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            title="일보 저장",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                exporter = KogasDailyWorkLogExporter()
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
        top.title("누적 진도보고서 엑셀 출력")
        # Allow enough vertical space for all fields and the generate button.
        top.geometry("620x430")
        top.minsize(620, 430)
        top.transient(self)
        top.grab_set()
        
        ttk.Label(top, text="📊 누적 진도보고서 자동 생성", font=('맑은 고딕', 12, 'bold')).pack(pady=10)
        
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
                    f"누적진도보고서_{ym.replace('-', '')}_"
                    f"{'월간' if scope_var.get() == 'month' else '전체누적'}.xlsx"
                ),
                title="저장 위치 선택",
                filetypes=[("Excel", "*.xlsx")]
            )
            if not save_path: return
            
            try:
                # Assuming MonthlyReportManager is in the python path
                manager = KogasMonthlyReportManager(tmpl_path)
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
                    messagebox.showinfo("성공", f"누적 진도보고서가 생성되었습니다.\n{result_path}")
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
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def save_history(self, history):
        with open(self.history_path, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=4)

    def save_current_history(self):
        history = self.load_history()
        current_date = self.date_entry.get()
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
                    row_dict[col] = str(val) if val else ""
            data['ndt_results'].append(row_dict)
            
        with open('debug_save.txt', 'w', encoding='utf-8') as debug_f:
            import json
            json.dump(data['ndt_results'], debug_f, ensure_ascii=False, indent=2)
            
        history[current_date] = data
        self.save_history(history)

    def on_date_change(self, event=None):
        current_date = self.date_entry.get()
        history = self.load_history()
        
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
                if field == '예상량' and not str(current_value).strip():
                    previous_value = (
                        prev_data.get('qty_data', {})
                        .get(comp_key, {})
                        .get('예상량', '')
                    )
                    if str(previous_value).strip():
                        current_value = previous_value
                    else:
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
                        expected_value = curr_qty.get('예상량', '')
                        if not str(expected_value).strip():
                            expected_value = prev_qty.get('예상량', '')
                        if not str(expected_value).strip():
                            method, spec = comp_key.split('_', 1)
                            expected_value = self.default_qty.get((method, spec), '')
                        try:
                            total_num = float(str(previous_total).replace(',', ''))
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
                else:
                    val = row_data.get(col, '')
                if col == '구간정보':
                    ent.set(val)
                elif isinstance(ent, tk.ttk.Combobox):
                    ent.set(val)
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
                    ent.insert(0, val)
            return f"{v:.1f}" if v % 1 else f"{int(v)}"
        today_qty = {comp_key: 0.0 for comp_key in self.qty_entries.keys()}
        
        for row_entries in self.ndt_grid_entries:
            if not hasattr(row_entries['검사방법'], 'get'): continue
            method = row_entries['검사방법'].get().upper().strip()
            if not method: continue
            
            spec_str = row_entries['규격'].get().strip() if hasattr(row_entries['규격'], 'get') else ""
            
            if method == 'RT':
                val = float(row_entries['RT_OR'].get() or 0) + float(row_entries['RT_RE'].get() or 0)
                if val > 0:
                    if '17' in spec_str or 'B' in spec_str.upper():
                        spec_key = 'B필름: 3⅓"x17"'
                    elif '12' in spec_str:
                        spec_key = 'A필름: 3⅓"x12"'
                    elif '6' in spec_str:
                        spec_key = 'A/2필름: 3⅓"x6"'
                    else:
                        spec_key = 'B필름: 3⅓"x17"' # Default
                    comp = f"RT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
            
            elif method == 'UT':
                val = float(row_entries['UT'].get() or 0)
                if val > 0:
                    comp = "UT_초음파탐상"
                    if comp in today_qty: today_qty[comp] += val
                        
            elif method == 'MT':
                val = float(row_entries['MT'].get() or 0)
                if val > 0:
                    comp = "MT_자분탐상"
                    if comp in today_qty: today_qty[comp] += val
                    
            elif method == 'PT':
                val = float(row_entries['PT'].get() or 0)
                if val > 0:
                    comp = "PT_침투탐상"
                    if comp in today_qty: today_qty[comp] += val
                    
        # Subtotals
        today_qty['RT_소계'] = sum([v for k, v in today_qty.items() if k.startswith('RT_') and k != 'RT_소계'])
                    
        # Update 금일작업 in UI
        def format_val(ckey, v):
            if ckey.startswith(('UT', 'PT')):
                return f"{v:.4f}"
            return f"{v:.1f}" if v % 1 else f"{int(v)}"

        for comp_key, val in today_qty.items():
            if '소계' not in comp_key or comp_key == 'RT_소계':
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                if val != 0:
                    ent.insert(0, format_val(comp_key, val))
                
        # 2. Update Totals based on previous date        # 2. Update Totals based on previous date
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
        paut_keys = ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간']
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
        
        messagebox.showinfo("완료", "자동 집계 및 데이터 저장이 완료되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test Tab")
    root.geometry("1400x800")
    tab = DailyWorkLogTab(root)
    tab.pack(fill="both", expand=True)
    root.mainloop()

