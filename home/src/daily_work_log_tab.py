import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from ndt_summary_exporter import NDTSummaryExporter
from tkcalendar import DateEntry
from datetime import datetime
import os
import sys
import json

# Import the exporter
try:
    from daily_work_log_exporter import DailyWorkLogExporter
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

class DailyWorkLogTab(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.history_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'daily_work_history.json')
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
        
        self.left_frame.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
        
        def _on_left_canvas_configure(event):
            new_width = max(event.width, self.left_frame.winfo_reqwidth())
            self.left_canvas.itemconfig(self.left_window, width=new_width)
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
        
        self.right_frame.bind("<Configure>", lambda e: self.right_canvas.configure(scrollregion=self.right_canvas.bbox("all")))
        self.right_window = self.right_canvas.create_window((0, 0), window=self.right_frame, anchor="nw")
        self.right_canvas.configure(xscrollcommand=self.right_xscroll.set, yscrollcommand=self.right_yscroll.set)
        
        def _on_right_canvas_configure(event):
            # Only stretch if the canvas is wider than the required width
            if event.width > self.right_frame.winfo_reqwidth():
                self.right_canvas.itemconfig(self.right_window, width=event.width)
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
        
        btn_export = ttk.Button(top_frame, text="엑셀 출력 (일보 생성)", command=self.export_excel)
        btn_export.grid(row=0, column=5, padx=10, pady=5)
        
        # --- Middle Section: Work Quantities ---
        mid_frame = ttk.LabelFrame(parent, text="1. 작업 물량 및 누계 현황", padding=10)
        mid_frame.pack(fill="x", padx=5, pady=5)
        
        headers = ['방법', '규격', '예상량', '전일누계', '금일작업', '총누계', '공정률(%)', '불량', '불량률(%)', '비고']
        for col, h in enumerate(headers):
            ttk.Label(mid_frame, text=h, font=('맑은 고딕', 9, 'bold')).grid(row=0, column=col, padx=1, pady=2)
            
        self.qty_rows = [
            ('PAUT', '300A이상'), ('PAUT', '300A이상-야간'), ('PAUT', '250A'), ('PAUT', '200A'), ('PAUT', '200A-야간'), ('PAUT', '소계'),
            ('RT', '150A~100A'), ('RT', '150A~100A-야간'), ('RT', '80A이하'), ('RT', '80A이하-야간'), ('RT', '소계'),
            ('MT', '전체(주간)'), ('MT', '전체(야간)'),
            ('PT', '전체(주간)'), ('PT', '전체(야간)')
        ]
        
        self.qty_entries = {}
        for row_idx, (method, spec) in enumerate(self.qty_rows, start=1):
            ttk.Label(mid_frame, text=method).grid(row=row_idx, column=0, padx=1, pady=2)
            ttk.Label(mid_frame, text=spec).grid(row=row_idx, column=1, padx=1, pady=2)
            
            default_qty = {
                ('PAUT', '300A이상'): '121', ('PAUT', '300A이상-야간'): '584', 
                ('PAUT', '250A'): '4', ('PAUT', '200A'): '4', ('PAUT', '200A-야간'): '2', ('PAUT', '소계'): '715',
                ('RT', '150A~100A'): '293', ('RT', '150A~100A-야간'): '43',
                ('RT', '80A이하'): '105', ('RT', '80A이하-야간'): '49', ('RT', '소계'): '490',
                ('MT', '전체(주간)'): '26', ('MT', '전체(야간)'): '0',
                ('PT', '전체(주간)'): '26', ('PT', '전체(야간)'): '0'
            }
            
            row_dict = {}
            for col_idx, key in enumerate(['예상량', '전일누계', '금일작업', '총누계', '공정률', '불량', '불량률', '비고'], start=2):
                ent = ttk.Entry(mid_frame, width=7)
                ent.grid(row=row_idx, column=col_idx, padx=1, pady=2)
                if key == '예상량':
                    val = default_qty.get((method, spec), '')
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
        self.remarks_text = tk.Text(rm_frame, width=20, height=8)
        self.remarks_text.pack(fill="both", expand=True)

    def _build_right_pane(self, parent):
        # --- NDT Results Grid ---
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", pady=5)
        ttk.Label(title_frame, text="2. 비파괴검사결과서 (연속 30줄 입력 가능)", font=("맑은 고딕", 12, "bold")).pack(side="left", padx=10)
        
        btn_export_ndt = ttk.Button(title_frame, text="NDT 누계 대장 엑셀 출력", command=self.export_ndt_summary)
        btn_export_ndt.pack(side="left", padx=10)
        
        # Grid Container
        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.ndt_cols = ('업체', '검사방법', '구간', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격', 
                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
        history = self.load_history()
        sections = set()
        lines = set()
        companies = set()
        for date_str, data in history.items():
            for r in data.get('ndt_results', []):
                if r.get('구간'): sections.add(r['구간'].strip())
                if r.get('라인번호'): lines.add(r['라인번호'].strip())
                if r.get('업체'): companies.add(r['업체'].strip())
        
        self.history_sections = [''] + sorted(list(sections))
        self.history_lines = [''] + sorted(list(lines))
        self.history_companies = [''] + sorted(list(companies))

        
        # Draw Headers
        for col_idx, c in enumerate(self.ndt_cols):
            ttk.Label(grid_frame, text=c, font=("맑은 고딕", 9, "bold")).grid(row=0, column=col_idx, padx=1, pady=2)
            if c in ('구간정보', '라인번호'):
                grid_frame.grid_columnconfigure(col_idx, weight=3)
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
                if c in ('검사방법', '결과', '규격', '관경'): w = 6
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
                elif c == '용접사':
                    ent = ttk.Combobox(grid_frame, width=w, values=['', 'W-2023-A-10', 'W-2023-A-13', 'W-2023-A-25'], justify='center')
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
                        r['규격']['values'] = ['', '31/3 X12"주간', '31/3 X12"야간', '31/3 X6"주간', '31/3 X6"야간']
                    elif method in ['PAUT', 'MT', 'PT']:
                        r['규격']['values'] = ['', '주간', '야간', '재검']
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
                
                for col in ['업체', '검사방법', '구간', '라인번호', '관경', '용접사', '결과', '규격']:
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
                else:
                    widget.bind('<Return>', on_joint_enter)
            
            self.ndt_grid_entries.append(row_entries)
            
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
        
        data = {
            'weather': self.weather_entry.get(),
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {},
            'remarks': self.remarks_text.get("1.0", "end-1c"),
            'ndt_results': []
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

        for comp_key, entries in self.qty_entries.items():
            # Load from today if exists
            curr_qty = curr_data.get('qty_data', {}).get(comp_key, {})
            for field in ['예상량', '금일작업', '총누계', '공정률', '불량', '불량률', '비고']:
                entries[field].delete(0, tk.END)
                if field in curr_qty:
                    val = load_format_val(comp_key, field, curr_qty[field])
                    entries[field].insert(0, val)
                    
            # Always calculate 전일누계 from past
            prev_total = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0')
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, prev_total)
            
        # Update Equip
        for eq, entries in self.equip_entries.items():
            curr_eq = curr_data.get('equip_data', {}).get(eq, {})
            for field in ['금일', '누계']:
                entries[field].delete(0, tk.END)
                if field in curr_eq:
                    entries[field].insert(0, curr_eq[field])
                    
        # Update Personnel
        for p_key, ent in self.personnel_entries.items():
            ent.delete(0, tk.END)
            if p_key in curr_data.get('personnel_data', {}):
                ent.insert(0, curr_data['personnel_data'][p_key])
                
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
        for i, row_entries in enumerate(self.ndt_grid_entries):
            row_data = curr_ndt[i] if i < len(curr_ndt) else {}
            for col, ent in row_entries.items():
                val = row_data.get(col, '')
                if col == '구간정보':
                    ent.set(val)
                elif isinstance(ent, tk.ttk.Combobox):
                    ent.set(val)
                elif hasattr(ent, 'delete'):
                    ent.delete(0, tk.END)
                    ent.insert(0, val)

    def auto_calculate_and_save(self):
        current_date = self.date_entry.get()
        history = self.load_history()
        
        # 1. Aggregate NDT Results -> 금일작업
        today_qty = {comp_key: 0.0 for comp_key in self.qty_entries.keys()}
        
        for row_entries in self.ndt_grid_entries:
            if not hasattr(row_entries['검사방법'], 'get'): continue
            method = row_entries['검사방법'].get().upper().strip()
            if not method: continue
            
            size_str = row_entries['관경'].get().strip() if hasattr(row_entries['관경'], 'get') else ""
            spec_str = row_entries['규격'].get().strip() if hasattr(row_entries['규격'], 'get') else ""
            
            import re
            size_match = re.search(r'\d+', size_str)
            size_val = int(size_match.group()) if size_match else 0
            
            if method == 'RT':
                val = float(row_entries['RT_OR'].get() or 0) + float(row_entries['RT_RE'].get() or 0)
                if val > 0:
                    spec_key = '80A이하' if size_val <= 80 else '150A~100A'
                    if spec_str == '야간': spec_key += '-야간'
                    comp = f"RT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
            
            elif method == 'PAUT':
                val = float(row_entries['PAUT'].get() or 0)
                if val > 0:
                    spec_key = '200A'
                    if size_val >= 300: spec_key = '300A이상'
                    elif size_val == 250: spec_key = '250A'
                    if spec_str == '야간' and spec_key in ['300A이상', '200A']:
                        spec_key += '-야간'
                    comp = f"PAUT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
                        
            elif method == 'MT':
                val = float(row_entries['MT'].get() or 0)
                if val > 0:
                    spec_key = '전체(야간)' if spec_str == '야간' else '전체(주간)'
                    comp = f"MT_{spec_key}"
                    if comp in today_qty: today_qty[comp] += val
                    
            elif method == 'PT':
                val = float(row_entries['PT'].get() or 0)
                if val > 0:
                    spec_key = '전체(야간)' if spec_str == '야간' else '전체(주간)'
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
            entries['금일작업'].insert(0, format_val(comp_key, today_val))
            
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, format_val(comp_key, prev_total))
            
            total = prev_total + today_val
            entries['총누계'].delete(0, tk.END)
            entries['총누계'].insert(0, format_val(comp_key, total))
            
            expected = float(entries['예상량'].get() or 0)
            if expected > 0:
                entries['예상량'].delete(0, tk.END)
                entries['예상량'].insert(0, f"{int(expected)}")
                progress = (total / expected) * 100
                entries['공정률'].delete(0, tk.END)
                entries['공정률'].insert(0, f"{progress:.1f}")

        # Subtotals for Qty
        paut_expected = sum(float(self.qty_entries[f"PAUT_{s}"]['예상량'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        if paut_expected > 0:
            self.qty_entries['PAUT_소계']['예상량'].delete(0, tk.END)
            self.qty_entries['PAUT_소계']['예상량'].insert(0, f"{int(paut_expected)}")
        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].insert(0, format_val('PAUT', paut_total))
        
        rt_expected = sum(float(self.qty_entries[f"RT_{s}"]['예상량'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        if rt_expected > 0:
            self.qty_entries['RT_소계']['예상량'].delete(0, tk.END)
            self.qty_entries['RT_소계']['예상량'].insert(0, f"{int(rt_expected)}")
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].insert(0, format_val('RT', rt_total))
                
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

