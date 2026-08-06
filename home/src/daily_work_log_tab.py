import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
import os
import sys

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

    def _build_left_pane(self, parent):
        # --- Top Section: General Info ---
        top_frame = ttk.LabelFrame(parent, text="기본 정보", padding=10)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(top_frame, text="검사일자:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.date_entry = DateEntry(top_frame, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(top_frame, text="날씨:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.weather_entry = ttk.Entry(top_frame, width=15)
        self.weather_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        btn_export = ttk.Button(top_frame, text="엑셀 출력 (일보 생성)", command=self.export_excel)
        btn_export.grid(row=0, column=4, padx=20, pady=5)
        
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
                ('RT', '150A~100A'): '95', ('RT', '150A~100A-야간'): '14',
                ('RT', '80A이하'): '34', ('RT', '80A이하-야간'): '16', ('RT', '소계'): '159',
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
            self.qty_entries[spec] = row_dict
            
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
        
        self.personnel_entries = {}
        for i, lbl in enumerate(['인원', '현장대리인', '누계'], start=1):
            ttk.Label(pe_frame, text=lbl).grid(row=i, column=0, sticky="w")
            ent = ttk.Entry(pe_frame, width=10)
            ent.grid(row=i, column=1, padx=2, pady=2)
            self.personnel_entries[f'검사원_{lbl}'] = ent
            
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
        
        # Grid Container
        grid_frame = ttk.Frame(parent)
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.ndt_cols = ('검사방법', '구간', '라인번호', 'Joint No.', '관경', '용접사', '구간정보', '결과', '규격', 
                         'RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
        
        # Draw Headers
        for col_idx, c in enumerate(self.ndt_cols):
            ttk.Label(grid_frame, text=c, font=("맑은 고딕", 9, "bold")).grid(row=0, column=col_idx, padx=1, pady=2)
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
                if c in ('검사방법', '구간', '결과', '규격', '관경', '용접사'): w = 6
                elif c in ('라인번호', 'Joint No.', '구간정보'): w = 12
                else: w = 5
                
                if c == '검사방법':
                    ent = ttk.Combobox(grid_frame, width=w, values=['', 'RT', 'PAUT', 'UT', 'MT', 'PT', 'PMI', 'ETC'])
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                elif c == '구간정보':
                    frame = ttk.Frame(grid_frame)
                    frame.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    entries = []
                    for i in range(4):
                        e = ttk.Entry(frame, width=3)
                        e.pack(side='left', fill='x', expand=True, padx=(0, 1 if i<3 else 0))
                        entries.append(e)
                    frame.entries = entries
                    def get_val(ents=entries, r_dict=row_entries):
                        method = r_dict['검사방법'].get().strip().upper() if '검사방법' in r_dict else ''
                        if method in ['PAUT', 'PT', 'MT']:
                            return ents[0].get()
                        parts = [e.get().strip() for e in ents]
                        return ','.join([p for p in parts if p])
                    def set_val(val, ents=entries):
                        for e in ents: e.delete(0, tk.END)
                        ents[0].insert(0, val)
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
                    ent = ttk.Combobox(grid_frame, width=w, values=[''] + list(SIZE_LENGTH.keys()))
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                elif c == '결과':
                    ent = ttk.Combobox(grid_frame, width=w, values=['', '합격', '불합격', '재촬영', '보류'])
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                elif c == '규격':
                    ent = ttk.Combobox(grid_frame, width=w, values=[''])
                    ent.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="ew")
                    row_entries[c] = ent
                else:
                    ent = ttk.Entry(grid_frame, width=w)
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
                        for i in range(1, 4):
                            r['구간정보'].entries[i].pack_forget()
                    else:
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
                
                for col in ['검사방법', '구간', '라인번호', '관경', '용접사', '규격']:
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
        for spec, entries in self.qty_entries.items():
            data['qty_data'][spec] = {k: v.get() for k, v in entries.items()}
            
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

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Test Tab")
    root.geometry("1400x800")
    tab = DailyWorkLogTab(root)
    tab.pack(fill="both", expand=True)
    root.mainloop()
