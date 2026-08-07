import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update UI
old_ui = """        ttk.Label(pe_frame, text="구분").grid(row=0, column=0)
        ttk.Label(pe_frame, text="검사원").grid(row=0, column=1)
        
        self.personnel_entries = {}
        for i, lbl in enumerate(['인원', '현장대리인', '누계'], start=1):
            ttk.Label(pe_frame, text=lbl).grid(row=i, column=0, sticky="w")
            ent = ttk.Entry(pe_frame, width=10)
            ent.grid(row=i, column=1, padx=2, pady=2)
            self.personnel_entries[f'검사원_{lbl}'] = ent"""

new_ui = """        ttk.Label(pe_frame, text="구분").grid(row=0, column=0)
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
            self.personnel_entries[f'안전_{lbl}'] = ent2"""
code = code.replace(old_ui, new_ui)

# 2. Update logic
old_logic = """        # Personnel
        prev_p_total_str = prev_data.get('personnel_data', {}).get('검사원_누계', '0') or '0'
        try: prev_p_total = float(prev_p_total_str.replace(',',''))
        except ValueError: prev_p_total = 0.0
        
        today_p = float(self.personnel_entries['검사원_인원'].get() or 0) + float(self.personnel_entries['검사원_현장대리인'].get() or 0)
        total_p = prev_p_total + today_p
        
        self.personnel_entries['검사원_누계'].delete(0, tk.END)
        self.personnel_entries['검사원_누계'].insert(0, f"{total_p:.1f}" if total_p % 1 else f"{int(total_p)}")"""

new_logic = """        # Personnel
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
        self.personnel_entries['안전_누계'].insert(0, f"{total_s:.1f}" if total_s % 1 else f"{int(total_s)}")"""
code = code.replace(old_logic, new_logic)

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated daily_work_log_tab.py")
