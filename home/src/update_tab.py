import re
import os

with open('daily_work_log_tab.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Add json and history_path
if "import json" not in code:
    code = code.replace(
        "import sys",
        "import sys\nimport json"
    )

if "self.history_path =" not in code:
    code = code.replace(
        "def __init__(self, parent, *args, **kwargs):\n        super().__init__(parent, *args, **kwargs)\n        self.parent = parent",
        "def __init__(self, parent, *args, **kwargs):\n        super().__init__(parent, *args, **kwargs)\n        self.parent = parent\n        self.history_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'daily_work_history.json')"
    )

# 2. Add Auto Calculate button and date bindings
button_code = """        btn_calc = ttk.Button(top_frame, text="자동 집계 및 저장", command=self.auto_calculate_and_save)
        btn_calc.grid(row=0, column=4, padx=5, pady=5)
        
        btn_export = ttk.Button(top_frame, text="엑셀 출력 (일보 생성)", command=self.export_excel)
        btn_export.grid(row=0, column=5, padx=20, pady=5)"""

if "btn_calc" not in code:
    code = code.replace(
        "        btn_export = ttk.Button(top_frame, text=\"엑셀 출력 (일보 생성)\", command=self.export_excel)\n        btn_export.grid(row=0, column=4, padx=20, pady=5)",
        button_code
    )

date_binding_code = """        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.date_entry.bind("<<DateEntrySelected>>", self.on_date_change)
        self.date_entry.bind("<FocusOut>", self.on_date_change)"""

if "<<DateEntrySelected>>" not in code:
    code = code.replace(
        "        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky=\"w\")",
        date_binding_code
    )

# 3. Fix self.qty_entries key to be composite
if "self.qty_entries[f\"{method}_{spec}\"]" not in code:
    code = code.replace(
        "self.qty_entries[spec] = row_dict",
        "self.qty_entries[f\"{method}_{spec}\"] = row_dict"
    )

# Fix export_excel to use composite key
if "for comp_key, entries in self.qty_entries.items():" not in code:
    code = code.replace(
        "        for spec, entries in self.qty_entries.items():\n            data['qty_data'][spec] = {k: v.get() for k, v in entries.items()}",
        "        for comp_key, entries in self.qty_entries.items():\n            data['qty_data'][comp_key] = {k: v.get() for k, v in entries.items()}"
    )

# Also save history in export_excel
if "self.save_current_history()" not in code:
    code = code.replace(
        "        # Save File Dialog",
        "        self.save_current_history()\n        # Save File Dialog"
    )

# 4. Add new methods at the end of the class
methods_code = """
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
            'qty_data': {},
            'equip_data': {},
            'personnel_data': {}
        }
        for comp_key, entries in self.qty_entries.items():
            data['qty_data'][comp_key] = {k: v.get() for k, v in entries.items()}
        for eq, entries in self.equip_entries.items():
            data['equip_data'][eq] = {k: v.get() for k, v in entries.items()}
        for p_key, ent in self.personnel_entries.items():
            data['personnel_data'][p_key] = ent.get()
            
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
            
        # Update Qty 전일누계
        for comp_key, entries in self.qty_entries.items():
            prev_total = prev_data.get('qty_data', {}).get(comp_key, {}).get('총누계', '0')
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, prev_total)
            
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
        for comp_key, val in today_qty.items():
            if '소계' not in comp_key:
                ent = self.qty_entries[comp_key]['금일작업']
                ent.delete(0, tk.END)
                ent.insert(0, f"{val:.1f}" if val % 1 else f"{int(val)}")
                
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
            
            entries['전일누계'].delete(0, tk.END)
            entries['전일누계'].insert(0, f"{prev_total:.1f}" if prev_total % 1 else f"{int(prev_total)}")
            
            total = prev_total + today_val
            entries['총누계'].delete(0, tk.END)
            entries['총누계'].insert(0, f"{total:.1f}" if total % 1 else f"{int(total)}")
            
            expected = float(entries['예상량'].get() or 0)
            if expected > 0:
                progress = (total / expected) * 100
                entries['공정률'].delete(0, tk.END)
                entries['공정률'].insert(0, f"{progress:.1f}")

        # Subtotals for Qty
        paut_total = sum(float(self.qty_entries[f"PAUT_{s}"]['총누계'].get() or 0) for s in ['300A이상', '300A이상-야간', '250A', '200A', '200A-야간'])
        self.qty_entries['PAUT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['PAUT_소계']['총누계'].insert(0, f"{paut_total:.1f}" if paut_total % 1 else f"{int(paut_total)}")
        
        rt_total = sum(float(self.qty_entries[f"RT_{s}"]['총누계'].get() or 0) for s in ['150A~100A', '150A~100A-야간', '80A이하', '80A이하-야간'])
        self.qty_entries['RT_소계']['총누계'].delete(0, tk.END)
        self.qty_entries['RT_소계']['총누계'].insert(0, f"{rt_total:.1f}" if rt_total % 1 else f"{int(rt_total)}")
                
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
        
        # Save to history
        self.save_current_history()
        
        messagebox.showinfo("완료", "자동 집계 및 데이터 저장이 완료되었습니다.")
"""

if "def auto_calculate_and_save" not in code:
    code += methods_code

with open('daily_work_log_tab.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Update completed.")
