import re

file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

def replace_regex(pattern, repl, text):
    res, count = re.subn(pattern, repl, text, count=1, flags=re.MULTILINE | re.DOTALL)
    if count == 0:
        print(f"Failed pattern: {pattern[:50]}...")
    else:
        print(f"Success for pattern: {pattern[:50]}...")
    return res

# 1. UI 개편
p1 = r"(\s*)rows = \[\s*\(\"현장명:\", \"cb_budget_site\"\),.*?self\.budget_widgets\[attr_name\] = widget"
r1 = r"""\1# Grid layout headers
\1headers = ["항목", "사전예산(계획)", "사후원가(실적)", "잔여예산(차액)"]
\1for col, h in enumerate(headers):
\1    ttk.Label(form_frame, text=h, font=('Malgun Gothic', 9, 'bold')).grid(row=0, column=col, padx=10, pady=5)
\1    
\1rows = [
\1    ("현장명", "cb_budget_site", None, None),
\1    ("계약금액(Revenue)", "ent_budget_revenue", "ent_budget_actual_revenue", "ent_budget_diff_revenue"),
\1    ("매출금액(UnitPrice)", "ent_budget_unit_price", "ent_budget_actual_unit_price", "ent_budget_diff_unit_price"),
\1    ("실행 노무비(Labor)", "ent_budget_labor", "ent_budget_actual_labor", "ent_budget_diff_labor"),
\1    ("실행 재료비(Material)", "ent_budget_material", "ent_budget_actual_material", "ent_budget_diff_material"),
\1    ("실행 경비(Expense)", "ent_budget_expense", "ent_budget_actual_expense", "ent_budget_diff_expense"),
\1    ("실행 외주비(Outsource)", "ent_budget_outsource", "ent_budget_actual_outsource", "ent_budget_diff_outsource"),
\1    ("영업이익(Profit)", "ent_budget_profit", "ent_budget_actual_profit", "ent_budget_diff_profit"),
\1    ("이익률(%)", "ent_budget_margin", "ent_budget_actual_margin", "ent_budget_diff_margin"),
\1    ("비고", "ent_budget_note", "ent_budget_actual_note", None)
\1]
\1
\1self.budget_widgets = {}
\1for r_idx, (label, w_plan, w_actual, w_diff) in enumerate(rows, start=1):
\1    ttk.Label(form_frame, text=label).grid(row=r_idx, column=0, sticky='e', padx=10, pady=2)
\1    if w_plan == "cb_budget_site":
\1        w = ttk.Combobox(form_frame, width=20)
\1        w['values'] = getattr(self, 'sites', [])
\1        w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2, columnspan=3)
\1    else:
\1        w = ttk.Entry(form_frame, width=20)
\1        w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2)
\1    setattr(self, w_plan, w)
\1    self.budget_widgets[w_plan] = w
\1    
\1    if w_actual:
\1        w_a = ttk.Entry(form_frame, width=20)
\1        w_a.grid(row=r_idx, column=2, sticky='ew', padx=5, pady=2)
\1        setattr(self, w_actual, w_a)
\1        self.budget_widgets[w_actual] = w_a
\1        
\1    if w_diff:
\1        w_d = ttk.Entry(form_frame, width=20, state='readonly')
\1        w_d.grid(row=r_idx, column=3, sticky='ew', padx=5, pady=2)
\1        setattr(self, w_diff, w_d)
\1        self.budget_widgets[w_diff] = w_d"""
content = replace_regex(p1, r1, content)

# 2. _update_budget_kpis
p2 = r"(\s*)def _update_budget_kpis\(self\):.*?print\(f\"DEBUG: Budget KPI sync error: \{e\}\"\)"
r2 = r"""\1def _update_budget_kpis(self):
\1    \"\"\"Update the top KPI summary labels based on current budget form values\"\"\"
\1    try:
\1        def _get_val(widget_name):
\1            try: 
\1                w = getattr(self, widget_name, None)
\1                if w:
\1                    v = str(w.get()).replace(',', '').split(' ')[0]
\1                    return float(v or 0)
\1                return 0.0
\1            except: return 0.0
\1            
\1        def _set_val(widget_name, val, is_percent=False):
\1            w = getattr(self, widget_name, None)
\1            if w:
\1                st = w.cget('state')
\1                if st == 'readonly': w.config(state='normal')
\1                w.delete(0, 'end')
\1                w.insert(0, f"{val:,.1f}" if is_percent else f"{val:,.0f}")
\1                if st == 'readonly': w.config(state='readonly')
\1
\1        def _calc_profit(prefix):
\1            rev = _get_val(f'{prefix}revenue')
\1            if rev == 0: rev = _get_val(f'{prefix}unit_price')
\1            lab = _get_val(f'{prefix}labor')
\1            mat = _get_val(f'{prefix}material')
\1            exp = _get_val(f'{prefix}expense')
\1            out = _get_val(f'{prefix}outsource')
\1            
\1            if prefix == 'ent_budget_' and hasattr(self, 'expense_detail_widget') and hasattr(self.expense_detail_widget, 'lbl_grand_total_cost'):
\1                try:
\1                    raw_t = self.expense_detail_widget.lbl_grand_total_cost.cget('text')
\1                    tc = float("".join(c for c in raw_t if c.isdigit() or c == '.') or 0)
\1                except: tc = lab + mat + exp + out
\1            else:
\1                tc = lab + mat + exp + out
\1                
\1            prof = rev - tc
\1            mar = (prof / rev * 100) if rev > 0 else 0.0
\1            _set_val(f'{prefix}profit', prof)
\1            _set_val(f'{prefix}margin', mar, True)
\1            return rev, lab, mat, exp, out, prof, mar
\1            
\1        p_r, p_l, p_m, p_e, p_o, p_p, p_mg = _calc_profit('ent_budget_')
\1        a_r, a_l, a_m, a_e, a_o, a_p, a_mg = _calc_profit('ent_budget_actual_')
\1        
\1        _set_val('ent_budget_diff_revenue', p_r - a_r)
\1        _set_val('ent_budget_diff_unit_price', _get_val('ent_budget_unit_price') - _get_val('ent_budget_actual_unit_price'))
\1        _set_val('ent_budget_diff_labor', p_l - a_l)
\1        _set_val('ent_budget_diff_material', p_m - a_m)
\1        _set_val('ent_budget_diff_expense', p_e - a_e)
\1        _set_val('ent_budget_diff_outsource', p_o - a_o)
\1        _set_val('ent_budget_diff_profit', p_p - a_p)
\1        _set_val('ent_budget_diff_margin', p_mg - a_mg, True)
\1        
\1        if hasattr(self, 'lbl_kpi_rev'):
\1            self.lbl_kpi_rev.config(text=f"{p_r:,.0f}원")
\1            self.lbl_kpi_cost.config(text=f"{p_l+p_m+p_e+p_o:,.0f}원")
\1            self.lbl_kpi_profit.config(text=f"{p_p:,.0f}원", foreground="#ef4444" if p_p < 0 else "#10b981")
\1            self.lbl_kpi_margin.config(text=f"{p_mg:.1f}%", foreground="#ef4444" if p_mg < 0 else "#10b981")
\1        self.root.update_idletasks()
\1    except Exception as e:
\1        print(f"DEBUG: Budget KPI sync error: {e}")"""
content = replace_regex(p2, r2, content)

# 3. _load_budget_to_form
p3 = r"(\s*)def _load_budget_to_form\(self, site, silent=False\):.*?self\._update_budget_kpis\(\)"
r3 = r"""\1def _load_budget_to_form(self, site, silent=False):
\1    self.clear_budget_form()
\1    if site and hasattr(self, 'cb_budget_site'):
\1        self.cb_budget_site.set(site)
\1        
\1    if not hasattr(self, 'budget_df') or self.budget_df.empty:
\1        if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
\1        return
\1
\1    match = self.budget_df[self.budget_df['Site'] == site]
\1    if match.empty:
\1        if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
\1        return
\1
\1    row = match.iloc[0]
\1    self.cb_budget_site.set(site)
\1    
\1    mappings = [
\1        ('ent_budget_revenue', 'Revenue'), ('ent_budget_unit_price','UnitPrice'),
\1        ('ent_budget_labor', 'LaborCost'), ('ent_budget_material', 'MaterialCost'),
\1        ('ent_budget_expense', 'Expense'), ('ent_budget_outsource', 'OutsourceCost'),
\1        ('ent_budget_profit', 'Profit'), ('ent_budget_note', 'Note'),
\1        ('ent_budget_actual_revenue', 'Actual_Revenue'), ('ent_budget_actual_unit_price','Actual_UnitPrice'),
\1        ('ent_budget_actual_labor', 'Actual_LaborCost'), ('ent_budget_actual_material', 'Actual_MaterialCost'),
\1        ('ent_budget_actual_expense', 'Actual_Expense'), ('ent_budget_actual_outsource', 'Actual_OutsourceCost'),
\1        ('ent_budget_actual_profit', 'Actual_Profit'), ('ent_budget_actual_note', 'Actual_Note')
\1    ]
\1    
\1    for attr, col in mappings:
\1        w = getattr(self, attr, None)
\1        if w and col in match.columns:
\1            st = w.cget('state')
\1            if st == 'readonly': w.config(state='normal')
\1            w.delete(0, 'end')
\1            val = row[col]
\1            if not pd.isna(val) and str(val).lower() != 'nan':
\1                if isinstance(val, (int, float)) and 'note' not in attr.lower():
\1                    w.insert(0, f"{val:,.0f}")
\1                else:
\1                    w.insert(0, str(val))
\1            if st == 'readonly': w.config(state='readonly')
\1            
\1    if not silent: messagebox.showinfo("로드 완료", f"'{site}' 현장의 예산/실적을 불러왔습니다.")
\1    self._update_budget_kpis()"""
content = replace_regex(p3, r3, content)

# 4. fill_budget_from_actuals (just the end part)
p4 = r"(\s*)self\.cb_budget_site\.set\(site\)\s+# \[NEW\].*?self\._update_budget_kpis\(\)"
r4 = r"""\1self.cb_budget_site.set(site)
\1
\1if hasattr(self, 'ent_budget_actual_revenue'):
\1    self.ent_budget_actual_revenue.delete(0, 'end')
\1    self.ent_budget_actual_revenue.insert(0, f"{total_net_revenue:,.0f}")
\1if hasattr(self, 'ent_budget_actual_unit_price'):
\1    self.ent_budget_actual_unit_price.delete(0, 'end')
\1    self.ent_budget_actual_unit_price.insert(0, f"{total_net_revenue:,.0f}")
\1    
\1if hasattr(self, 'ent_budget_actual_material'):
\1    self.ent_budget_actual_material.delete(0, 'end')
\1    self.ent_budget_actual_material.insert(0, f"{total_mat_cost:,.0f}")
\1    
\1if hasattr(self, 'ent_budget_actual_expense'):
\1    self.ent_budget_actual_expense.delete(0, 'end')
\1    self.ent_budget_actual_expense.insert(0, f"{total_travel + total_meal:,.0f}")
\1    
\1if hasattr(self, 'ent_budget_actual_labor'):
\1    # Actual labor could be mapped here from lab_total if lab_total is computed in this scope
\1    # Let's see if lab_total is available. Yes, it was available.
\1    try:
\1        self.ent_budget_actual_labor.delete(0, 'end')
\1        self.ent_budget_actual_labor.insert(0, f"{lab_total:,.0f}")
\1    except: pass
\1    
\1self._update_budget_kpis()"""
content = replace_regex(p4, r4, content)

# 5. save_budget_entry
p5 = r"(\s*)def save_budget_entry\(self\):.*?self\.update_budget_view\(\)"
r5 = r"""\1def save_budget_entry(self):
\1    \"\"\"Save or update budget entry\"\"\"
\1    site = self.cb_budget_site.get().strip()
\1    if not site:
\1        messagebox.showwarning("입력 오류", "현장명을 선택하거나 입력하세요.")
\1        return
\1
\1    def _get(attr):
\1        try: 
\1            w = getattr(self, attr, None)
\1            if w: return float(str(w.get()).replace(',', '') or 0)
\1            return 0.0
\1        except: return 0.0
\1
\1    rev = _get('ent_budget_revenue')
\1    unit = _get('ent_budget_unit_price')
\1    lab = _get('ent_budget_labor')
\1    mat = _get('ent_budget_material')
\1    exp = _get('ent_budget_expense')
\1    out = _get('ent_budget_outsource')
\1    prof = _get('ent_budget_profit')
\1    note = getattr(self, 'ent_budget_note').get().strip() if hasattr(self, 'ent_budget_note') else ""
\1    
\1    a_rev = _get('ent_budget_actual_revenue')
\1    a_unit = _get('ent_budget_actual_unit_price')
\1    a_lab = _get('ent_budget_actual_labor')
\1    a_mat = _get('ent_budget_actual_material')
\1    a_exp = _get('ent_budget_actual_expense')
\1    a_out = _get('ent_budget_actual_outsource')
\1    a_prof = _get('ent_budget_actual_profit')
\1    a_note = getattr(self, 'ent_budget_actual_note').get().strip() if hasattr(self, 'ent_budget_actual_note') else ""
\1
\1    import pandas as pd
\1    row_df = pd.DataFrame([{
\1        'Site': site,
\1        'Revenue': rev, 'UnitPrice': unit, 'LaborCost': lab,
\1        'MaterialCost': mat, 'Expense': exp, 'OutsourceCost': out,
\1        'Profit': prof, 'Note': note,
\1        'Actual_Revenue': a_rev, 'Actual_UnitPrice': a_unit, 'Actual_LaborCost': a_lab,
\1        'Actual_MaterialCost': a_mat, 'Actual_Expense': a_exp, 'Actual_OutsourceCost': a_out,
\1        'Actual_Profit': a_prof, 'Actual_Note': a_note
\1    }])
\1
\1    if not hasattr(self, 'budget_df') or self.budget_df is None:
\1        self.budget_df = pd.DataFrame()
\1
\1    if not self.budget_df.empty and site in self.budget_df['Site'].values:
\1        idx = self.budget_df[self.budget_df['Site'] == site].index[0]
\1        for col in row_df.columns:
\1            self.budget_df.loc[idx, col] = row_df.loc[0, col]
\1    else:
\1        self.budget_df = pd.concat([self.budget_df, row_df], ignore_index=True)
\1
\1    messagebox.showinfo("성공", f"'{site}' 현장의 예산/실적 정보가 저장/수정되었습니다.")
\1    self.update_budget_view()"""
content = replace_regex(p5, r5, content)

# 6. clear_budget_form
p6 = r"(\s*)def clear_budget_form\(self\):.*?self\._update_budget_kpis\(\)"
r6 = r"""\1def clear_budget_form(self):
\1    \"\"\"Reset budget form fields while maintaining site selection context\"\"\"
\1    if hasattr(self, 'budget_widgets'):
\1        for k, w in self.budget_widgets.items():
\1            if k == "cb_budget_site": continue
\1            if hasattr(w, 'cget'):
\1                st = w.cget('state')
\1                if st == 'readonly': w.config(state='normal')
\1                w.delete(0, 'end')
\1                if st == 'readonly': w.config(state='readonly')
\1            else:
\1                try: w.delete(0, 'end')
\1                except: pass
\1    if hasattr(self, 'labor_detail_widget'): self.labor_detail_widget.reset()
\1    if hasattr(self, 'material_detail_widget'): self.material_detail_widget.reset()
\1    if hasattr(self, 'expense_detail_widget'): self.expense_detail_widget.reset()
\1    self._update_budget_kpis()"""
content = replace_regex(p6, r6, content)


with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
