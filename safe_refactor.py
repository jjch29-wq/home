import os

file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def replace_between(start_str, end_str, new_content):
    start_idx = -1
    end_idx = -1
    for i, line in enumerate(lines):
        if start_str in line and start_idx == -1:
            start_idx = i
        if end_str in line and start_idx != -1 and i > start_idx and end_idx == -1:
            end_idx = i
    if start_idx != -1 and end_idx != -1:
        new_lines = new_content.split('\n')
        new_lines = [l + '\n' for l in new_lines]
        del lines[start_idx:end_idx+1]
        for l in reversed(new_lines):
            lines.insert(start_idx, l)
        print(f"Replaced {start_idx} to {end_idx}")
    else:
        print(f"Failed to find {start_str} or {end_str}")

# 1. form_frame UI
replace_between(
    'rows = [',
    'self.budget_widgets[attr_name] = widget',
    """        # Grid layout headers
        headers = ["항목", "사전예산(계획)", "사후원가(실적)", "잔여예산(차액)"]
        for col, h in enumerate(headers):
            ttk.Label(form_frame, text=h, font=('Malgun Gothic', 9, 'bold')).grid(row=0, column=col, padx=10, pady=5)
            
        rows = [
            ("현장명", "cb_budget_site", None, None),
            ("계약금액(Revenue)", "ent_budget_revenue", "ent_budget_actual_revenue", "ent_budget_diff_revenue"),
            ("매출금액(UnitPrice)", "ent_budget_unit_price", "ent_budget_actual_unit_price", "ent_budget_diff_unit_price"),
            ("실행 노무비(Labor)", "ent_budget_labor", "ent_budget_actual_labor", "ent_budget_diff_labor"),
            ("실행 재료비(Material)", "ent_budget_material", "ent_budget_actual_material", "ent_budget_diff_material"),
            ("실행 경비(Expense)", "ent_budget_expense", "ent_budget_actual_expense", "ent_budget_diff_expense"),
            ("실행 외주비(Outsource)", "ent_budget_outsource", "ent_budget_actual_outsource", "ent_budget_diff_outsource"),
            ("영업이익(Profit)", "ent_budget_profit", "ent_budget_actual_profit", "ent_budget_diff_profit"),
            ("이익률(%)", "ent_budget_margin", "ent_budget_actual_margin", "ent_budget_diff_margin"),
            ("비고", "ent_budget_note", "ent_budget_actual_note", None)
        ]

        self.budget_widgets = {}
        for r_idx, (label, w_plan, w_actual, w_diff) in enumerate(rows, start=1):
            ttk.Label(form_frame, text=label).grid(row=r_idx, column=0, sticky='e', padx=10, pady=2)
            if w_plan == "cb_budget_site":
                w = ttk.Combobox(form_frame, width=20)
                w['values'] = getattr(self, 'sites', [])
                w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2, columnspan=3)
            else:
                w = ttk.Entry(form_frame, width=20)
                w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2)
            setattr(self, w_plan, w)
            self.budget_widgets[w_plan] = w
            
            if w_actual:
                w_a = ttk.Entry(form_frame, width=20)
                w_a.grid(row=r_idx, column=2, sticky='ew', padx=5, pady=2)
                setattr(self, w_actual, w_a)
                self.budget_widgets[w_actual] = w_a
                
            if w_diff:
                w_d = ttk.Entry(form_frame, width=20, state='readonly')
                w_d.grid(row=r_idx, column=3, sticky='ew', padx=5, pady=2)
                setattr(self, w_diff, w_d)
                self.budget_widgets[w_diff] = w_d"""
)

# 2. _update_budget_kpis
replace_between(
    'def _update_budget_kpis(self):',
    'self.root.update_idletasks()',
    """    def _update_budget_kpis(self):
        \"\"\"Update the top KPI summary labels based on current budget form values\"\"\"
        try:
            def _get_val(widget_name):
                try: 
                    w = getattr(self, widget_name, None)
                    if w:
                        v = str(w.get()).replace(',', '').split(' ')[0]
                        return float(v or 0)
                    return 0.0
                except: return 0.0
                
            def _set_val(widget_name, val, is_percent=False):
                w = getattr(self, widget_name, None)
                if w:
                    st = w.cget('state')
                    if st == 'readonly': w.config(state='normal')
                    w.delete(0, 'end')
                    w.insert(0, f"{val:,.1f}" if is_percent else f"{val:,.0f}")
                    if st == 'readonly': w.config(state='readonly')

            def _calc_profit(prefix):
                rev = _get_val(f'{prefix}revenue')
                if rev == 0: rev = _get_val(f'{prefix}unit_price')
                lab = _get_val(f'{prefix}labor')
                mat = _get_val(f'{prefix}material')
                exp = _get_val(f'{prefix}expense')
                out = _get_val(f'{prefix}outsource')
                
                # Check for detail widget grand total only for plan
                if prefix == 'ent_budget_' and hasattr(self, 'expense_detail_widget') and hasattr(self.expense_detail_widget, 'lbl_grand_total_cost'):
                    try:
                        raw_t = self.expense_detail_widget.lbl_grand_total_cost.cget('text')
                        tc = float("".join(c for c in raw_t if c.isdigit() or c == '.') or 0)
                    except: tc = lab + mat + exp + out
                else:
                    tc = lab + mat + exp + out
                    
                prof = rev - tc
                mar = (prof / rev * 100) if rev > 0 else 0.0
                _set_val(f'{prefix}profit', prof)
                _set_val(f'{prefix}margin', mar, True)
                return rev, lab, mat, exp, out, prof, mar
                
            p_r, p_l, p_m, p_e, p_o, p_p, p_mg = _calc_profit('ent_budget_')
            a_r, a_l, a_m, a_e, a_o, a_p, a_mg = _calc_profit('ent_budget_actual_')
            
            _set_val('ent_budget_diff_revenue', p_r - a_r)
            _set_val('ent_budget_diff_unit_price', _get_val('ent_budget_unit_price') - _get_val('ent_budget_actual_unit_price'))
            _set_val('ent_budget_diff_labor', p_l - a_l)
            _set_val('ent_budget_diff_material', p_m - a_m)
            _set_val('ent_budget_diff_expense', p_e - a_e)
            _set_val('ent_budget_diff_outsource', p_o - a_o)
            _set_val('ent_budget_diff_profit', p_p - a_p)
            _set_val('ent_budget_diff_margin', p_mg - a_mg, True)
            
            # update top KPI labels
            if hasattr(self, 'lbl_kpi_rev'):
                self.lbl_kpi_rev.config(text=f"{p_r:,.0f}원")
                self.lbl_kpi_cost.config(text=f"{p_l+p_m+p_e+p_o:,.0f}원")
                self.lbl_kpi_profit.config(text=f"{p_p:,.0f}원")
                self.lbl_kpi_margin.config(text=f"{p_mg:.1f}%")
            self.root.update_idletasks()"""
)

# 3. clear_budget_form
replace_between(
    '# self.cb_budget_site.set(',
    'self.ent_budget_note.delete(0, tk.END)',
    """        for k, w in self.budget_widgets.items():
            if k == 'cb_budget_site': continue
            st = w.cget('state')
            if st == 'readonly': w.config(state='normal')
            w.delete(0, 'end')
            if st == 'readonly': w.config(state='readonly')"""
)

# 4. _load_budget_to_form
replace_between(
    "        for attr, col in [('ent_budget_revenue',   'Revenue')",
    "                    w.insert(0, str(val))",
    """        mappings = [
            ('ent_budget_revenue', 'Revenue'), ('ent_budget_unit_price','UnitPrice'),
            ('ent_budget_labor', 'LaborCost'), ('ent_budget_material', 'MaterialCost'),
            ('ent_budget_expense', 'Expense'), ('ent_budget_outsource', 'OutsourceCost'),
            ('ent_budget_profit', 'Profit'), ('ent_budget_note', 'Note'),
            ('ent_budget_actual_revenue', 'Actual_Revenue'), ('ent_budget_actual_unit_price','Actual_UnitPrice'),
            ('ent_budget_actual_labor', 'Actual_LaborCost'), ('ent_budget_actual_material', 'Actual_MaterialCost'),
            ('ent_budget_actual_expense', 'Actual_Expense'), ('ent_budget_actual_outsource', 'Actual_OutsourceCost'),
            ('ent_budget_actual_profit', 'Actual_Profit'), ('ent_budget_actual_note', 'Actual_Note')
        ]
        
        for attr, col in mappings:
            w = getattr(self, attr, None)
            if w and col in match.columns:
                st = w.cget('state')
                if st == 'readonly': w.config(state='normal')
                w.delete(0, 'end')
                val = row[col]
                if not pd.isna(val) and str(val).lower() != 'nan':
                    if isinstance(val, (int, float)) and 'note' not in attr.lower():
                        w.insert(0, f"{val:,.0f}")
                    else:
                        w.insert(0, str(val))
                if st == 'readonly': w.config(state='readonly')"""
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(lines)
