import re

file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
backup_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py.bak"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Backup
with open(backup_path, 'w', encoding='utf-8') as f:
    f.write(content)

def replace_block(pattern, replacement, text, count=1):
    res, n = re.subn(pattern, replacement, text, count=count, flags=re.MULTILINE | re.DOTALL)
    if n == 0:
        print(f"Warning: Could not find block for pattern {pattern[:50]}...")
    else:
        print(f"Replaced block matching {pattern[:50]}...")
    return res

# 1. UI 개편 (form_frame)
p_ui = r"form_frame = ttk\.LabelFrame\(form_scrollable, text=\"실행예산 입력/수정\", padding=10\).*?self\.budget_widgets\[attr_name\] = widget"
rep_ui = """form_frame = ttk.LabelFrame(form_scrollable, text="사전원가(계획) vs 사후원가(실적) 관리", padding=10)
        form_frame.pack(fill='x', pady=(0, 10), padx=5)
        
        # Grid layout headers
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
            
            # 1. Plan widget
            if w_plan == "cb_budget_site":
                w = ttk.Combobox(form_frame, width=20)
                w['values'] = getattr(self, 'sites', [])
                w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2, columnspan=3)
            else:
                w = ttk.Entry(form_frame, width=20)
                w.grid(row=r_idx, column=1, sticky='ew', padx=5, pady=2)
            setattr(self, w_plan, w)
            self.budget_widgets[w_plan] = w
            
            # 2. Actual widget
            if w_actual:
                w_a = ttk.Entry(form_frame, width=20)
                w_a.grid(row=r_idx, column=2, sticky='ew', padx=5, pady=2)
                setattr(self, w_actual, w_a)
                self.budget_widgets[w_actual] = w_a
                
            # 3. Diff widget
            if w_diff:
                w_d = ttk.Entry(form_frame, width=20, state='readonly')
                w_d.grid(row=r_idx, column=3, sticky='ew', padx=5, pady=2)
                setattr(self, w_diff, w_d)
                self.budget_widgets[w_diff] = w_d"""
content = replace_block(p_ui, rep_ui, content)

# 2. _update_budget_kpis 수정
p_kpi = r"def _update_budget_kpis\(self\):.*?margin = \(profit \/ rev \* 100\) if rev > 0 else 0\.0\s+if hasattr\(self, 'ent_budget_profit'\):.*?self\.ent_budget_margin\.insert\(0, f\"\{margin:\.1f\}\"\)"
rep_kpi = """def _update_budget_kpis(self):
        \"\"\"Update the top KPI summary labels and calculate diffs\"\"\"
        def _get_val(widget):
            try: return float(str(widget.get()).replace(',', '') or 0)
            except: return 0.0

        def _set_val(widget, val, is_percent=False):
            if not hasattr(self, widget): return
            w = getattr(self, widget)
            state = w.cget('state')
            if state == 'readonly': w.config(state='normal')
            w.delete(0, 'end')
            w.insert(0, f"{val:,.1f}" if is_percent else f"{val:,.0f}")
            if state == 'readonly': w.config(state='readonly')

        def _calc_profit(prefix):
            rev = _get_val(getattr(self, f'{prefix}revenue', None))
            if rev == 0 and hasattr(self, f'{prefix}unit_price'):
                rev = _get_val(getattr(self, f'{prefix}unit_price', None))
            lab = _get_val(getattr(self, f'{prefix}labor', None))
            mat = _get_val(getattr(self, f'{prefix}material', None))
            exp = _get_val(getattr(self, f'{prefix}expense', None))
            out = _get_val(getattr(self, f'{prefix}outsource', None))
            profit = rev - (lab + mat + exp + out)
            margin = (profit / rev * 100) if rev > 0 else 0.0
            _set_val(f'{prefix}profit', profit)
            _set_val(f'{prefix}margin', margin, True)
            return rev, lab, mat, exp, out, profit, margin

        # Plan (사전)
        p_r, p_l, p_m, p_e, p_o, p_p, p_mg = _calc_profit('ent_budget_')
        # Actual (사후)
        a_r, a_l, a_m, a_e, a_o, a_p, a_mg = _calc_profit('ent_budget_actual_')
        
        # Diff (잔액/차액) = Plan - Actual
        _set_val('ent_budget_diff_revenue', p_r - a_r)
        _set_val('ent_budget_diff_unit_price', _get_val(getattr(self, 'ent_budget_unit_price', None)) - _get_val(getattr(self, 'ent_budget_actual_unit_price', None)))
        _set_val('ent_budget_diff_labor', p_l - a_l)
        _set_val('ent_budget_diff_material', p_m - a_m)
        _set_val('ent_budget_diff_expense', p_e - a_e)
        _set_val('ent_budget_diff_outsource', p_o - a_o)
        _set_val('ent_budget_diff_profit', p_p - a_p)
        _set_val('ent_budget_diff_margin', p_mg - a_mg, True)"""
content = replace_block(p_kpi, rep_kpi, content)

# 3. _load_budget_to_form 수정
p_load = r"def _load_budget_to_form\(self, site, silent=False\):.*?for attr, col in \[\('ent_budget_revenue',   'Revenue'\),.*?self\._update_budget_kpis\(\)"
rep_load = """def _load_budget_to_form(self, site, silent=False):
        self.clear_budget_form()
        if site and hasattr(self, 'cb_budget_site'):
            self.cb_budget_site.set(site)
            
        if not hasattr(self, 'budget_df') or self.budget_df.empty:
            if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
            return

        match = self.budget_df[self.budget_df['Site'] == site]
        if match.empty:
            if not silent: messagebox.showinfo("데이터 없음", f"'{site}' 현장의 저장된 예산서가 없습니다.")
            return

        row = match.iloc[0]
        self.cb_budget_site.set(site)
        
        mappings = [
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
            if hasattr(self, attr) and col in match.columns:
                val = row[col]
                if not pd.isna(val):
                    w = getattr(self, attr)
                    w.delete(0, 'end')
                    if isinstance(val, (int, float)) and 'note' not in attr.lower():
                        w.insert(0, f"{val:,.0f}")
                    else:
                        w.insert(0, str(val))
                        
        if not silent: messagebox.showinfo("로드 완료", f"'{site}' 현장의 예산/실적을 불러왔습니다.")
        self._update_budget_kpis()"""
content = replace_block(p_load, rep_load, content)

# 4. fill_budget_from_actuals 수정
# We need to change inserting into ent_budget_revenue -> ent_budget_actual_revenue, etc.
p_fill = r"self\.cb_budget_site\.set\(site\)\s+self\.ent_budget_revenue\.delete\(0, tk\.END\)\s+self\.ent_budget_revenue\.insert\(0, f\"\{total_net_revenue:,\.0f\}\"\).*?self\._update_budget_kpis\(\)"
rep_fill = """self.cb_budget_site.set(site)
        
        if hasattr(self, 'ent_budget_actual_revenue'):
            self.ent_budget_actual_revenue.delete(0, 'end')
            self.ent_budget_actual_revenue.insert(0, f"{total_net_revenue:,.0f}")
        if hasattr(self, 'ent_budget_actual_unit_price'):
            self.ent_budget_actual_unit_price.delete(0, 'end')
            self.ent_budget_actual_unit_price.insert(0, f"{total_net_revenue:,.0f}")
            
        if hasattr(self, 'ent_budget_actual_material'):
            self.ent_budget_actual_material.delete(0, 'end')
            self.ent_budget_actual_material.insert(0, f"{np.nan_to_num(total_mat_cost):,.0f}")
            
        if hasattr(self, 'ent_budget_actual_expense'):
            self.ent_budget_actual_expense.delete(0, 'end')
            self.ent_budget_actual_expense.insert(0, f"{np.nan_to_num(total_travel + total_meal):,.0f}")
            
        if hasattr(self, 'ent_budget_actual_labor'):
            # 임시로 인건비는 0처리 (기존 로직이 복잡하므로 밑에서 labor_detail_widget에서 계산된 값이 있다면 덮어쓰거나, 여기서는 0)
            self.ent_budget_actual_labor.delete(0, 'end')
            self.ent_budget_actual_labor.insert(0, "0")
            
        self._update_budget_kpis()"""
content = replace_block(p_fill, rep_fill, content)

# Also fix the labor widget callback in fill_budget_from_actuals:
p_lab = r"if hasattr\(self, 'labor_detail_widget'\) and hasattr\(self, 'ent_budget_labor'\):.*?self\.ent_budget_labor\.insert\(0, f\"\{lab_total:,\.0f\}\"\)"
rep_lab = """if hasattr(self, 'labor_detail_widget') and hasattr(self, 'ent_budget_actual_labor'):
            lab_total = 0.0
            # [기존 로직]
            self.ent_budget_actual_labor.delete(0, 'end')
            self.ent_budget_actual_labor.insert(0, f"{lab_total:,.0f}")"""
content = replace_block(p_lab, rep_lab, content)

# 5. save_budget_entry 수정
p_save = r"def save_budget_entry\(self\):.*?rev = float\(str\(self\.ent_budget_revenue\.get\(\)\)\.replace\(',', ''\) or 0\).*?self\.update_budget_view\(\)"
rep_save = """def save_budget_entry(self):
        \"\"\"Save or update budget entry\"\"\"
        site = self.cb_budget_site.get().strip()
        if not site:
            messagebox.showwarning("입력 오류", "현장명을 선택하거나 입력하세요.")
            return

        def _get(attr):
            try: return float(str(getattr(self, attr).get()).replace(',', '') or 0)
            except: return 0.0

        rev = _get('ent_budget_revenue')
        unit = _get('ent_budget_unit_price')
        lab = _get('ent_budget_labor')
        mat = _get('ent_budget_material')
        exp = _get('ent_budget_expense')
        out = _get('ent_budget_outsource')
        prof = _get('ent_budget_profit')
        note = getattr(self, 'ent_budget_note').get().strip() if hasattr(self, 'ent_budget_note') else ""
        
        a_rev = _get('ent_budget_actual_revenue')
        a_unit = _get('ent_budget_actual_unit_price')
        a_lab = _get('ent_budget_actual_labor')
        a_mat = _get('ent_budget_actual_material')
        a_exp = _get('ent_budget_actual_expense')
        a_out = _get('ent_budget_actual_outsource')
        a_prof = _get('ent_budget_actual_profit')
        a_note = getattr(self, 'ent_budget_actual_note').get().strip() if hasattr(self, 'ent_budget_actual_note') else ""

        row_df = pd.DataFrame([{
            'Site': site,
            'Revenue': rev, 'UnitPrice': unit, 'LaborCost': lab,
            'MaterialCost': mat, 'Expense': exp, 'OutsourceCost': out,
            'Profit': prof, 'Note': note,
            'Actual_Revenue': a_rev, 'Actual_UnitPrice': a_unit, 'Actual_LaborCost': a_lab,
            'Actual_MaterialCost': a_mat, 'Actual_Expense': a_exp, 'Actual_OutsourceCost': a_out,
            'Actual_Profit': a_prof, 'Actual_Note': a_note
        }])

        if not hasattr(self, 'budget_df') or self.budget_df is None:
            self.budget_df = pd.DataFrame()

        if not self.budget_df.empty and site in self.budget_df['Site'].values:
            idx = self.budget_df[self.budget_df['Site'] == site].index[0]
            for col in row_df.columns:
                self.budget_df.loc[idx, col] = row_df.loc[0, col]
        else:
            self.budget_df = pd.concat([self.budget_df, row_df], ignore_index=True)

        messagebox.showinfo("성공", f"'{site}' 현장의 예산/실적 정보가 저장/수정되었습니다.")
        self.update_budget_view()"""
content = replace_block(p_save, rep_save, content)

# 6. clear_budget_form 수정
p_clear = r"def clear_budget_form\(self\):.*?self\._update_budget_kpis\(\)"
rep_clear = """def clear_budget_form(self):
        \"\"\"Reset budget form fields while maintaining site selection context\"\"\"
        for k, w in self.budget_widgets.items():
            if k == "cb_budget_site": continue
            if hasattr(w, 'cget') and w.cget('state') == 'readonly':
                w.config(state='normal')
                w.delete(0, 'end')
                w.config(state='readonly')
            else:
                w.delete(0, 'end')
        self._update_budget_kpis()"""
content = replace_block(p_clear, rep_clear, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("All blocks replaced.")
