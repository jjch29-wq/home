import sys
file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def replace_lines(start_idx, end_idx, new_content):
    new_lines = new_content.split('\n')
    new_lines = [l + '\n' for l in new_lines]
    del lines[start_idx:end_idx+1]
    for l in reversed(new_lines):
        lines.insert(start_idx, l)

# Define exact line numbers based on our previous discovery
# setup_budget_tab form:
# I will find the line number of "form_frame = ttk.LabelFrame(form_scrollable, text=\"실행예산 입력/수정\", padding=10)"
idx_form_start = next(i for i, l in enumerate(lines) if 'form_frame = ttk.LabelFrame(form_scrollable, text="실행예산 입력/수정"' in l)
idx_form_end = next(i for i, l in enumerate(lines) if 'self.budget_widgets[attr_name] = widget' in l and i > idx_form_start)

# _update_budget_kpis:
idx_kpi_start = next(i for i, l in enumerate(lines) if 'def _update_budget_kpis(self):' in l)
idx_kpi_end = next(i for i, l in enumerate(lines) if 'self.root.update_idletasks()' in l and i > idx_kpi_start)

# _load_budget_to_form:
idx_load_start = next(i for i, l in enumerate(lines) if 'for attr, col in [(\'ent_budget_revenue\',   \'Revenue\'),' in l)
idx_load_end = next(i for i, l in enumerate(lines) if 'w.insert(0, str(val))' in l and i > idx_load_start and i < idx_load_start + 40)

# fill_budget_from_actuals:
idx_fill_start = next(i for i, l in enumerate(lines) if 'self.cb_budget_site.set(site)' in l and i > 7300)
# wait, there's multiple cb_budget_site.set(site). Let's use the one around line 7750
idx_fill_start = next(i for i, l in enumerate(lines) if 'self.cb_budget_site.set(site)' in l and i > 7700)
idx_fill_end = next(i for i, l in enumerate(lines) if 'self._update_budget_kpis()' in l and i > idx_fill_start)

# save_budget_entry:
idx_save_start = next(i for i, l in enumerate(lines) if 'try:' in l and i > 7930) # line 7931
idx_save_end = next(i for i, l in enumerate(lines) if 'self.update_budget_view()' in l and i > idx_save_start) # line 7998


print(f"Form: {idx_form_start} to {idx_form_end}")
print(f"KPI: {idx_kpi_start} to {idx_kpi_end}")
print(f"Load: {idx_load_start} to {idx_load_end}")
print(f"Fill: {idx_fill_start} to {idx_fill_end}")
print(f"Save: {idx_save_start} to {idx_save_end}")
