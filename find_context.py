import sys
file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def print_around(match_str, n=5):
    for i, line in enumerate(lines):
        if match_str in line:
            print(f"--- MATCH at {i} ---")
            for j in range(max(0, i-n), min(len(lines), i+n+1)):
                print(f"{j}: {lines[j].rstrip()}")
            print("")

print_around('form_frame = ttk.LabelFrame(form_scrollable, text="실행예산 입력/수정"', 2)
print_around('self.budget_widgets[attr_name] = widget', 2)

print_around('def _update_budget_kpis(self):', 2)
print_around('self.root.update_idletasks()', 2)

print_around("for attr, col in [('ent_budget_revenue',   'Revenue'),", 2)
print_around('if not silent: messagebox.showinfo("로드 완료", f"\'{site}\' 현장의 예산서를 불러왔습니다.")', 2)

print_around('self.cb_budget_site.set(site)', 2)
print_around('def save_budget_entry(self):', 2)
