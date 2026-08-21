import sys
file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def print_context(func_name):
    print(f"--- {func_name} ---")
    start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith(f"def {func_name}"):
            start = i
            break
    if start != -1:
        for i in range(start, min(start+40, len(lines))):
            print(f"Line {i+1}: {lines[i].rstrip()}")
        print("...")

print_context("_update_budget_kpis")
print_context("_load_budget_to_form")
print_context("fill_budget_from_actuals")
print_context("save_budget_entry")
print_context("clear_budget_form")
