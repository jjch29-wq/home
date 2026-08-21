import re

file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "_update_budget_kpis" in line:
        print(f"Line {i+1}: {line.strip()}")
