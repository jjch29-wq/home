import re

file_path = r"C:\Users\-\PMI\home\src\한국지역난방 중앙지사.py"
out_path = r"C:\Users\-\PMI\search_out.txt"

keywords = ['사전원가', '단가', '투입', '예산', '중앙지사', '한국지역난방', 'budget', '공사실행예산서']

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

with open(out_path, 'w', encoding='utf-8') as out_f:
    for i, line in enumerate(lines):
        if any(k in line for k in keywords):
            out_f.write(f"Line {i+1}: {line.strip()}\n")
