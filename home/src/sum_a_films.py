import re

file_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\rt_data.txt'
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

day_section = False
a_film_sum = 0
details = []

for line in lines:
    if '□ RT주간' in line:
        day_section = True
    elif '□ RT야간' in line:
        day_section = False
    
    if day_section and '관리소' in line and 'A' in line and 'A/2' not in line:
        parts = line.strip().split('\t')
        # We need to find the "ORIGINAL" column for "가산~가평"
        # The columns typically look like: ... 관리소  137 ... 5  관리소  B  685  ...
        # So the film type is near the end. Let's just print to see.
        details.append(line.strip())

for i, d in enumerate(details):
    print(f'Row {i}: {d}')
