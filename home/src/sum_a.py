with open(r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\rt_data.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

day_section = False
total = 0
for line in lines:
    if '□ RT주간' in line:
        day_section = True
    elif '□ RT야간' in line:
        day_section = False
        
    if day_section and '관리소' in line and '\tA\t' in line:
        parts = line.split('\t')
        try:
            film_idx = parts.index('A')
            # Look at values right after 'A'
            val = parts[film_idx + 1].strip()
            if val and val != '-':
                print(f'Found: {val} in line: {line.strip()}')
                total += int(val)
        except:
            pass
            
print(f'Total A film for Day Station: {total}')
