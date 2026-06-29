with open(r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\rt_data.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

def get_sums(time_label):
    section = False
    total_A = 0
    total_A2 = 0
    for line in lines:
        if '□ RT' in line:
            if time_label in line:
                section = True
            else:
                section = False
            continue
            
        if section and '관리소' in line:
            parts = line.split('\t')
            if '\tA\t' in line:
                try:
                    idx = parts.index('A')
                    val = parts[idx+1].strip()
                    if val and val != '-': total_A += int(val)
                except: pass
            elif '\tA/2\t' in line:
                try:
                    idx = parts.index('A/2')
                    val = parts[idx+1].strip()
                    if val and val != '-': total_A2 += int(val)
                except: pass
    return total_A, total_A2

d_a, d_a2 = get_sums('주간')
n_a, n_a2 = get_sums('야간')
h_a, h_a2 = get_sums('휴일')

print(f'Day: A={d_a}, A/2={d_a2}')
print(f'Night: A={n_a}, A/2={n_a2}')
print(f'Holiday: A={h_a}, A/2={h_a2}')
