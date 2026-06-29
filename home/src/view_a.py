with open(r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\rt_data.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()
for i in range(130, 260):
    if 'A' in lines[i] and '관리소' in lines[i]:
        print(f'{i}: {lines[i].strip()}')
