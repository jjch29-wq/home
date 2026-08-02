import json
config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'
with open(config_path, 'r', encoding='utf-8') as f:
    c = json.load(f)

# UT Updates (All assigned to 열배관 as per excel)
c['CONTRACT_QTY']['열배관']['일반']['UT'] = 237.66
c['CONTRACT_QTY']['열배관']['야간']['UT'] = 54.24
c['CONTRACT_QTY']['열배관']['휴일']['UT'] = 27.12

c['CONTRACT_QTY']['플랜트(관리소)']['일반']['UT'] = 0.0
c['CONTRACT_QTY']['플랜트(관리소)']['야간']['UT'] = 0.0
c['CONTRACT_QTY']['플랜트(관리소)']['휴일']['UT'] = 0.0

# PT Updates (Split between 열배관 and 관리소)
c['CONTRACT_QTY']['열배관']['일반']['PT'] = 285.19
c['CONTRACT_QTY']['열배관']['야간']['PT'] = 65.08
c['CONTRACT_QTY']['열배관']['휴일']['PT'] = 32.54

c['CONTRACT_QTY']['플랜트(관리소)']['일반']['PT'] = 21.56
c['CONTRACT_QTY']['플랜트(관리소)']['야간']['PT'] = 3.25
c['CONTRACT_QTY']['플랜트(관리소)']['휴일']['PT'] = 2.96

with open(config_path, 'w', encoding='utf-8') as f:
    json.dump(c, f, indent=4, ensure_ascii=False)
print('Config UT/PT updated!')
