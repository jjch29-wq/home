import json
config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'
with open(config_path, 'r', encoding='utf-8') as f:
    c = json.load(f)
c['CONTRACT_QTY']['플랜트(관리소)']['일반']['RT_A'] = 1969
c['CONTRACT_QTY']['플랜트(관리소)']['일반']['RT_A2'] = 1359
c['CONTRACT_QTY']['플랜트(관리소)']['야간']['RT_A'] = 256
c['CONTRACT_QTY']['플랜트(관리소)']['야간']['RT_A2'] = 168
c['CONTRACT_QTY']['플랜트(관리소)']['휴일']['RT_A'] = 239
c['CONTRACT_QTY']['플랜트(관리소)']['휴일']['RT_A2'] = 177
with open(config_path, 'w', encoding='utf-8') as f:
    json.dump(c, f, indent=4, ensure_ascii=False)
print('Config forcibly updated!')
