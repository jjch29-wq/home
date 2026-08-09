import json
with open(r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json', 'r', encoding='utf-8') as f:
    history = json.load(f)

for date, data in history.items():
    if 'ndt_results' in data:
        for r in data['ndt_results']:
            if r.get('검사', '') == 'PAUT' or r.get('검사방법', '') == 'PAUT' or r.get('검사방법', '') == 'PAUT':
                keys = list(r.keys())
                for k in keys:
                    if '번호' in k:
                        print(f"Key containing '번호': '{k}' (hex: {[hex(ord(c)) for c in k]})")
                break
        break
