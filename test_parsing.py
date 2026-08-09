import json

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json', 'r', encoding='utf-8') as f:
    history = json.load(f)

paut_records = []
target_month_str = "2026-08"

for date_key, log_data in history.items():
    if date_key.startswith(target_month_str):
        ndt_results = log_data.get('ndt_results', [])
        for r in ndt_results:
            if str(r.get('검사방법', '')).strip().upper() == 'PAUT':
                r['_date'] = date_key
                paut_records.append(r)

print(f"Total PAUT records found: {len(paut_records)}")

if paut_records:
    paut_records.sort(key=lambda x: x['_date'])
    
    groups = {}
    for r in paut_records:
        key = (str(r.get('업체', '')), str(r.get('구간', '')), str(r.get('라인번호', '')), str(r.get('관경', '')), str(r.get('Joint No.', '')))
        
        paut_val = str(r.get('PAUT', '0')).strip()
        try:
            val = float(paut_val)
        except:
            val = 0.0
            
        if val == 0: continue
            
        shift = str(r.get('규격', '주간')).strip()
        
        if key not in groups:
            groups[key] = {'ORI': 0.0, 'RE': 0.0, 'shift': shift}
            
        if groups[key]['ORI'] == 0.0:
            groups[key]['ORI'] += val
        else:
            groups[key]['RE'] += val
            
    print(f"Total grouped records: {len(groups)}")
    for i, (key, data) in enumerate(groups.items()):
        if i < 5:
            print(f"Row {i}: Key={key}, Data={data}")
else:
    print("No records found!")
