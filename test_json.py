import json

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json', 'r', encoding='utf-8') as f:
    history = json.load(f)

with open('out_json.txt', 'w', encoding='utf-8') as out:
    for date, data in history.items():
        if date.startswith('2026-08'):
            raw = data.get('ndt_results', [])
            for r in raw:
                out.write(json.dumps(r, ensure_ascii=False) + '\n')
