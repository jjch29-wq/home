import sys, json, openpyxl
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
from paut_writer import find_paut_section, write_paut_data

history_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json'
with open(history_path, 'r', encoding='utf-8') as f:
    history = json.load(f)

target_month = '2026-08'

# Extract PAUT records
paut_raw = []
for date_key, log_data in history.items():
    if date_key.startswith(target_month):
        for r in log_data.get('ndt_results', []):
            if str(r.get('검사방법', '')).strip().upper() == 'PAUT':
                r['_date'] = date_key
                paut_raw.append(r)

paut_raw.sort(key=lambda x: x['_date'])

groups = {}
for r in paut_raw:
    key = (str(r.get('업체','')), str(r.get('구간','')), str(r.get('라인번호','')),
           str(r.get('관경','')), str(r.get('Joint No.','')), str(r.get('용접사','')))
    val = float(str(r.get('PAUT','0') or '0').strip() or 0)
    if val == 0: continue
    if key not in groups:
        groups[key] = {'ORI': 0.0, 'RE': 0.0, 'shift': str(r.get('규격','주간'))}
    if groups[key]['ORI'] == 0.0:
        groups[key]['ORI'] += val
    else:
        groups[key]['RE'] += val

print(f"PAUT groups: {len(groups)}")

# Write to template
wb = openpyxl.load_workbook(r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V74.xlsx')
ws = wb.worksheets[0]

# Verify headers are intact
print("\n=== Checking headers ===")
for row in [403, 404, 448, 449]:
    cells = {}
    for col in range(1, 25):
        c = ws.cell(row=row, column=col)
        if c.value:
            cells[col] = str(c.value).strip()
    print(f"  Row {row}: {cells}")

print("\n=== Writing PAUT to 1.2.1 section (header=403, data starts=405) ===")
header_row, data_start, col_map = find_paut_section(ws)
print(f"  Auto-detected: header={header_row}, data_start={data_start}")
print(f"  col_map: {col_map}")

written = write_paut_data(ws, groups, header_row, data_start, col_map)
print(f"  Written: {written} rows")

wb.save(r'C:\Users\jjch2\Desktop\Test_Final.xlsx')
print("\nSaved to C:\\Users\\jjch2\\Desktop\\Test_Final.xlsx")
print("Please check rows 405+ for PAUT data")
