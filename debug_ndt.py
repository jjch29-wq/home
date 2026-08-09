import sys, json, openpyxl, traceback
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')

history_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json'
template_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V74.xlsx'
output_path = r'C:\Users\jjch2\Desktop\Debug_NDT.xlsx'
target_month = '2026-08'

print("=== STEP 1: history 로드 ===")
with open(history_path, 'r', encoding='utf-8') as f:
    history = json.load(f)
print(f"  날짜 키: {list(history.keys())}")

print("\n=== STEP 2: ndt_results 원본 확인 ===")
for date_key, log_data in history.items():
    if date_key.startswith(target_month):
        ndt = log_data.get('ndt_results', [])
        print(f"  {date_key}: {len(ndt)}건")
        for r in ndt[:3]:
            keys = list(r.keys())
            print(f"    키 목록: {keys}")
            for k in ['검사방법', '업체', '구간', '라인번호', 'Joint No.', 'PAUT', 'RT', 'RT_OR']:
                print(f"    {k}: {repr(r.get(k, '(없음)'))}")
            print()

print("\n=== STEP 3: extract_ndt_records 결과 ===")
from ndt_section_writer import extract_ndt_records
paut_list = extract_ndt_records(history, target_month, 'PAUT')
rt_list = extract_ndt_records(history, target_month, 'RT')
mt_list = extract_ndt_records(history, target_month, 'MT')
pt_list = extract_ndt_records(history, target_month, 'PT')
print(f"  PAUT: {len(paut_list)}건")
print(f"  RT: {len(rt_list)}건")
print(f"  MT: {len(mt_list)}건")
print(f"  PT: {len(pt_list)}건")
if paut_list:
    print(f"  PAUT 첫 번째: {paut_list[0]}")

print("\n=== STEP 4: 엑셀 섹션 탐지 ===")
wb = openpyxl.load_workbook(template_path)
ws = wb.worksheets[0]

from paut_writer import find_paut_section
hr, ds, cm = find_paut_section(ws)
print(f"  1.2.1 PAUT 섹션: header_row={hr}, data_start={ds}")
print(f"  col_map: {cm}")

from ndt_section_writer import find_ndt_section
for kws in [['2.1', 'PAUT'], ['2.3', 'RT'], ['MT'], ['PT']]:
    hr2, ds2, cm2 = find_ndt_section(ws, kws)
    print(f"  {kws}: header_row={hr2}, data_start={ds2}")

print("\n=== STEP 5: PAUT 데이터 기입 시도 ===")
try:
    from paut_writer import write_paut_data
    if paut_list:
        paut_groups = {}
        for r in paut_list:
            key = (r['업체'], r['구간'], r['라인번호'], r['관경'], r['Joint No.'], r['용접사'])
            paut_groups[key] = {'ORI': r['ORI'], 'RE': r['RE'], 'shift': r['규격']}
        
        hr, ds, cm = find_paut_section(ws)
        if hr:
            written = write_paut_data(ws, paut_groups, hr, ds, cm)
            print(f"  기입 완료: {written}건")
        else:
            print("  섹션을 찾지 못함!")
    wb.save(output_path)
    print(f"  저장: {output_path}")
except Exception as e:
    print(f"  오류: {e}")
    traceback.print_exc()
