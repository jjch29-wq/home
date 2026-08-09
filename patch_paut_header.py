import codecs
import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the PAUT block in do_export (around line 12085 onwards)
# We need to replace from:
#   paut_sheet_name = wb.sheetnames[0]
# to:
#   current_row += 1
# with new header-based logic

start_marker = "                paut_sheet_name = wb.sheetnames[0]"
end_marker = "                current_row += 1"

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if start_marker in line and start_idx == -1:
        start_idx = i
    if start_idx != -1 and end_marker in line:
        end_idx = i
        break

print(f"Found block: lines {start_idx+1} to {end_idx+1}")

if start_idx != -1 and end_idx != -1:
    new_block = '''                # --- PAUT 데이터: 헤더 자동 감지 방식으로 기입 ---
                try:
                    from paut_writer import find_paut_section, write_paut_data
                    target_month_str = f"{year}-{month:02d}"
                    
                    # PAUT 레코드 추출
                    paut_raw = []
                    for date_key, log_data in history.items():
                        if date_key.startswith(target_month_str):
                            for r in log_data.get('ndt_results', []):
                                if str(r.get('검사방법', '')).strip().upper() == 'PAUT':
                                    r['_date'] = date_key
                                    paut_raw.append(r)
                    
                    paut_raw.sort(key=lambda x: x['_date'])
                    
                    groups = {}
                    for r in paut_raw:
                        key = (
                            str(r.get('업체', '')),
                            str(r.get('구간', '')),
                            str(r.get('라인번호', '')),
                            str(r.get('관경', '')),
                            str(r.get('Joint No.', '')),
                            str(r.get('용접사', ''))
                        )
                        paut_val = str(r.get('PAUT', '0')).strip()
                        try: val = float(paut_val)
                        except: val = 0.0
                        if val == 0: continue
                        shift = str(r.get('규격', '주간')).strip()
                        if key not in groups:
                            groups[key] = {'ORI': 0.0, 'RE': 0.0, 'shift': shift}
                        if groups[key]['ORI'] == 0.0:
                            groups[key]['ORI'] += val
                        else:
                            groups[key]['RE'] += val
                    
                    if groups:
                        ws_main = wb.worksheets[0]
                        header_row, data_start, col_map = find_paut_section(ws_main)
                        if header_row:
                            written = write_paut_data(ws_main, groups, header_row, data_start, col_map)
                            log(f"✅ PAUT 데이터 {written}건 기입 완료 (헤더 자동 감지, row {data_start}부터)")
                        else:
                            log("⚠️ PAUT 표 헤더를 찾을 수 없습니다. (1.2.1 위상배열초음파탐상검사 섹션 확인 필요)")
                    else:
                        log(f"⚠️ {target_month_str} 기간의 PAUT 데이터가 없습니다.")
                except Exception as e:
                    log(f"⚠️ PAUT 기입 오류: {e}")
                    import traceback
                    log(traceback.format_exc())
'''
    
    lines[start_idx:end_idx+1] = [new_block]
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print(f"Replaced PAUT block (was lines {start_idx+1}-{end_idx+1})")
else:
    print("Block not found!")
