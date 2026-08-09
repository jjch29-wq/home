import re

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the "MonthlyReportExporter" call block and REPLACE it with full NDT writing
# The block starts around: "# --- 4.5 태그 변환 (MonthlyReportExporter) ---"
# and ends after: "exporter.generate(save_path)"

start_marker = "                # --- 4.5"
end_marker = "                    exporter.generate(save_path)"

start_idx = -1
end_idx = -1
for i, line in enumerate(lines):
    if start_marker in line and start_idx == -1:
        start_idx = i
    if start_idx != -1 and end_marker in line:
        end_idx = i
        break

print(f"Found block at lines {start_idx+1} - {end_idx+1}")

if start_idx != -1 and end_idx != -1:
    new_block = '''                # --- 4.5 NDT 결과서 섹션 헤더 기반 자동 기입 ---
                try:
                    import sys as _sys
                    import os as _os
                    _src = _os.path.dirname(_os.path.abspath(__file__))
                    if _src not in _sys.path:
                        _sys.path.insert(0, _src)
                    
                    from ndt_section_writer import find_ndt_section, write_ndt_section_data, extract_ndt_records
                    from paut_writer import find_paut_section, write_paut_data
                    
                    # 이미 저장된 파일을 다시 열어 NDT 기입
                    import openpyxl as _opx
                    wb2 = _opx.load_workbook(save_path)
                    ws2 = wb2.worksheets[0]
                    
                    # PAUT (1.2.1 섹션)
                    paut_list = extract_ndt_records(history, target_month_str, 'PAUT')
                    if paut_list:
                        paut_groups = {}
                        for _r in paut_list:
                            _k = (_r['업체'], _r['구간'], _r['라인번호'], _r['관경'], _r['Joint No.'], _r['용접사'])
                            paut_groups[_k] = {'ORI': _r['ORI'], 'RE': _r['RE'], 'shift': _r['규격']}
                        _hr, _ds, _cm = find_paut_section(ws2)
                        if _hr:
                            _w = write_paut_data(ws2, paut_groups, _hr, _ds, _cm)
                            log(f"✅ 1.2.1 PAUT {_w}건 기입 완료 (row {_ds}~)")
                    
                    # 2.1 비파괴(PAUT) 섹션
                    if paut_list:
                        _hr, _ds, _cm = find_ndt_section(ws2, ['2.1', 'PAUT'])
                        if _hr:
                            _w = write_ndt_section_data(ws2, paut_list, _hr, _ds, _cm, 'PAUT')
                            log(f"✅ 2.1 PAUT {_w}건 기입 완료 (row {_ds}~)")
                    
                    # 2.3 RT 섹션
                    rt_list = extract_ndt_records(history, target_month_str, 'RT')
                    if rt_list:
                        _hr, _ds, _cm = find_ndt_section(ws2, ['2.3', 'RT'])
                        if _hr:
                            _w = write_ndt_section_data(ws2, rt_list, _hr, _ds, _cm, 'RT')
                            log(f"✅ 2.3 RT {_w}건 기입 완료 (row {_ds}~)")
                    
                    # 2.1 MT 섹션
                    mt_list = extract_ndt_records(history, target_month_str, 'MT')
                    if mt_list:
                        _hr, _ds, _cm = find_ndt_section(ws2, ['MT'])
                        if _hr:
                            _w = write_ndt_section_data(ws2, mt_list, _hr, _ds, _cm, 'MT')
                            log(f"✅ MT {_w}건 기입 완료 (row {_ds}~)")
                    
                    # PT 섹션
                    pt_list = extract_ndt_records(history, target_month_str, 'PT')
                    if pt_list:
                        _hr, _ds, _cm = find_ndt_section(ws2, ['PT'])
                        if _hr:
                            _w = write_ndt_section_data(ws2, pt_list, _hr, _ds, _cm, 'PT')
                            log(f"✅ PT {_w}건 기입 완료 (row {_ds}~)")
                    
                    wb2.save(save_path)
                    wb2.close()
                    log("✅ NDT 결과서 전체 기입 완료")
                    
                except Exception as ex:
                    log(f"⚠️ NDT 결과서 기입 오류 (무시됨): {ex}")
                    import traceback
                    log(traceback.format_exc())
'''
    lines[start_idx:end_idx+1] = [new_block]
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print(f"Replaced MonthlyReportExporter block with full NDT section writer!")
else:
    print("Block not found!")
