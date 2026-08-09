with open(r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

start = -1
end = -1
for i, line in enumerate(lines):
    if '# --- 4.5 NDT 결과서 섹션 헤더 기반 자동 기입' in line and start == -1:
        start = i
    if start != -1 and 'wb2.close()' in line:
        end = i
        break

print(f"Found block at lines {start+1} to {end+1}")

new_block = '''                # --- 4.5 NDT 결과서 섹션 헤더 기반 자동 기입 ---
                try:
                    import sys as _sys
                    import os as _os
                    _src = _os.path.dirname(_os.path.abspath(__file__))
                    if _src not in _sys.path:
                        _sys.path.insert(0, _src)
                    
                    from ndt_section_writer import write_all_ndt_sections, extract_records_by_method
                    from paut_writer import find_paut_section, write_paut_data
                    
                    # 저장된 파일을 다시 열어 NDT 기입
                    import openpyxl as _opx
                    wb2 = _opx.load_workbook(save_path)
                    ws2 = wb2.worksheets[0]
                    
                    # PAUT (1.2.1 섹션) - paut_writer의 헤더 자동감지 사용
                    paut_recs = extract_records_by_method(history, target_month_str, 'PAUT')
                    if paut_recs:
                        paut_groups = {}
                        for _r in paut_recs:
                            _k = (_r['업체'], _r['구간'], _r['라인번호'], _r['관경'], _r['Joint No.'], _r['용접사'])
                            paut_groups[_k] = {'ORI': _r['ORI'], 'RE': _r['RE'], 'shift': _r['규격']}
                        _hr, _ds, _cm = find_paut_section(ws2)
                        if _hr:
                            _w = write_paut_data(ws2, paut_groups, _hr, _ds, _cm)
                            log(f"✅ 1.2.1 PAUT {_w}건 기입 완료 (row {_ds}~)")
                    
                    # 나머지 모든 NDT 섹션 (1.2.2 MT, 1.2.3 RT, 1.2.4 PT, 2.x 등) 자동 기입
                    write_all_ndt_sections(ws2, history, target_month_str, log_func=log)
                    
                    wb2.save(save_path)
                    wb2.close()
                    log("✅ NDT 결과서 전체 기입 완료")
                    
'''

lines[start:end+1] = [new_block]

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f"Updated NDT block!")
