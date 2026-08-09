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

new_block = '''                # --- 4.5 NDT 결과서 섹션 태그 기반 자동 기입 ---
                try:
                    import sys as _sys
                    import os as _os
                    _src = _os.path.dirname(_os.path.abspath(__file__))
                    if _src not in _sys.path:
                        _sys.path.insert(0, _src)
                    
                    from tagged_ndt_writer import write_all_tagged_sections
                    
                    # 저장된 파일을 다시 열어 NDT 기입
                    import openpyxl as _opx
                    wb2 = _opx.load_workbook(save_path)
                    ws2 = wb2.worksheets[0]
                    
                    # 태그 기반 NDT 섹션 기입
                    write_all_tagged_sections(ws2, history, target_month_str, log_func=log)
                    
                    wb2.save(save_path)
                    wb2.close()
                    log("✅ 태그 기반 NDT 결과서 전체 기입 완료")
                    
'''

lines[start:end+1] = [new_block]

with open(r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Updated NDT block with tagged_ndt_writer!")
