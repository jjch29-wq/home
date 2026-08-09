import codecs

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\한국지역난방 중앙지사.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    lines = f.readlines()

injection_code = """
                # 1.1 비파괴검사 물량표 작성 (PAUT)
                paut_sheet_name = '1.1.2.1 위상배열초음파탐상검사'
                if paut_sheet_name in sheet_names:
                    try:
                        paut_records = []
                        target_month_str = f"{year}-{month:02d}"
                        
                        for date_key, log_data in history.items():
                            if date_key.startswith(target_month_str):
                                raw_data = log_data.get('raw_data', [])
                                for r in raw_data:
                                    if str(r.get('검사', '')).strip().upper() == 'PAUT':
                                        r['_date'] = date_key
                                        paut_records.append(r)
                                        
                        if paut_records:
                            # 날짜순 정렬
                            paut_records.sort(key=lambda x: x['_date'])
                            
                            groups = {}
                            for r in paut_records:
                                key = (str(r.get('업체', '')), str(r.get('구간', '')), str(r.get('도면번호', '')), str(r.get('관경', '')), str(r.get('Joint No.', '')))
                                
                                paut_val = str(r.get('PAUT', '0')).strip()
                                try:
                                    val = float(paut_val)
                                except:
                                    val = 0.0
                                    
                                if val == 0: continue
                                    
                                shift = str(r.get('주야간', '주간')).strip()
                                
                                if key not in groups:
                                    groups[key] = {'ORI': 0.0, 'RE': 0.0, 'shift': shift}
                                    
                                if groups[key]['ORI'] == 0.0:
                                    groups[key]['ORI'] += val
                                else:
                                    groups[key]['RE'] += val
                                    
                            ws_paut = wb.Sheets(paut_sheet_name)
                            # 기존 내용 지우기 (B405:M1000)
                            ws_paut.Range("B405:M1000").ClearContents()
                            
                            current_row = 405
                            for i, (key, data) in enumerate(groups.items()):
                                업체, 구간, 도면번호, 관경, joint = key
                                ws_paut.Cells(current_row, 2).Value = 업체
                                ws_paut.Cells(current_row, 3).Value = i + 1
                                ws_paut.Cells(current_row, 4).Value = 구간
                                ws_paut.Cells(current_row, 5).Value = 도면번호
                                ws_paut.Cells(current_row, 6).Value = 관경
                                ws_paut.Cells(current_row, 7).Value = joint
                                ws_paut.Cells(current_row, 8).Value = data['shift']
                                ws_paut.Cells(current_row, 9).Value = 'M'
                                
                                ori = data['ORI']
                                re_val = data['RE']
                                tot = ori + re_val
                                
                                if ori > 0: ws_paut.Cells(current_row, 10).Value = round(ori, 4)
                                if re_val > 0: ws_paut.Cells(current_row, 11).Value = round(re_val, 4)
                                if tot > 0: ws_paut.Cells(current_row, 12).Value = round(tot, 4)
                                
                                current_row += 1
                                
                            log(f"▶ '{paut_sheet_name}' 시트 기입 완료 (총 {len(groups)}건)")
                    except Exception as e:
                        log(f"⚠️ '{paut_sheet_name}' 작성 중 오류: {e}")
"""

# Find the injection point right before wb.Save()
for i in range(12110, 12150):
    if 'wb.Save()' in lines[i]:
        lines.insert(i, injection_code)
        break

# Also, fix the sheet_names bug that was missed for mgmt_agg!
for i in range(12110, 12150):
    if 'if sheet_name in wb.sheetnames:' in lines[i]:
        lines[i] = lines[i].replace('wb.sheetnames', 'sheet_names')
    if 'write_ndt_sheet(wb[sheet_name]' in lines[i]:
        lines[i] = lines[i].replace('wb[sheet_name]', 'wb.Sheets(sheet_name)')

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.writelines(lines)
