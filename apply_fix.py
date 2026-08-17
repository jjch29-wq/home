import sys

file_path = r'C:\Users\jjch2\Desktop\PMI\home\src\services\monthly_report_manager.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

target1 = '''            else:
                data_items = [
                    (k, {'count': v['count'], 'qty': v['qty'], 'ori': v['qty'], 're': 0.0})
                    for k, v in ndt_groups[method].items()
                ]'''
                
replace1 = '''            else:
                data_items = []
                current_company = None
                sub_count = 0
                sub_qty = 0.0
                
                # 업체명 기준으로 1차 정렬
                sorted_groups = sorted(ndt_groups[method].items(), key=lambda x: str(x[0][0]))
                
                for k, v in sorted_groups:
                    comp = k[0]
                    if current_company is not None and comp != current_company:
                        data_items.append(
                            ((f"[{current_company} 소계]", "", "", "", ""), {'count': sub_count, 'qty': sub_qty, 'ori': sub_qty, 're': 0.0, 'is_subtotal': True})
                        )
                        sub_count = 0
                        sub_qty = 0.0
                        
                    current_company = comp
                    data_items.append(
                        (k, {'count': v['count'], 'qty': v['qty'], 'ori': v['qty'], 're': 0.0, 'is_subtotal': False})
                    )
                    sub_count += v['count']
                    sub_qty += v['qty']
                    
                if current_company is not None:
                    data_items.append(
                        ((f"[{current_company} 소계]", "", "", "", ""), {'count': sub_count, 'qty': sub_qty, 'ori': sub_qty, 're': 0.0, 'is_subtotal': True})
                    )'''

target2 = '''            current_row = start_row
            for idx, ((comp, sec, line, size, spec), vals) in enumerate(data_items):
                count = vals['count']
                qty = vals['qty']
                unit = "매" if method == "RT" else "m"'''
                
replace2 = '''            current_row = start_row
            seq_num = 1
            for idx, ((comp, sec, line, size, spec), vals) in enumerate(data_items):
                count = vals['count']
                qty = vals['qty']
                unit = "매" if method == "RT" else "m"
                is_subtotal = vals.get('is_subtotal', False)
                if is_subtotal:
                    unit = ""'''

target3 = '''                # 셀 위치: 2:업체, 3:순번, 4:Section, 6:Line No., 10:관경, 12:용접개소, 14:규격, 16:단위, 17:길이
                ws.cell(row=current_row, column=2).value = comp
                ws.cell(row=current_row, column=3).value = str(idx+1)'''
                
replace3 = '''                # 셀 위치: 2:업체, 3:순번, 4:Section, 6:Line No., 10:관경, 12:용접개소, 14:규격, 16:단위, 17:길이
                ws.cell(row=current_row, column=2).value = comp
                if is_subtotal:
                    ws.cell(row=current_row, column=3).value = ""
                else:
                    ws.cell(row=current_row, column=3).value = str(seq_num)
                    seq_num += 1'''

target4 = '''                # 폰트, 정렬 적용 (병합된 칸 전체에 정렬 속성을 먹여야 엑셀이 줄바꿈을 정상 인식함)
                for c_idx in range(2, 24):
                    c_cell = ws.cell(row=current_row, column=c_idx)
                    c_cell.font = self.font_normal'''

replace4 = '''                # 폰트, 정렬 적용 (병합된 칸 전체에 정렬 속성을 먹여야 엑셀이 줄바꿈을 정상 인식함)
                for c_idx in range(2, 24):
                    c_cell = ws.cell(row=current_row, column=c_idx)
                    if is_subtotal:
                        import copy
                        bold_font = copy.copy(self.font_normal)
                        bold_font.bold = True
                        c_cell.font = bold_font
                    else:
                        c_cell.font = self.font_normal'''

if target1 not in content: print('target1 not found')
if target2 not in content: print('target2 not found')
if target3 not in content: print('target3 not found')
if target4 not in content: print('target4 not found')

content = content.replace(target1, replace1)
content = content.replace(target2, replace2)
content = content.replace(target3, replace3)
content = content.replace(target4, replace4)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('Patched successfully!')
