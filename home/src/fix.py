import re

with open(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Fix 1: apply_custom_dimensions separator
old1 = '''                    if '-' in part or '~' in part:
                        sep = '-' if '-' in part else '~'
                        start, end = map(int, part.split(sep))'''
new1 = '''                    if '-' in part or '~' in part or ':' in part:
                        if '-' in part: sep = '-'
                        elif '~' in part: sep = '~'
                        else: sep = ':'
                        start, end = map(int, part.split(sep))'''
text = text.replace(old1, new1)

# Fix 2 & 3: print orientation and print area
old2 = '''            ws.page_setup.paperSize = 9
            # PAUT 갑지와 을지는 모두 A4 세로 방향으로 출력한다.
            if mode in ("PMI", "PAUT"):
                ws.page_setup.orientation = 'portrait'
            elif context == "DATA" or mode == "RT":
                ws.page_setup.orientation = 'landscape'
            else:
                ws.page_setup.orientation = 'portrait'
            
            # [FIX] 상황별 인쇄영역 동적 지정 (RT_COVER 외 COVER 시트들)
            if mode == "RT":
                if context == "COVER":
                    ws.print_area = 'A1:V35'
                else:
                    ws.print_area = 'A1:V37'
            else:
                if context == "COVER":
                    # [NEW] Highly Dynamic print area for Gapji
                    key = f"{mode}_GAPJI_PRINT_END_ROW" if mode != "PMI" else "GAPJI_PRINT_END_ROW"
                    end_r = int(self.config.get(key, 51))
                    if end_r > 0:
                        ws.print_area = f'A1:T{end_r}'
                else:
                    # [NEW] Highly Dynamic print area for Eulji
                    key = f"{mode}_PRINT_END_ROW" if mode != "PMI" else "PRINT_END_ROW"
                    end_r = int(self.config.get(key, 47))
                    if end_r > 0:
                        if mode == "PAUT":
                            ws.print_area = f'A1:AH{end_r}'
                        else:
                            ws.print_area = f'A1:M{end_r}\''''

new2 = '''            ws.page_setup.paperSize = 9
            # PAUT 갑지와 을지는 모두 A4 세로 방향으로 출력한다.
            if mode in ("PMI", "PAUT", "PT"):
                ws.page_setup.orientation = 'portrait'
            elif context == "DATA" or mode == "RT":
                ws.page_setup.orientation = 'landscape'
            else:
                ws.page_setup.orientation = 'portrait'
            
            # [FIX] 상황별 인쇄영역 동적 지정 (RT_COVER 외 COVER 시트들)
            if mode == "PT" and context == "COVER":
                ws.print_area = 'A1:S47'
            elif mode == "RT":
                if context == "COVER":
                    ws.print_area = 'A1:V35'
                else:
                    ws.print_area = 'A1:V37'
            else:
                if context == "COVER":
                    # [NEW] Highly Dynamic print area for Gapji
                    key = f"{mode}_GAPJI_PRINT_END_ROW" if mode != "PMI" else "GAPJI_PRINT_END_ROW"
                    end_r = int(self.config.get(key, 51))
                    if end_r > 0:
                        ws.print_area = f'A1:T{end_r}'
                else:
                    # [NEW] Highly Dynamic print area for Eulji
                    key = f"{mode}_PRINT_END_ROW" if mode != "PMI" else "PRINT_END_ROW"
                    end_r = int(self.config.get(key, 47))
                    if end_r > 0:
                        if mode == "PAUT":
                            ws.print_area = f'A1:AH{end_r}'
                        elif mode == "PT":
                            ws.print_area = f'A1:S{end_r}'
                        else:
                            ws.print_area = f'A1:M{end_r}\''''

text = text.replace(old2, new2)

# Fix 4: PT data loop rewrite
pattern = r'                if is_new_pt_template:.*?for c in range\(1, 12\):'
new4 = '''                # PT 고정 열 하드코딩 (H:Q 병합 주의)
                self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=1).coordinate, item.get('Dwg', ''))
                self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=4).coordinate, item.get('Joint', ''))
                
                res_val = str(item.get('Result', 'Acc')).strip().upper()
                is_acc = (res_val in ['ACC', 'ACCEPT', 'PASS', '합격', 'O', 'OK', 'V', ''])
                
                if is_acc:
                    self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=5).coordinate, "V")
                    self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=8).coordinate, "NO RECORDABLE INDICATION")
                else:
                    self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=6).coordinate, "V")
                    self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=8).coordinate, res_val)
                    
                self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=18).coordinate, item.get('Welder', ''))
                self.safe_set_value(current_ws, current_ws.cell(row=current_row, column=19).coordinate, item.get('NPS', ''))
                
                # 스타일링
                for c in range(1, 20):'''
text = re.sub(pattern, new4, text, flags=re.DOTALL)

with open(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write(text)
print('Re-applied all fixes successfully.')
