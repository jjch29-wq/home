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

# Fix 2: print area
old2 = '''            # [FIX] 상황별 인쇄영역 동적 지정 (RT_COVER 외 COVER 시트들)
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
                            ws.print_area = f'A1:M{end_r}''''

new2 = '''            # [FIX] 상황별 인쇄영역 동적 지정 (RT_COVER 외 COVER 시트들)
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

with open(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py', 'w', encoding='utf-8') as f:
    f.write(text)
print('Applied missing fixes!')
