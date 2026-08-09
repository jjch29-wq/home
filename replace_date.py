import win32com.client

file_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V68.xlsx'
out_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V69.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Sheets(1)
    
    count = 0
    for r in range(1, 400):
        for c in range(1, 20):
            val = str(ws.Cells(r, c).Value or '')
            if '2026년 07월 31일' in val:
                new_val = val.replace('2026년 07월 31일 ~ 2026년 08월 05일', '2026년 08월 05일 ~ 2027년 08월 05일')
                ws.Cells(r, c).Value = new_val
                count += 1
                
    wb.SaveAs(out_path)
    wb.Close(False)
    print(f'Successfully updated {count} cells in V69.xlsx.')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
