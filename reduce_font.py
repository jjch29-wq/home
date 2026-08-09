import win32com.client

file_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V69.xlsx'
out_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V71.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Sheets(1)
    
    count = 0
    for r in range(1, 600):
        for c in range(1, 20):
            val = str(ws.Cells(r, c).Value or '')
            if '단가계약' in val and '\n' in val:
                # Turn WrapText back on just in case
                ws.Cells(r, c).WrapText = True
                
                # Reduce font size of only the first line to 9.5
                idx = val.find('\n')
                if idx != -1:
                    ws.Cells(r, c).Characters(1, idx).Font.Size = 9.5
                count += 1
                
    wb.SaveAs(out_path)
    wb.Close(False)
    print(f'Successfully reduced first-line font size to 9.5 for {count} cells in V71.xlsx.')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
