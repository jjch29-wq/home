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
                # Keep it multiline
                ws.Cells(r, c).WrapText = True
                ws.Cells(r, c).ShrinkToFit = False
                
                # Reduce font size of the whole cell to 10.0
                ws.Cells(r, c).Font.Size = 10.0
                count += 1
                
    wb.SaveAs(out_path)
    wb.Close(False)
    print(f'Successfully reduced font size to 10.0 for {count} cells in V71.xlsx.')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
