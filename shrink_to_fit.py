import win32com.client

file_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V69.xlsx'
out_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V70.xlsx'

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
            # Looking for common header phrases
            if '2026년' in val and '단가계약' in val:
                # Enable Shrink to Fit
                ws.Cells(r, c).ShrinkToFit = True
                
                # Check if it has a manual newline character (\n). If so, replace with space.
                # Sometimes people press Alt+Enter causing it to break manually.
                # Actually, if wrap text is on, it might break. We should turn off wrap text.
                ws.Cells(r, c).WrapText = False
                
                count += 1
                
    wb.SaveAs(out_path)
    wb.Close(False)
    print(f'Successfully applied ShrinkToFit and disabled WrapText to {count} cells in V70.xlsx.')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
