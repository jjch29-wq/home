import win32com.client

file_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V71.xlsx'
out_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V72.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(file_path)
    
    count = 0
    for s_idx in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(s_idx)
        for shape in ws.Shapes:
            # Method 1: TextFrame2 (newer Excel)
            try:
                if shape.TextFrame2.HasText:
                    text = shape.TextFrame2.TextRange.Text
                    if '동탄지사' in text:
                        shape.TextFrame2.TextRange.Text = text.replace('동탄지사', '중앙지사')
                        count += 1
            except Exception: pass
            
            # Method 2: TextFrame (older Excel)
            try:
                text = shape.TextFrame.Characters().Text
                if '동탄지사' in text:
                    shape.TextFrame.Characters().Text = text.replace('동탄지사', '중앙지사')
                    count += 1
            except Exception: pass
            
            # Method 3: Group Items
            try:
                if shape.Type == 6: # Group
                    for child in shape.GroupItems:
                        try:
                            if child.TextFrame2.HasText:
                                text = child.TextFrame2.TextRange.Text
                                if '동탄지사' in text:
                                    child.TextFrame2.TextRange.Text = text.replace('동탄지사', '중앙지사')
                                    count += 1
                        except Exception: pass
            except Exception: pass
            
    wb.SaveAs(out_path)
    wb.Close(False)
    print(f'Successfully updated {count} shapes in V72.xlsx.')
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
