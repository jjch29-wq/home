import win32com.client

file_path = r'C:\Users\jjch2\Desktop\템플릿_최종완성본_V71.xlsx'

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(file_path)
    
    for s_idx in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(s_idx)
        print(f'--- Sheet: {ws.Name} ---')
        
        # Check shapes
        for shape in ws.Shapes:
            try:
                # Type 17 is msoTextBox, 1 is msoAutoShape, 6 is msoGroup
                print(f'Shape Name: {shape.Name}, Type: {shape.Type}')
                
                if shape.Type == 6: # Group
                    for child in shape.GroupItems:
                        print(f'  Child Name: {child.Name}, Type: {child.Type}')
                        if child.HasTextFrame:
                            if child.TextFrame.HasText:
                                print(f'  Child Text: {repr(child.TextFrame.Characters().Text)}')
                
                if shape.HasTextFrame:
                    if shape.TextFrame.HasText:
                        print(f'Shape Text: {repr(shape.TextFrame.Characters().Text)}')
            except Exception as e:
                print(f'Error reading shape {shape.Name}: {e}')
                
    wb.Close(False)
except Exception as e:
    print('Error:', e)
finally:
    excel.Quit()
