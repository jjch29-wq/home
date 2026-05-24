import openpyxl

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\가스공사 의뢰서.xlsx"
try:
    wb = openpyxl.load_workbook(file_path)
    print(f"Sheet count: {len(wb.worksheets)}")
    for i, ws in enumerate(wb.worksheets):
        print(f"  sheets[{i}]: '{ws.title}'")
except Exception as e:
    print(f"Error: {e}")

# Also check if RT template file exists
rt_template = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT KS양식.xlsx"
try:
    wb2 = openpyxl.load_workbook(rt_template)
    print(f"\nRT Template Sheet count: {len(wb2.worksheets)}")
    for i, ws in enumerate(wb2.worksheets):
        print(f"  sheets[{i}]: '{ws.title}'")
except Exception as e:
    print(f"RT Template Error: {e}")
