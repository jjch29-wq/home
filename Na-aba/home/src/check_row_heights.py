import openpyxl

report_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\RT_Report_20260510_203428.xlsx"

try:
    wb = openpyxl.load_workbook(report_path)
    ws = wb.worksheets[0]
    print(f"Row 43 height: {ws.row_dimensions[43].height}")
    print(f"Row 46 height: {ws.row_dimensions[46].height}")
    print(f"Row 43 hidden: {ws.row_dimensions[43].hidden}")
    print(f"Row 46 hidden: {ws.row_dimensions[46].hidden}")
except Exception as e:
    print(f"ERROR: {e}")
