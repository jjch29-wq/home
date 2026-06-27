import openpyxl

try:
    file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\resources\Template_DailyWorkReport.xlsx'
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    
    print("MERGED CELLS:")
    for merged_range in sheet.merged_cells.ranges:
        print(merged_range)
except Exception as e:
    print(f"Error: {e}")
