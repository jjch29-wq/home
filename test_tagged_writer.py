import sys, json, openpyxl
sys.path.insert(0, r'c:\Users\jjch2\Desktop\PMI\home\src')
from tagged_ndt_writer import write_all_tagged_sections

history_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json'
with open(history_path, 'r', encoding='utf-8') as f:
    history = json.load(f)

template_path = r'C:/Users/jjch2/Desktop/템플릿_최종완성본_V74.xlsx'
wb = openpyxl.load_workbook(template_path)
ws = wb.worksheets[0]

target_month = '2026-08'
print("=== 태그 기반 NDT 기입 테스트 ===")
write_all_tagged_sections(ws, history, target_month)

wb.save(r'C:\Users\jjch2\Desktop\Test_Tagged.xlsx')
print("Done! Saved to C:\\Users\\jjch2\\Desktop\\Test_Tagged.xlsx")
