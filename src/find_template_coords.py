import openpyxl
import os

def find_text_in_excel(file_path, search_text):
    if not os.path.exists(file_path):
        return f"File not found: {file_path}"
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        results = []
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and search_text in str(cell.value):
                    results.append(f"Found '{search_text}' at {cell.coordinate} (Row {cell.row}, Col {cell.column_letter})")
        return results if results else f"Text '{search_text}' not found."
    except Exception as e:
        return f"Error: {e}"

# Standard paths
current_dir = os.path.dirname(os.path.abspath(__file__)) # src
home_dir = os.path.dirname(current_dir)
template_path = os.path.join(home_dir, 'resources', 'Template_DailyWorkReport.xlsx')

print(f"Searching in {template_path}...")
print(find_text_in_excel(template_path, "시건장치"))
print(find_text_in_excel(template_path, "차량번호"))
print(find_text_in_excel(template_path, "안함"))
