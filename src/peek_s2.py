import openpyxl
import os

template_file = os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "Project PROVIDENCE 작업일보(Template).xlsx")
# Search for the template file in parent directories if not found
if not os.path.exists(template_file):
    for root, dirs, files in os.walk(os.path.join(os.getcwd(), "..")):
        if "Project PROVIDENCE 작업일보(Template).xlsx" in files:
            template_file = os.path.join(root, "Project PROVIDENCE 작업일보(Template).xlsx")
            break

print(f"Loading {template_file}")
wb = openpyxl.load_workbook(template_file, data_only=True)
ws = wb.active
for r in range(11, 16):
    print(f'Row {r}:', [c.value for c in ws[r]])
