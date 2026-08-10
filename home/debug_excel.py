import openpyxl
import os
import glob

files = glob.glob(r'C:\Users\-\PMI\home\assets\*.xlsx')
template_path = files[0]
print("Using:", template_path)
wb = openpyxl.load_workbook(template_path, data_only=True)
with open('debug_out.txt', 'w', encoding='utf-8') as f:
    for ws in wb.worksheets:
        for r in range(1, 15):
            for c in range(1, 15):
                val = ws.cell(row=r, column=c).value
                if val:
                    f.write(f'[{ws.title}] R{r}C{c}: {repr(val)}\n')
print("Done")
