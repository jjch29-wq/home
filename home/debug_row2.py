import openpyxl
import glob

# Try to find the exact template file
files = glob.glob('C:/Users/**/템플릿*.xlsx', recursive=True)
if not files:
    print("Template not found!")
else:
    desktop_template = [f for f in files if 'V70' in f and ('바탕' in f or 'Desktop' in f or '' in f)][0]
    print(f"Loading: {desktop_template}")
    wb = openpyxl.load_workbook(desktop_template, data_only=True)
    
    with open('row2_debug.txt', 'w', encoding='utf-8') as f:
        for ws in wb.worksheets:
            f.write(f"=== Sheet: {ws.title} ===\n")
            for r in range(1, 4):
                for c in range(1, 20):
                    val = ws.cell(row=r, column=c).value
                    if val:
                        f.write(f"  R{r}C{c}: {repr(val)}\n")
    print("Done writing row2_debug.txt")
