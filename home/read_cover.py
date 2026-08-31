import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx', data_only=True)
ws = wb['표지']
for r in range(1, 15):
    for c in range(1, 10):
        val = str(ws.cell(row=r, column=c).value)
        if val != 'None':
            print(f"[표지] R{r}C{c}: {val}")
