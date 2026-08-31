import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx', data_only=True)
ws = wb['5.2 용접사 불량률-주배관(완료)']
for r in range(2, 26):
    row_vals = []
    for c in range(1, 10):
        val = str(ws.cell(row=r, column=c).value)
        row_vals.append(val if val != 'None' else "")
    if any(row_vals):
        print(f"Row {r}: " + " | ".join(row_vals))
