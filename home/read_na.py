import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx', data_only=True)
ws = wb['나. 검사물량 세부내역 (주배관)']
for r in range(1, 20):
    row_vals = []
    for c in range(1, 9):
        val = ws.cell(row=r, column=c).value
        row_vals.append(str(val) if val is not None else "")
    print(f"Row {r}: " + " | ".join(row_vals))
