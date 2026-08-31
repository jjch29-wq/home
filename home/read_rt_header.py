import openpyxl
wb = openpyxl.load_workbook(r'C:\Users\-\OneDrive\바탕 화면\산출내역서(가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역)_감독경유.xlsx', data_only=True)
ws = wb['6-1. RT']
for r in range(1, 10):
    row_vals = []
    for c in range(1, 15):
        val = str(ws.cell(row=r, column=c).value)
        row_vals.append(val if val != 'None' else "")
    print(f"Row {r}: " + " | ".join(row_vals))
