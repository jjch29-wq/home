from openpyxl import load_workbook
import os

path = os.path.abspath("../resources/Template_DailyWorkReport.xlsx")
wb = load_workbook(path)
sheet = wb.active

for r in range(13, 40):
    val = sheet[f'B{r}'].value
    if val:
        print(f"Row {r}: {val}")

wb.close()
