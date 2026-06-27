from openpyxl import load_workbook
import os

path = os.path.abspath("../resources/Template_DailyWorkReport.xlsx")
wb = load_workbook(path)
sheet = wb.active

for r in range(25, 30):
    b = sheet[f'B{r}'].value
    c = sheet[f'C{r}'].value
    d = sheet[f'D{r}'].value
    print(f"Row {r}: B={b}, C={c}, D={d}")

wb.close()
