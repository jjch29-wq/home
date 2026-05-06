from openpyxl import load_workbook
import os

path = os.path.abspath("../resources/Template_DailyWorkReport.xlsx")
wb = load_workbook(path)
sheet = wb.active

print(f"B19 value: {sheet['B19'].value}")
for m_rng in sheet.merged_cells.ranges:
    if "B19" in str(m_rng):
        print(f"B19 is part of merge: {m_rng}")

wb.close()
