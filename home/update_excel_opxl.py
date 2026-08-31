import openpyxl

excel_path = r"C:\Users\-\OneDrive\바탕 화면\템플릿 가산가평월간용역진도보고서.xlsx"
wb = openpyxl.load_workbook(excel_path)

# 2. 공정율 (주배관)
ws_main = wb["2. 공정율 (주배관)"]
ws_main["C9"].value = 14205
ws_main["C10"].value = 4920
ws_main["C12"].value = 237.66
ws_main["C13"].value = 81.36
ws_main["C15"].value = 237.66
ws_main["C16"].value = 81.35
ws_main["C22"].value = 19125
ws_main["C23"].value = 0
ws_main["C24"].value = 0

# 2. 공정율 (관리소)
ws_station = wb["2. 공정율 (관리소)"]
ws_station["C9"].value = 4314
ws_station["C10"].value = 1097
ws_station["C12"].value = 0
ws_station["C13"].value = 0
ws_station["C15"].value = 15.4
ws_station["C16"].value = 4.22
ws_station["C22"].value = 1243
ws_station["C23"].value = 2464
ws_station["C24"].value = 1704

# 2. 공정율 (전체)
ws_total = wb["2. 공정율 (전체)"]
ws_total["C9"].value = 18519
ws_total["C10"].value = 6017
ws_total["C12"].value = 237.66
ws_total["C13"].value = 81.36
ws_total["C15"].value = 253.06
ws_total["C16"].value = 85.57
ws_total["C22"].value = 20368
ws_total["C23"].value = 2464
ws_total["C24"].value = 1704

wb.save(excel_path)
print("Successfully updated the Excel file using openpyxl.")
