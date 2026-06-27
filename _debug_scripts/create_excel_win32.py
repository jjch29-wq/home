import win32com.client as win32
import os

try:
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Add()
    ws = wb.ActiveSheet
    ws.Name = "투입인원 명단"

    # Title
    ws.Range("A1:F1").Merge()
    ws.Range("A1").Value = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역 투입인원 명단"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = -4108 # xlCenter
    ws.Range("A1").VerticalAlignment = -4108 # xlCenter
    ws.Rows(1).RowHeight = 40

    # Headers
    headers = ["연번", "직책 (담당분야)", "성명", "생년월일", "서명 (인)", "비고"]
    for col, header in enumerate(headers, 1):
        cell = ws.Cells(3, col)
        cell.Value = header
        cell.Font.Bold = True
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108

    # Widths
    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 25
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 20
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 25

    # Data
    sample_data = [
        [1, "총괄 책임자", "", "", "", ""],
        [2, "방사선투과검사(RT)", "", "", "", ""],
        [3, "초음파탐상검사(UT)", "", "", "", ""],
        [4, "자기탐상검사(MT)", "", "", "", ""],
        [5, "침투탐상검사(PT)", "", "", "", ""],
    ]

    for i in range(10):
        if i < len(sample_data):
            row_data = sample_data[i]
        else:
            row_data = [i+1, "", "", "", "", ""]
        
        for col, val in enumerate(row_data, 1):
            cell = ws.Cells(i+4, col)
            cell.Value = val
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108

    # Borders
    rng = ws.Range("A3:F13")
    for border_id in [7, 8, 9, 10, 11, 12]:
        rng.Borders(border_id).LineStyle = 1
        rng.Borders(border_id).Weight = 2

    # Row heights
    for r in range(3, 14):
        ws.Rows(r).RowHeight = 30

    save_path = r"c:\Users\-\OneDrive\바탕 화면\home\투입인원_명단.xlsx"
    if os.path.exists(save_path):
        os.remove(save_path)
    wb.SaveAs(save_path)
    wb.Close()
    print(f"SUCCESS: {save_path}")
except Exception as e:
    print(f"ERROR: {e}")
finally:
    try:
        excel.Quit()
    except:
        pass
