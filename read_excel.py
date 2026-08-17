import pandas as pd

file_path = r"C:\Users\jjch2\Desktop\누적진도보고서_202708.xlsx"

try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheet names: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        print(f"\n--- Sheet: {sheet} ---")
        df = xl.parse(sheet)
        print(f"Shape: {df.shape}")
        print("Columns:")
        print(df.columns.tolist())
        print("Head:")
        print(df.head(3))
except Exception as e:
    print(f"Error reading file: {e}")
