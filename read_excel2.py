import pandas as pd
import sys

file_path = r"C:\Users\jjch2\Desktop\누적진도보고서_202708.xlsx"
out_path = r"C:\Users\jjch2\Desktop\PMI\excel_output.txt"

with open(out_path, "w", encoding="utf-8") as f:
    try:
        xl = pd.ExcelFile(file_path)
        f.write(f"Sheet names: {xl.sheet_names}\n")
        for sheet in xl.sheet_names:
            f.write(f"\n--- Sheet: {sheet} ---\n")
            df = xl.parse(sheet, header=None) # Read without header to see raw structure
            f.write(f"Shape: {df.shape}\n")
            f.write("First 15 rows:\n")
            f.write(df.head(15).to_string(index=True))
            f.write("\n")
    except Exception as e:
        f.write(f"Error reading file: {e}\n")
