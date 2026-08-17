import pandas as pd

file_path = r"C:\Users\jjch2\Desktop\누적진도보고서_202708.xlsx"
out_path = r"C:\Users\jjch2\Desktop\PMI\excel_summary.txt"

with open(out_path, "w", encoding="utf-8") as f:
    try:
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            f.write(f"\n--- Sheet: {sheet} ---\n")
            df = xl.parse(sheet, header=None)
            
            # Print row by row, skipping rows that are completely NaN
            for i, row in df.iterrows():
                row_vals = [str(x).strip() for x in row if pd.notna(x) and str(x).strip() != ""]
                if row_vals:
                    f.write(f"Row {i}: " + " | ".join(row_vals) + "\n")
    except Exception as e:
        f.write(f"Error reading file: {e}\n")
