import pandas as pd
import sys

file_path = r"C:\Users\jjch2\Desktop\4.4.1_위험성평가표(RT_표준양식).xlsx"
try:
    xl = pd.ExcelFile(file_path)
    output_lines = []
    output_lines.append(f"Sheets: {xl.sheet_names}")
    
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)
        output_lines.append(f"\n--- Sheet: {sheet} ---")
        output_lines.append(f"Shape: {df.shape}")
        
        # Write first 25 rows and all columns as CSV
        output_lines.append(df.iloc[:25, :].to_csv(index=False, header=False))
        
    with open(r"C:\Users\jjch2\Desktop\PMI\home\src\read_excel_output.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
        
    print("Done")
except Exception as e:
    print("Error:", e)
    sys.exit(1)
