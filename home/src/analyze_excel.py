import pandas as pd
import json

file_path = r'C:\Users\jjch2\Desktop\26년07월 월간용역진도보고서.xlsx'
out_path = r'C:\Users\jjch2\Desktop\PMI\home\src\excel_summary.md'

try:
    xl = pd.ExcelFile(file_path)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(f"# Excel File: {file_path}\n\n")
        f.write(f"**Sheets**: {xl.sheet_names}\n\n")
        
        for sheet in xl.sheet_names:
            f.write(f"## Sheet: {sheet}\n")
            df = xl.parse(sheet)
            f.write(f"**Original Shape**: {df.shape}\n")
            # Drop empty cols and rows
            df = df.dropna(axis=1, how='all')
            df = df.dropna(axis=0, how='all')
            f.write(f"**Cleaned Shape**: {df.shape}\n")
            f.write("### First 20 valid rows:\n")
            f.write(df.head(20).to_markdown(index=False))
            f.write("\n\n")
except Exception as e:
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(f"Error: {e}")
