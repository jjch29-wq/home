import pandas as pd

file_path = r"Na-aba/K1-JHC-750-CCS1-027.xlsm"
try:
    xls = pd.ExcelFile(file_path)
    print(f"Sheets: {xls.sheet_names}")
    for sheet_name in xls.sheet_names[1:4]: # print first few sheets (excluding first sheet if it's main index)
        print(f"\n--- Sheet: {sheet_name} ---")
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        # Find non-empty rows up to row 25
        for idx, row in df.iterrows():
            if idx > 25: break
            non_nulls = [f"{i}: {v}" for i, v in enumerate(row.values) if pd.notna(v)]
            if non_nulls:
                print(f"Row {idx}: {non_nulls}")
except Exception as e:
    print(f"Error: {e}")
