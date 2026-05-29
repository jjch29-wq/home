import pandas as pd

fp = r"Na-aba/home/data/RT KS양식.xlsx"
try:
    xls = pd.ExcelFile(fp)
    print("Sheets in RT KS양식.xlsx:", xls.sheet_names)
    for sheet_name in xls.sheet_names[:2]:
        df = pd.read_excel(fp, sheet_name=sheet_name, header=None)
        print(f"\n--- Sheet: {sheet_name} (Shape: {df.shape}) ---")
        for idx, row in df.iterrows():
            if idx > 30: break
            non_nulls = [f"{i}: {v}" for i, v in enumerate(row.values) if pd.notna(v)]
            if non_nulls:
                print(f"Row {idx}: {non_nulls}")
except Exception as e:
    print("Error:", e)
