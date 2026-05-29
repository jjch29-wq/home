import pandas as pd

fp = r"Na-aba/K1-JHC-760-HCS1-006.xlsm"
try:
    xls = pd.ExcelFile(fp)
    # print sheet names
    print("Sheets:", xls.sheet_names)
    # read sheet index 1 (which was _001)
    df = pd.read_excel(fp, sheet_name=xls.sheet_names[1], header=None)
    print("\nRows 5 to 11, Cols 9 to 16:")
    for r in range(5, 12):
        row_vals = [f"Col {c}: {df.iloc[r, c]}" for c in range(9, 16) if c < df.shape[1]]
        print(f"Row {r}: {row_vals}")
except Exception as e:
    print("Error:", e)
