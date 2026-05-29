import pandas as pd

fp = r"Na-aba/home/data/SISGI-KOGAS-RT-001.xlsx"
try:
    xls = pd.ExcelFile(fp)
    print("Sheets in SISGI-KOGAS-RT-001.xlsx:", xls.sheet_names)
    df = pd.read_excel(fp, sheet_name=xls.sheet_names[1], skiprows=8)
    print("Sheet Name:", xls.sheet_names[1])
    print("Columns:", list(df.columns))
    print("\nNon-empty rows:")
    df_clean = df.dropna(how='all')
    print(df_clean.iloc[:20, :11].to_string())
except Exception as e:
    print("Error:", e)
