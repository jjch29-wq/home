import pandas as pd

fp = r"Na-aba/home/data/가스공사 의뢰서.xlsx"
try:
    df = pd.read_excel(fp, sheet_name="6월5일-1", header=None)
    print("Shape:", df.shape)
    for idx, row in df.iterrows():
        non_nulls = [f"{i}: {v}" for i, v in enumerate(row.values) if pd.notna(v)]
        if non_nulls:
            print(f"Row {idx}: {non_nulls}")
except Exception as e:
    print("Error:", e)
