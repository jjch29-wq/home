import pandas as pd, os, glob

folder = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'
files = sorted(glob.glob(folder + r'\Final_Smart_Merged_v2.8_*.xlsx'))
latest = max(files, key=os.path.getmtime)
print('Latest file:', os.path.basename(latest))

df = pd.read_excel(latest)
print('Columns:', list(df.columns))
print('Total rows:', len(df))
print()

subtotals = df[df.astype(str).apply(lambda r: r.str.contains('Sub-Total', case=False, na=False)).any(axis=1)]
print('=== All Sub-Totals ===')
print(subtotals.head(10).to_string())
print(f'Total Sub-Total rows: {len(subtotals)}')
