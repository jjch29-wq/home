import pandas as pd

df = pd.read_excel(r".\Na-aba\Final_Smart_Merged_v2.8_220225.xlsx")
print("Total rows count:", len(df))
totals = df[df.astype(str).apply(lambda r: r.str.contains("Total|Grand|Sub|소계|합계", case=False).any(), axis=1)]
if len(totals) > 0:
    print("Found total rows:")
    print(totals.to_string())
else:
    print("No total rows found!")
