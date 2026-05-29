import pandas as pd

df = pd.read_excel(r".\Na-aba\Final_Smart_Merged_v2.8_220225.xlsx")
print("Total rows:", len(df))
print("Unique Dwgs:")
print(df['Dwg'].dropna().unique())
print("\nTHK value counts:")
print(df['THK'].value_counts(dropna=False))
