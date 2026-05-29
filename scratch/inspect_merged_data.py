import pandas as pd

df = pd.read_excel(r".\Na-aba\Final_Smart_Merged_v2.8_220225.xlsx")
print("Columns:", list(df.columns))
print("First 30 rows:")
print(df.head(30).to_string())
