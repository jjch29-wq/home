import pandas as pd
import os

path = '../data/Material_Inventory.xlsx'
if not os.path.exists(path):
    print(f"File not found: {path}")
    exit(1)

m_df = pd.read_excel(path, sheet_name='Materials')
d_df = pd.read_excel(path, sheet_name='DailyUsage')

print("MAT COLS:", m_df.columns.tolist())
print("DAILY COLS:", d_df.columns.tolist())

if not m_df.empty:
    # Filter for Carestream or NDT drugs
    subset = m_df[m_df['품목명'].astype(str).str.contains('Carestream|PT약품|MT약품', na=False, case=False)]
    print("MAT SUBSET:")
    print(subset[['MaterialID', '품목명', '모델명', '수량']].head(10).to_string())
else:
    print("MAT EMPTY")

if not d_df.empty:
    print("DAILY USAGE LAST 5:")
    print(d_df.tail(5).to_string())
else:
    print("DAILY USAGE EMPTY")
