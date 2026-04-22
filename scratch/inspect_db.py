import pandas as pd
import os

db_path = r"c:\Users\-\OneDrive\바탕 화면\PMI Report\home\data\Material_Inventory.xlsx"
if os.path.exists(db_path):
    print(f"File found: {db_path}")
    try:
        df = pd.read_excel(db_path, sheet_name='DailyUsage')
        print("Columns in DailyUsage (Unicode Escaped):")
        for c in df.columns:
            print(f"{repr(c)} -> {c.encode('unicode_escape').decode('utf-8')}")
        print("\nData Sample for columns 14-22:")
        print(df.iloc[:, 14:22].head())
    except Exception as e:
        print(f"Error reading Excel: {e}")
else:
    print(f"File NOT found: {db_path}")
