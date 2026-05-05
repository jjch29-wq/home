import pandas as pd
import os

db_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\Material_Inventory.xlsx"
if not os.path.exists(db_path):
    # Try alternate path
    db_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\Material_Inventory.xlsx"

if os.path.exists(db_path):
    try:
        df = pd.read_excel(db_path, sheet_name='DailyUsage')
        print(f"Total records: {len(df)}")
        print("Last 5 records:")
        cols = ['Date', 'Site', 'MaterialID', 'Usage', '차량번호', '차량점검', 'EntryTime', '입력시간']
        existing_cols = [c for c in cols if c in df.columns]
        print(df[existing_cols].tail(5))
    except Exception as e:
        print(f"Error: {e}")
else:
    print(f"DB not found at {db_path}")
