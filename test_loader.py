import pandas as pd
import os

db_path = r'c:\Users\jjch2\Desktop\PMI\home\data\Material_Inventory.xlsx'
df = pd.read_excel(db_path, sheet_name='DailyUsage', 
                   dtype={'Site': str, 'Note': str, 'User': str,
                          '차량번호': str, '주행거리': str, '차량점검': str, '차량비고': str,
                          'MaterialID': object})
print(f"Read {len(df)} records directly.")

import sys
sys.path.append(r'c:\Users\jjch2\Desktop\PMI\home\src')
from services.data_loader import DataLoader
loader = DataLoader(db_path)
loader.load_data()
print(f"DataLoader loaded {len(loader.daily_usage_df)} records.")
