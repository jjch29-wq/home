import pandas as pd
import os

file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\Material_Inventory.xlsx'
if os.path.exists(file_path):
    df = pd.read_excel(file_path, sheet_name='DailyUsage')
    print("Columns:", list(df.columns))
    print("\nTail data:")
    print(df.tail(5).to_string())
else:
    print("File not found")
