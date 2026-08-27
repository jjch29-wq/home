import pandas as pd

df_backup = pd.read_excel('temp_test.xlsx', sheet_name='DailyUsage')
xlsx_path = r'c:\Users\jjch2\Desktop\PMI\home\data\Material_Inventory.xlsx'

with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_backup.to_excel(writer, sheet_name='DailyUsage', index=False)

print(f'Recovered {len(df_backup)} records!')
