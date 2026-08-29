import pandas as pd

def check_excel(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='DailyUsage')
        last_10 = df.tail(10)
        with open('excel_dump.txt', 'a', encoding='utf-8') as f:
            f.write(f"\n--- {file_path} ---\n")
            f.write(last_10.to_string())
    except Exception as e:
        with open('excel_dump.txt', 'a', encoding='utf-8') as f:
            f.write(f"\n--- Error reading {file_path}: {e} ---\n")

with open('excel_dump.txt', 'w', encoding='utf-8') as f:
    f.write("Excel Dump\n")

check_excel('home/data/Material_Inventory.xlsx')
check_excel('home/data/Kogas_Material_Inventory.xlsx')
