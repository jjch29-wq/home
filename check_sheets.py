import pandas as pd
excel_file = pd.ExcelFile('home/data/Material_Inventory.xlsx')
print("Material_Inventory.xlsx sheets:", excel_file.sheet_names)
excel_file2 = pd.ExcelFile('home/data/Kogas_Material_Inventory.xlsx')
print("Kogas_Material_Inventory.xlsx sheets:", excel_file2.sheet_names)
