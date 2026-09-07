import pandas as pd
import json

df_dl = pd.read_excel('C:/Users/-/Downloads/공사실행예산서(중앙지사) (4).xlsx', sheet_name='사전원가')
df_inv = pd.read_excel('c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx', sheet_name='Budget')
central_row = df_inv[df_inv['Site'] == '중앙지사'].iloc[0]

print("=== App Budget Data ===")
print("LaborCost:", central_row['LaborCost'])
print("MaterialCost:", central_row['MaterialCost'])
print("Expense:", central_row['Expense'])
print("Profit:", central_row['Profit'])
print("\nLaborDetail:")
print(json.dumps(json.loads(central_row['LaborDetail']), ensure_ascii=False, indent=2))
print("\nMaterialDetail:")
print(json.dumps(json.loads(central_row['MaterialDetail']), ensure_ascii=False, indent=2))
print("\nExpenseDetail:")
print(json.dumps(json.loads(central_row['ExpenseDetail']), ensure_ascii=False, indent=2))
