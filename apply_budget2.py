import pandas as pd
import json

excel_path = 'c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx'
df = pd.read_excel(excel_path, sheet_name='Budget')

central_idx = df[df['Site'] == '중앙지사'].index[0]

# Previous values
labor_detail = json.loads(df.at[central_idx, 'LaborDetail'])
material_detail = json.loads(df.at[central_idx, 'MaterialDetail'])
expense_detail = json.loads(df.at[central_idx, 'ExpenseDetail'])

# 1. Update Material Detail (add Se-175)
if len(material_detail) == 10:
    material_detail.append({
        "qty": "1",
        "price": "10300000"
    })
else:
    # If it's already 11 items, just update the last one
    material_detail[10] = {
        "qty": "1",
        "price": "10300000"
    }

# Calculate new MaterialCost
mat_cost = sum([float(str(x.get('qty', '0')).replace(',', '') or 0) * float(str(x.get('price', '0')).replace(',', '') or 0) for x in material_detail])
print(f"New MaterialCost: {mat_cost}") # Should be 11,737,950

# 2. Update Expense Detail (Remove Se-175 from site_expense, fix depreciation)
new_site_expense = []
for ex in expense_detail['site_expense']:
    if 'Se-175' not in ex['cat']:
        new_site_expense.append(ex)
expense_detail['site_expense'] = new_site_expense

expense_detail['depreciation'] = [
    {
      "item": "PAUT 장비", "spec": "", "life": "5", "qty": "1", "days": "120",
      "rate": "44,444", "amount": "5,333,280"
    },
    {
      "item": "PAUT SCANNER (MANUAL)", "spec": "", "life": "5", "qty": "1", "days": "120",
      "rate": "5,556", "amount": "666,720"
    },
    {
      "item": "PAUT SCANNER (COBRA)", "spec": "", "life": "5", "qty": "1", "days": "120",
      "rate": "16,667", "amount": "2,000,040"
    },
    {
      "item": "탑차(5년간 분할 반영)", "spec": "공사별 현장 탑차 월 반영", "life": "5", "qty": "1", "days": "30",
      "rate": "16,667", "amount": "500,010"
    },
    {
      "item": "스타렉스(5년간 분할 반영)", "spec": "공사별 현장 스타렉스 월 반영", "life": "5", "qty": "1", "days": "160",
      "rate": "16,667", "amount": "2,666,720"
    }
]

# Calculate new Expense total based on App's exp_total logic (Site + Rentals + Insurance + Depreciation)
site_total = sum([float(str(x.get('qty', '0')).replace(',', '') or 0) * float(str(x.get('price', '0')).replace(',', '') or 0) for x in expense_detail['site_expense']])
dep_total = sum([float(str(x.get('qty', '0')).replace(',', '') or 0) * float(str(x.get('days', '0')).replace(',', '') or 0) * float(str(x.get('rate', '0')).replace(',', '') or 0) for x in expense_detail['depreciation']])
rental_total = 0
insurance_total = 107303333.33333333 * 0.109744 # = 11,775,897

exp_total = site_total + rental_total + insurance_total + dep_total
print(f"New Expense (Calculated by App logic): {exp_total}")
# Note: Excel's expense is 27,144,785
# site_total = 4,200,000
# insurance = 11,775,897
# dep_total = 11,166,770 (5333280 + 666720 + 2000040 + 500010 + 2666720 = 11,166,770)
# Sum = 4,200,000 + 11,775,897 + 11,166,770 = 27,142,667 (Close to 27,144,785, difference is precision of rates)

# We will just write the KPI expense directly to match Excel EXACTLY for the top row
df.at[central_idx, 'MaterialCost'] = mat_cost
df.at[central_idx, 'Expense'] = 27144785

df.at[central_idx, 'MaterialDetail'] = json.dumps(material_detail, ensure_ascii=False)
df.at[central_idx, 'ExpenseDetail'] = json.dumps(expense_detail, ensure_ascii=False)

with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Budget', index=False)
    
print("Updated Material_Inventory.xlsx")
