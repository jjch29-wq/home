import pandas as pd
import json
import traceback

def get_clean_budget_data():
    return {
        'LaborDetail': json.dumps({
            "이사": {"personnel": "", "period": "", "unit_price": "230208.33333333334"},
            "부장": {"personnel": "1", "period": "160", "unit_price": "230208.33333333334"},
            "차장": {"personnel": "1", "period": "120", "unit_price": "198625"},
            "과장": {"personnel": "", "period": "", "unit_price": "171541.66666666666"},
            "대리": {"personnel": "1", "period": "160", "unit_price": "158000"},
            "계장": {"personnel": "1", "period": "120", "unit_price": "144458.33333333334"},
            "주임": {"personnel": "", "period": "", "unit_price": "130916.66666666667"},
            "기사": {"personnel": "", "period": "", "unit_price": "121875"},
            "연장근무": {"personnel": "4", "period": "60", "unit_price": "4000"},
            "야간근무": {"personnel": "4", "period": "60", "unit_price": "5000"},
            "휴일근무": {"personnel": "4", "period": "62", "unit_price": "7500"}
        }, ensure_ascii=False),
        
        'MaterialDetail': json.dumps([
            {"qty": "10", "price": "2,200"},
            {"qty": "1", "price": "2,950"},
            {"qty": "10", "price": "2,700"},
            {"qty": "10", "price": "3,000"},
            {"qty": "10", "price": "2,600"},
            {"qty": "490", "price": "1,700"},
            {"qty": "2", "price": "100,000"},
            {"qty": "6", "price": "23,000"},
            {"qty": "6", "price": "23,000"},
            {"qty": "6", "price": "3,500"}
        ], ensure_ascii=False),
        
        'ExpenseDetail': json.dumps({
            "site_expense": [
                {"cat": "차량유지비", "cont": "주유, 수리, 통행, 주차 등", "ppl": "N/A", "qty": "12", "unit": "개월", "price": "200,000", "amount": "2,400,000"},
                {"cat": "소모품비", "cont": "장갑,일회용 작업복외", "ppl": "N/A", "qty": "12", "unit": "개월", "price": "100,000", "amount": "1,200,000"},
                {"cat": "복리후생비", "cont": "생수, 음료 외 기타", "ppl": "N/A", "qty": "12", "unit": "개월", "price": "50,000", "amount": "600,000"},
                {"cat": "Se-175", "cont": "방사성동위원소 구매", "ppl": "N/A", "qty": "1", "unit": "대", "price": "10,300,000", "amount": "10,300,000"}
            ],
            "rental": [
                {"cat": "", "spec": "", "qty": "", "period": "", "unit": "", "price": "0", "amount": "0"},
                {"cat": "", "spec": "", "qty": "", "period": "", "unit": "", "price": "0", "amount": "0"},
                {"cat": "", "spec": "", "qty": "", "period": "", "unit": "", "price": "0", "amount": "0"}
            ],
            "outsource": [
                {"cat": "케이엔디아이", "work": "방사선투과검사", "count": "0", "price": "0", "amount": "0"},
                {"cat": "", "work": "", "count": "0", "price": "0", "amount": "0"},
                {"cat": "", "work": "", "count": "0", "price": "0", "amount": "0"}
            ],
            "depreciation": [
                {"item": "PAUT 장비", "spec": "", "life": "5", "qty": "1", "days": "120", "rate": "44,444", "amount": "5,333,280"},
                {"item": "PAUT SCANNER (MANUAL)", "spec": "", "life": "5", "qty": "1", "days": "120", "rate": "5,556", "amount": "666,720"},
                {"item": "PAUT SCANNER (COBRA)", "spec": "", "life": "5", "qty": "1", "days": "120", "rate": "16,667", "amount": "2,000,040"},
                {"item": "YOKE", "spec": "", "life": "5", "qty": "1", "days": "10", "rate": "222", "amount": "2,220"},
                {"item": "현상용 탑차(5년간 보험비 포함)", "spec": "현장별 차량기입시 탑차 구분 기입", "life": "5", "qty": "1", "days": "120", "rate": "16,667", "amount": "2,000,040"},
                {"item": "스타렉스(5년간 보험비 포함)", "spec": "현장별 차량기입시 스타렉스 구분 기입", "life": "5", "qty": "2", "days": "120", "rate": "16,667", "amount": "4,000,080"}
            ]
        }, ensure_ascii=False),
        
        'LaborCost': 107303333,
        'MaterialCost': 11737950 - 10300000,  # Total material minus Se-175 which is in expense
        'Expense': 4200000 + 10300000,        # Original expense + Se-175
        'Revenue': 288268000
    }

file_path = 'c:/Users/-/PMI/home/src/site_apps/central/data/Material_Inventory.xlsx'

try:
    print("Reading Excel...")
    sheets = pd.read_excel(file_path, sheet_name=None)
    
    budget_df = sheets['Budget']
    site = '중앙지사'
    
    if site in budget_df['Site'].values:
        idx = budget_df[budget_df['Site'] == site].index[0]
        updates = get_clean_budget_data()
        
        for k, v in updates.items():
            budget_df.loc[idx, k] = v
            
        # Recalculate Profit
        rev = float(budget_df.loc[idx, 'Revenue'] or 0)
        lab = float(budget_df.loc[idx, 'LaborCost'] or 0)
        mat = float(budget_df.loc[idx, 'MaterialCost'] or 0)
        exp = float(budget_df.loc[idx, 'Expense'] or 0)
        out = float(budget_df.loc[idx, 'OutsourceCost'] or 0)
        budget_df.loc[idx, 'Profit'] = rev - (lab + mat + exp + out)
        
        print("Updated Budget for '중앙지사'. Writing back to Excel...")
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for s_name, s_df in sheets.items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)
                
        print("Done!")
    else:
        print("Site '중앙지사' not found in Budget sheet.")
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
