import json
import math

config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'
with open(config_path, 'r', encoding='utf-8') as f:
    c = json.load(f)

MATERIAL_COST = c["MATERIAL_COST"]
LABOR_COST = c["LABOR_COST"]
CONTRACT_QTY = c["CONTRACT_QTY"]

overhead_rate = 59.7 / 100.0
tech_fee_rate = 5.3 / 100.0

total_contract_amt = 0
details = []

for loc, times in CONTRACT_QTY.items():
    for t_time, mats in times.items():
        for m_key, val in mats.items():
            if val == 0: continue
            
            base_mat = m_key.split('_')[0]
            lab_unit = LABOR_COST[loc][t_time][base_mat]
            
            mat_map = {
                "RT_B": 'RT (B필름: 3⅓"x17")',
                "RT_A": 'RT (A필름: 3⅓"x12")',
                "RT_A2": 'RT (A/2필름: 3⅓"x6")',
                "UT": "UT",
                "PT": "PT"
            }
            mat_unit = MATERIAL_COST.get(mat_map.get(m_key, m_key), 0)
            
            oh = int(lab_unit * overhead_rate)
            tech = int((lab_unit + oh) * tech_fee_rate)
            unit_cost = mat_unit + lab_unit + oh + tech
            
            amt = int(float(val) * unit_cost)
            total_contract_amt += amt
            
            details.append(f"{loc} {t_time} {m_key}: Qty {val} x Unit {unit_cost:,} = {amt:,}")

print("--- Contract Breakdown ---")
for d in details:
    print(d)
print(f"\nTotal NDT Contract Amount: {total_contract_amt:,} KRW")
