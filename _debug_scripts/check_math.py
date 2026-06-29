import sys
import os
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HOME_DIR = os.path.dirname(SCRIPT_DIR)
SRC_DIR = os.path.join(HOME_DIR, "home", "src")
sys.path.insert(0, SRC_DIR)

from ndt_billing_tab import load_config

CONFIG = load_config()
MATERIAL_COST = CONFIG["MATERIAL_COST"]
LABOR_COST = CONFIG["LABOR_COST"]
contract_qtys = CONFIG.get("CONTRACT_QTY", {})

print("| 구분 | 시간 | 항목 | 자동입력 수량 | 자동계산 단가(원) | 계약금액(원) |")
print("| :--- | :--- | :--- | :---: | :---: | :---: |")

for loc in contract_qtys:
    for t_time in contract_qtys[loc]:
        if isinstance(contract_qtys[loc][t_time], dict):
            for mat, val in contract_qtys[loc][t_time].items():
                qty = float(val)
                if qty == 0: continue
                
                base_mat = mat.split('_')[0]
                lab_unit = LABOR_COST[loc][t_time][base_mat]
                mat_map = {
                    "RT_B": 'RT (B필름: 3⅓"x17")', "RT_A": 'RT (A필름: 3⅓"x12")', "RT_A2": 'RT (A/2필름: 3⅓"x6")',
                    "UT": "UT", "PT": "PT"
                }
                mat_unit = MATERIAL_COST.get(mat_map.get(mat, mat), 0)
                
                oh1 = int(lab_unit * 0.8)
                tech1 = int((lab_unit + oh1) * 0.0586)
                unit1 = mat_unit + lab_unit + oh1 + tech1
                amt1 = unit1 * qty
                
                print(f"| {loc} | {t_time} | {mat} | {qty:,.2f} | {unit1:,} | {int(amt1):,} |")



