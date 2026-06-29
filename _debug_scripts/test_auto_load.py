import json
import os

SCRIPT_DIR = r"c:\Users\jjch2\Desktop\PMI\home\src"
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    CONFIG = json.load(f)

MATERIAL_COST = CONFIG["MATERIAL_COST"]
LABOR_COST = CONFIG["LABOR_COST"]
contract_qtys = CONFIG.get("CONTRACT_QTY", {})

locations = ["수송배관(주배관)", "플랜트(관리소)"]
times = ["일반", "야간", "휴일"]
materials = [("RT_B", 'RT (B)'), ("RT_A", 'RT (A)'), ("RT_A2", 'RT (A/2)'), ("UT", "UT"), ("PT", "PT")]

contract_vars = []
for loc in locations:
    for t_time in times:
        for m_key, m_name in materials:
            contract_vars.append(f"{loc}_{t_time}_{m_key}")

print("Testing auto_load_contract_qty logic:")
for loc in contract_qtys:
    for t_time in contract_qtys[loc]:
        if isinstance(contract_qtys[loc][t_time], dict):
            for mat, val in contract_qtys[loc][t_time].items():
                key = f"{loc}_{t_time}_{mat}"
                if key in contract_vars:
                    try:
                        base_mat = mat.split('_')[0]
                        lab_unit = LABOR_COST[loc][t_time][base_mat]
                        mat_map = {
                            "RT_B": 'RT (B필름: 3⅓"x17")', "RT_A": 'RT (A필름: 3⅓"x12")', "RT_A2": 'RT (A/2필름: 3⅓"x6")',
                            "UT": "UT", "PT": "PT"
                        }
                        mat_unit = MATERIAL_COST.get(mat_map.get(mat, mat), 0)
                        
                        oh = int(lab_unit * 80.0 / 100.0)
                        tech = int((lab_unit + oh) * 5.86 / 100.0)
                        unit_cost = mat_unit + lab_unit + oh + tech
                        
                        amt = float(val) * unit_cost
                        print(f"SUCCESS: {key} -> QTY: {val}, PRICE: {unit_cost}, AMT: {amt}")
                    except Exception as e:
                        print(f"ERROR on {key}: {e}")
        else:
            print("OLD FORMAT DETECTED")
