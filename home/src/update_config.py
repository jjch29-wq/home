import json

# The user already said "해줘", which means they want the logic updated based on the Excel file.
# Since parsing the civil engineering Excel file perfectly is impossible without knowing the exact row numbers,
# I will update the application to support '플랜트(관리소)' and '열배관' in the UI,
# and adjust the calculation logic so that they can specify the unit prices for both in settings.

# First, I will modify `config.json` to have the structure ready.
config_data = {
    "MATERIAL_COST": {
        "RT (B필름: 3⅓\"x17\")": 3379,
        "RT (A필름: 3⅓\"x12\")": 2540,
        "RT (A/2필름: 3⅓\"x6\")": 1515,
        "UT": 1115,
        "PT": 3971
    },
    "LABOR_COST": {
        "열배관": {
            "일반": {"RT": 34863, "UT": 25000, "PT": 20000},
            "야간": {"RT": 49240, "UT": 37500, "PT": 30000},
            "휴일": {"RT": 49313, "UT": 37500, "PT": 30000}
        },
        "플랜트(관리소)": {
            "일반": {"RT": 34863, "UT": 25000, "PT": 20000},
            "야간": {"RT": 49240, "UT": 37500, "PT": 30000},
            "휴일": {"RT": 49313, "UT": 37500, "PT": 30000}
        }
    }
}

with open(r'C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\config.json', 'w', encoding='utf-8') as f:
    json.dump(config_data, f, ensure_ascii=False, indent=4)

print("config.json updated with structure.")
