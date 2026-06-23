import json

# Pipeline (수송배관)
# RT: 45126 (from previous)
# UT: 43734 (from previous)
# PT: 9142000 / 237.66 = 38466

# Plant (관리소)
# RT: 40438 (from previous)
# UT: 43734 (fallback)
# PT: 758000 / 15.4 = 49220

config_data = {
    "MATERIAL_COST": {
        "RT (B필름: 3⅓\"x17\")": 3379,
        "RT (A필름: 3⅓\"x12\")": 2540,
        "RT (A/2필름: 3⅓\"x6\")": 1515,
        "UT": 1115,
        "PT": 3971
    },
    "LABOR_COST": {
        "수송배관(주배관)": {
            "일반": {"RT": 45126, "UT": 43734, "PT": 38466},
            "야간": {"RT": 67689, "UT": 65601, "PT": 53457},
            "휴일": {"RT": 67689, "UT": 65601, "PT": 53429}
        },
        "플랜트(관리소)": {
            "일반": {"RT": 40438, "UT": 43734, "PT": 49220},
            "야간": {"RT": 60657, "UT": 65601, "PT": 70142},
            "휴일": {"RT": 60657, "UT": 65601, "PT": 68720}
        }
    }
}

with open(r'C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\config.json', 'w', encoding='utf-8') as f:
    json.dump(config_data, f, ensure_ascii=False, indent=4)

print("config.json updated with corrected PT rates from user image.")
