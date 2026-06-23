import json
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    config_data = json.load(f)

config_data["CONTRACT_QTY"] = {
    "수송배관(주배관)": {
        "RT_B": 19125,
        "RT_A": 0,
        "RT_A2": 0,
        "UT": 319.02,
        "PT": 319.01
    },
    "플랜트(관리소)": {
        "RT_B": 1243,
        "RT_A": 2464,
        "RT_A2": 1704,
        "UT": 0,
        "PT": 19.62
    }
}

with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
    json.dump(config_data, f, ensure_ascii=False, indent=4)

print("Added CONTRACT_QTY to config.json")
