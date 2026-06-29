import json
config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'
with open(config_path, 'r', encoding='utf-8') as f:
    c = json.load(f)
import pprint
pprint.pprint(c['CONTRACT_QTY'])
