import json
import os

config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'

try:
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Apply the correct RT_A and RT_A2 quantities for 플랜트(관리소)
    if 'CONTRACT_QTY' in config and '플랜트(관리소)' in config['CONTRACT_QTY']:
        plant = config['CONTRACT_QTY']['플랜트(관리소)']
        
        # Day (일반)
        if '일반' in plant:
            plant['일반']['RT_A'] = 1969
            plant['일반']['RT_A2'] = 1359
            
        # Night (야간)
        if '야간' in plant:
            plant['야간']['RT_A'] = 256
            plant['야간']['RT_A2'] = 168
            
        # Holiday (휴일)
        if '휴일' in plant:
            plant['휴일']['RT_A'] = 239
            plant['휴일']['RT_A2'] = 177

    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

    print('Config updated for RT_A and RT_A2 successfully.')
except Exception as e:
    print('Error:', e)
