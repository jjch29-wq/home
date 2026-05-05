import sys
import os
import pandas as pd
import json

# Add current directory to path
sys.path.append(os.getcwd())

def check_sync():
    db_path = r"c:\Users\-\OneDrive\바탕 화면\home\data\Material_Inventory.xls"
    config_path = r"c:\Users\-\.gemini\antigravity\tab_config.json" # Might be different
    
    # Try to find config
    if not os.path.exists(config_path):
        # Look for it in the app dir
        config_path = r"c:\Users\-\OneDrive\바탕 화면\home\src\tab_config.json"
    
    print(f"DB Path: {db_path}")
    if os.path.exists(db_path):
        df = pd.read_excel(db_path, sheet_name='Materials')
        print(f"Materials in DB: {len(df)}")
        print(f"Columns: {df.columns.tolist()}")
        print(f"First 5 items: {df['품목명'].head().tolist()}")
    else:
        print("DB not found!")

    print(f"\nConfig Path: {config_path}")
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
            materials = cfg.get('materials', [])
            print(f"Preferred Materials in Config: {len(materials)}")
            print(f"Items: {materials}")
    else:
        print("Config not found!")

if __name__ == "__main__":
    check_sync()
