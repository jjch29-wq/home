import sys
import os
import pandas as pd
import json

def check_materials():
    db_path = r"c:\Users\-\OneDrive\바탕 화면\home\data\Material_Inventory.xlsx"
    if os.path.exists(db_path):
        df = pd.read_excel(db_path, sheet_name='Materials')
        print(f"Total Materials: {len(df)}")
        # Check specific items
        targets = [
            'Carestream T200-10*12"',
            'Carestream T200-14*17"',
            'Carestream T200-3⅓*12"',
            'Carestream T200-3⅓*17"',
            'Carestream T200-4½*12"',
            'MT약품',
            'PT약품'
        ]
        
        print("\nChecking Target Items in DB:")
        for t in targets:
            # Case insensitive search
            match = df[df['품목명'].astype(str).str.strip().str.lower() == t.lower()]
            if not match.empty:
                active = match.iloc[0].get('Active', 'N/A')
                print(f"MATCH: {t} -> Active: {active}, Stock: {match.iloc[0].get('수량', 'N/A')}")
            else:
                print(f"MISSING: {t}")
    else:
        print("DB Not Found")

if __name__ == "__main__":
    check_materials()
