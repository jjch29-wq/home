import sys
import os
import pandas as pd

def check_duplicates():
    db_path = r"c:\Users\-\OneDrive\바탕 화면\home\data\Material_Inventory.xlsx"
    if os.path.exists(db_path):
        df = pd.read_excel(db_path, sheet_name='Materials')
        for name in ['PT약품', 'MT약품']:
            matches = df[df['품목명'].astype(str).str.strip().str.upper() == name]
            print(f"\nFound {len(matches)} matches for {name}:")
            for i, row in matches.iterrows():
                print(f"ID: {row.get('MaterialID')}, Name: {row.get('품목명')}, Model: {row.get('모델명')}, Spec: {row.get('규격')}, Active: {row.get('Active')}, Qty: {row.get('수량')}")
    else:
        print("DB Not Found")

if __name__ == "__main__":
    check_duplicates()
