import pandas as pd
import os
import traceback
import re

def diagnose_simulated():
    db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'
    print(f"Simulating loading logic for {db_path}...")
    
    try:
        # Materials
        print("Loading Materials...")
        materials_df = pd.read_excel(db_path, sheet_name='Materials')
        if '품명' in materials_df.columns and '품목명' not in materials_df.columns:
            materials_df.rename(columns={'품명': '품목명'}, inplace=True)
        
        # Transactions
        print("Loading Transactions...")
        transactions_df = pd.read_excel(db_path, sheet_name='Transactions')
        
        # Here is where it likely fails in MaterialManager-10.py line 460
        if not transactions_df.empty:
            print("Processing Transactions...")
            # Simulation of line 460
            if 'MaterialID' not in transactions_df.columns:
                print("WARNING: 'MaterialID' not in transactions_df columns!")
                print("Existing columns:", list(transactions_df.columns))
            
            transactions_df['MaterialID'] = pd.to_numeric(transactions_df['MaterialID'], errors='coerce')
            print("Successfully processed Transactions.")

        # Daily Usage
        print("Loading DailyUsage...")
        daily_usage_df = pd.read_excel(db_path, sheet_name='DailyUsage')
        daily_usage_df.columns = [re.sub(r'\s+', '', str(c)) for c in daily_usage_df.columns]
        
        print("Load complete!")

    except Exception as e:
        print(f"\nFATAL ERROR DURING SIMULATION: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    diagnose_simulated()
