import pandas as pd
import os

db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'

if not os.path.exists(db_path):
    print(f"Error: {db_path} not found.")
    exit(1)

try:
    df = pd.read_excel(db_path, sheet_name='DailyUsage')
    # Filter by date range
    df['Date'] = pd.to_datetime(df['날짜'], errors='coerce')
    start_date = '2026-01-01'
    end_date = '2026-02-21'
    
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    filtered_df = df.loc[mask]
    
    unique_sites = filtered_df['현장'].dropna().unique()
    
    print(f"Date Range: {start_date} to {end_date}")
    print(f"Total entries: {len(filtered_df)}")
    print(f"Unique Sites Count: {len(unique_sites)}")
    print("Sites List:")
    for site in unique_sites:
        print(f" - {site}")

except Exception as e:
    print(f"Error: {e}")
