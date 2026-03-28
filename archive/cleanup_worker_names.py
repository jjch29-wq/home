"""
Database Cleanup Script - Remove Shift Information from Worker Names
This script cleans up worker names in the database by extracting only the actual names
from entries like "(주간) 김진환" to just "김진환"
"""
import pandas as pd
import re
import os

# Path to the database
db_path = 'Material_Inventory.xlsx'

if not os.path.exists(db_path):
    print(f"Error: Database file '{db_path}' not found!")
    exit(1)

print("Loading database...")
# Load the DailyUsage sheet
daily_usage_df = pd.read_excel(db_path, sheet_name='DailyUsage', engine='openpyxl', dtype={'Site': str, 'Note': str, 'User': str})

print(f"Total records: {len(daily_usage_df)}")

# Define user columns
user_cols = ['User', 'User2', 'User3', 'User4', 'User5', 'User6', 'User7', 'User8', 'User9', 'User10']

# Function to extract actual name from "(Shift) Name" format
def clean_worker_name(name):
    if pd.isna(name) or not str(name).strip():
        return ''
    
    name_str = str(name).strip()
    
    # Check if it matches the pattern "(주간/야간/휴일) Name"
    match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", name_str)
    if match:
        actual_name = match.group(2).strip()
        return actual_name if actual_name else ''
    
    # If it's just a shift marker without a name, return empty
    if re.match(r"^\((주간|야간|휴일)\)$", name_str):
        return ''
    
    # Otherwise return as is
    return name_str

# Clean all user columns
print("\nCleaning worker names...")
changes_made = 0
changes_log = []

for col in user_cols:
    if col in daily_usage_df.columns:
        # Create a new cleaned series
        cleaned_series = daily_usage_df[col].apply(clean_worker_name)
        
        # Check for changes
        for idx in daily_usage_df.index:
            original = str(daily_usage_df.loc[idx, col]) if pd.notna(daily_usage_df.loc[idx, col]) else ''
            cleaned = cleaned_series.loc[idx]
            if original != cleaned:
                changes_log.append((original, cleaned))
                changes_made += 1
        
        # Replace the entire column with cleaned data
        daily_usage_df[col] = cleaned_series

# Print changes
for original, cleaned in changes_log:
    print(f"  Changed: '{original}' -> '{cleaned}'")


print(f"\nTotal changes made: {changes_made}")

if changes_made > 0:
    print("\nSaving updated database...")
    try:
        # Use openpyxl to handle the Excel file more robustly
        from openpyxl import load_workbook
        
        # First, save all dataframes to a temporary Excel file
        temp_path = db_path + '.tmp'
        
        # Load all sheets to preserve them
        try:
            materials_df = pd.read_excel(db_path, sheet_name='Materials', engine='openpyxl')
            transactions_df = pd.read_excel(db_path, sheet_name='Transactions', engine='openpyxl')
            monthly_usage_df = pd.read_excel(db_path, sheet_name='MonthlyUsage', engine='openpyxl')
        except Exception as e:
            print(f"Error loading sheets: {e}")
            print("The Excel file may be open in another program. Please close it and run this script again.")
            exit(1)
        
        # Write to temporary file first
        with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
            materials_df.to_excel(writer, sheet_name='Materials', index=False)
            transactions_df.to_excel(writer, sheet_name='Transactions', index=False)
            monthly_usage_df.to_excel(writer, sheet_name='MonthlyUsage', index=False)
            daily_usage_df.to_excel(writer, sheet_name='DailyUsage', index=False)
        
        # Replace original file with temporary file
        import shutil
        shutil.move(temp_path, db_path)
        
        print("Database updated successfully!")
        
    except PermissionError:
        print("\n" + "="*60)
        print("ERROR: Cannot save the database file!")
        print("The file is probably open in another program.")
        print("Please close the MaterialManager application and try again.")
        print("="*60)
        exit(1)
    except Exception as e:
        print(f"\nError saving database: {e}")
        print("Please close all applications that might be using the Excel file and try again.")
        exit(1)
    
    # Show unique worker names after cleanup
    print("\nUnique worker names after cleanup:")
    all_users = set()
    for col in user_cols:
        if col in daily_usage_df.columns:
            all_users.update([u for u in daily_usage_df[col].dropna().unique().tolist() if u])
    
    for user in sorted(all_users):
        print(f"  - {user}")
else:
    print("\nNo changes needed - database is already clean!")
