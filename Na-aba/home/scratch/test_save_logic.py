import os
import pandas as pd
import time
import datetime

db_path = "test_inventory.xlsx"

# Helper to simulate show_error_dialog
def show_error_dialog(title, message):
    print(f"DIALOG: {title}\n{message}")

# Simplified save_data logic for testing
def save_data_logic(materials_df, db_path):
    max_retries = 3
    retry_delay = 0.5
    
    for attempt in range(max_retries):
        try:
            # Simulate lock check
            if os.path.exists(db_path):
                with open(db_path, 'a'): pass
            
            with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
                materials_df.to_excel(writer, sheet_name='Materials', index=False)
            
            print(f"Save successful on attempt {attempt + 1}")
            return True
        except PermissionError as pe:
            print(f"Attempt {attempt + 1} failed: File locked")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                continue
            else:
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                conflict_path = db_path.replace('.xlsx', f'_Conflict_{ts}.xlsx')
                try:
                    with pd.ExcelWriter(conflict_path, engine='openpyxl') as writer:
                        materials_df.to_excel(writer, sheet_name='Materials', index=False)
                    show_error_dialog("데이터 저장 지연/충돌", f"임시 저장됨: {conflict_path}")
                    return True
                except Exception as e2:
                    print(f"Backup failed: {e2}")
                    return False

# Setup dummy data
df = pd.DataFrame({'MaterialID': [1, 2], 'Name': ['A', 'B']})
df.to_excel(db_path, index=False)

print("--- Test 1: Normal Save ---")
save_data_logic(df, db_path)

print("\n--- Test 2: Locked Save ---")
# Manually open file to lock it (simulated by opening in 'r' mode without closing in some OS, 
# but in Python we can use a context manager to hold it open)
f = open(db_path, 'a') 
# Note: On Windows, 'a' mode locks it for others.
try:
    save_data_logic(df, db_path)
finally:
    f.close()

if os.path.exists(db_path): os.remove(db_path)
# Clean up conflict files
for f in os.listdir('.'):
    if 'Conflict' in f: os.remove(f)
