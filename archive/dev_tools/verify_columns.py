import tkinter as tk
from tkinter import ttk
import pandas as pd
import sys
import os

# Add the directory to sys.path to import MaterialManager-4
sys.path.append(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI")

# Mocking necessary parts if import fails or to avoid full app launch
try:
    from importlib.machinery import SourceFileLoader
    mm_module = SourceFileLoader("MaterialManager", r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\MaterialManager-4.py").load_module()
    MaterialManager = mm_module.MaterialManager
except Exception as e:
    print(f"Failed to import MaterialManager: {e}")
    sys.exit(1)

def verify_columns():
    root = tk.Tk()
    app = MaterialManager(root)
    
    # Mock data
    app.materials_df = pd.DataFrame(columns=['Material ID', '품목명', '모델명', 'SN', '규격', '품목군코드', '관리단위', '재고하한', '수량'])
    app.daily_usage_df = pd.DataFrame([
        {
            'Date': '2024-01-01', 'Site': 'TestSite', 'Material ID': 1, 
            'NDT_형광': 10.0, 'NDT_자분': 5.0, # Old data keys
            'Entry Time': '2024-01-01 12:00:00'
        }
    ])
    
    # Setup tabs (partial)
    app.tab_daily_usage = ttk.Frame(app.notebook)
    app.notebook.add(app.tab_daily_usage, text="일일 사용량")
    app.setup_daily_usage_tab()
    
    # 1. Verify Columns in Treeview
    columns = app.daily_usage_tree['columns']
    print(f"Treeview Columns: {columns}")
    
    expected_ndt = ['형광자분', '흑색자분', '백색페인트', '침투제', '세척제', '현상제', '형광침투제']
    
    # Check if expected columns are present
    missing = [col for col in expected_ndt if col not in columns]
    if missing:
        print(f"FAILED: Missing columns: {missing}")
    else:
        print("SUCCESS: All expected NDT columns found.")
        
    # Check order
    # Finding indices
    indices = {col: columns.index(col) for col in expected_ndt if col in columns}
    print(f"Indices: {indices}")
    if indices['형광자분'] < indices['흑색자분'] and indices['현상제'] < indices['형광침투제']:
         print("SUCCESS: Column order seems correct (Fluorescent Mag first, Fluorescent Pen last)")
    else:
         print("FAILED: Column order is incorrect")

    # 2. Verify NDT Entries
    entry_keys = list(app.ndt_entries.keys())
    print(f"NDT Entry Keys: {entry_keys}")
    if entry_keys == expected_ndt:
        print("SUCCESS: NDT Entry fields are correct.")
    else:
        print("FAILED: NDT Entry fields mismatch.")

    # 3. Verify Backward Compatibility in View Update
    app.update_daily_usage_view()
    
    # Check the inserted item
    child_id = app.daily_usage_tree.get_children()[0]
    values = app.daily_usage_tree.item(child_id, 'values')
    
    # Map values to columns
    val_dict = dict(zip(columns, values))
    
    print(f"Row Values: {val_dict}")
    
    # Check if '형광침투제' (mapped from NDT_형광) is 10.0
    if val_dict.get('형광침투제') == '10.0':
        print("SUCCESS: Backward compatibility for '형광' -> '형광침투제' works.")
    else:
        print(f"FAILED: '형광침투제' value is {val_dict.get('형광침투제')}, expected 10.0")

    # Check if '형광자분' is 0.0 (default)
    if val_dict.get('형광자분') == '0.0':
        print("SUCCESS: '형광자분' default value is correct.")
    else:
        print(f"FAILED: '형광자분' value is {val_dict.get('형광자분')}, expected 0.0")
        
    root.destroy()

if __name__ == "__main__":
    verify_columns()
