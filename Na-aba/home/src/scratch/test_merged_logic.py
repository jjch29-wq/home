import sys
import os
import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox
import importlib.util

# Monkeypatch messagebox to prevent blocking UI dialogs
messagebox.showinfo = lambda title, msg: print(f"[POPUP INFO] {title}: {msg}")
messagebox.showerror = lambda title, msg: print(f"[POPUP ERROR] {title}: {msg}")
messagebox.showwarning = lambda title, msg: print(f"[POPUP WARNING] {title}: {msg}")

# Load the module dynamically
module_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\요청서 합치기.py"
spec = importlib.util.spec_from_file_location("excel_merger_module", module_path)
merger_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(merger_module)

class SilentExcelMergerApp:
    def __init__(self):
        self.selected_folder = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba"
        self.keyword_var = tk.StringVar(value="No, Joint, Dwg, Size, THK, Result, Date, Report No, Identification No")
        self.excel_files = []
        self.btn_merge = {"state": tk.NORMAL} # Mock widget
        self.status_var = tk.StringVar(value="Ready")
        
    def add_log(self, msg):
        # Safely print to cp949 terminal
        try:
            print(f"[LOG] {msg}")
        except Exception:
            clean_msg = msg.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            clean_msg = re.sub(r'[^\u0000-\u007F\uAC00-\uD7A3]', '', clean_msg)
            try:
                print(f"[LOG] {clean_msg}")
            except Exception:
                try:
                    print(f"[LOG] {clean_msg.encode('ascii', errors='ignore').decode('ascii')}")
                except Exception:
                    pass
        
    def normalize(self, text):
        if pd.isna(text): return ""
        t = str(text).lower()
        t = re.sub(r'[^a-z0-9가-힣]', '', t)
        return t.strip()
        
    def scan_files(self):
        self.excel_files = [f for f in os.listdir(self.selected_folder) 
                            if (f.endswith('.xlsx') or f.endswith('.xlsm')) and not f.startswith('~$') and "Smart_Merged" not in f]
        print(f"Scanned files count: {len(self.excel_files)}")

# Attach the merge_logic
SilentExcelMergerApp.merge_logic = merger_module.ExcelMergerApp.merge_logic

# Create headless TK root
root = tk.Tk()
root.withdraw()

app = SilentExcelMergerApp()
app.scan_files()
print("Starting headless merge logic...\n")
app.merge_logic()
print("\nMerge logic verification complete.")
