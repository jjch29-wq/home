import sys
import os
import json
import tkinter as tk
from unittest.mock import MagicMock

# Import app
sys.path.append(r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src")
import importlib
app_module = importlib.import_module("Archived-Main-App-20260405-RT-Fix")

# Mock UI components
root = tk.Tk()
app = app_module.PMIReportApp(root)

# Setup test data
app.kogas_extracted_data = [
    {"No": "1", "Joint": "J1", "Date": "2023-10-01", "Dwg": "DWG-01", "Mat": "SS304", "Welder": "W1", "Thk": "10", "Size": "2", "Result": "ACC", "Acc": "O", "selected": True, "date_filtered": True}
]
app.config['KOGAS_START_ROW'] = 14
app.config['KOGAS_DATA_END_ROW'] = 25

template_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\가스공사 의뢰서.xlsx"

# Mock messagebox
app_module.messagebox = MagicMock()

# Run generation
try:
    print("Running _run_rt_process for KOGAS...")
    app._run_rt_process(app.kogas_extracted_data, template_path, mode="KOGAS")
    print("Success. Messagebox calls:", app_module.messagebox.mock_calls)
except Exception as e:
    import traceback
    traceback.print_exc()

root.destroy()
