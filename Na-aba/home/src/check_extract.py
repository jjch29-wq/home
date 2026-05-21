import tkinter as tk
import sys
import os

sys.path.append(os.path.abspath('Na-aba/home/src'))

import importlib
main_module = importlib.import_module("Archived-Main-App-20260405-RT-Fix")

root = tk.Tk()
root.withdraw()

app = main_module.PMIReportApp(root)
file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\data\가스공사 의뢰서.xlsx'

# Mock tab selection to return KOGAS mode
# Mode logic:
# main_tab = self.mode_notebook.tab(self.mode_notebook.select(), 'text')
# if 'RT' in main_tab:
#     sub_tab_text = self.rt_preview_nb.tab(self.rt_preview_nb.select(), 'text')
#     mode = 'KOGAS' if '가스공사' in sub_tab_text else 'RT'
# ...

app.mode_notebook.tab = lambda tab_id, option: "RT" if option == "text" else None
app.rt_preview_nb.tab = lambda tab_id, option: "가스공사" if option == "text" else None

app.kogas_target_file_path.set(file_path)

# Run extract_only
app.extract_only(show_msg=False)

# Check extracted data
data = app.kogas_extracted_data
print(f"Extracted {len(data)} items into kogas_extracted_data.")
if data:
    print("First 3 items:")
    for i, item in enumerate(data[:3]):
        clean_item = {k: v for k, v in item.items() if k != '_src'}
        print(f"Item {i}: {clean_item}")

root.destroy()
