import sys
import os
import importlib.util
import tkinter as tk

script_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\Archived-Main-App-20260405-RT-Fix.py'

spec = importlib.util.spec_from_file_location("main_app", script_path)
main_app = importlib.util.module_from_spec(spec)
sys.modules["main_app"] = main_app
spec.loader.exec_module(main_app)

root = tk.Tk()

# Initialize app
app = main_app.PMIReportApp(root)

# Force select RT tab & KOGAS preview
app.mode_notebook.select(1)
app.rt_preview_nb.select(1)
root.update()

# Get columns from Treeview
cols = app.kogas_preview_tree["columns"]
print("Columns in kogas_preview_tree:")
for col in cols:
    heading_text = app.kogas_preview_tree.heading(col, "text")
    width = app.kogas_preview_tree.column(col, "width")
    print(f"  Column ID: {col:<10} | Heading: {heading_text:<15} | Width: {width}")

root.destroy()
