import tkinter as tk
import importlib.util
import os

spec = importlib.util.spec_from_file_location("unified", "JJCHSITPMI-V2-Unified.py")
unified = importlib.util.module_from_spec(spec)
spec.loader.exec_module(unified)

root = tk.Tk()
app = unified.PMIReportApp(root)

def find_entry(w):
    if w.winfo_class() in ('Entry', 'TEntry'):
        return w
    for c in w.winfo_children():
        res = find_entry(c)
        if res: return res
    return None

entry = find_entry(root)
print("Found Entry:", entry)
tags = entry.bindtags()
print("Bindtags:", tags)

for tag in tags:
    print(f"\nBindings for tag: {tag}")
    try:
        events = entry.bind_class(tag)
        for e in events:
            print(f"  {e} -> {entry.bind_class(tag, e)}")
    except Exception as ex:
        print(f"  Error getting bindings: {ex}")

root.destroy()
