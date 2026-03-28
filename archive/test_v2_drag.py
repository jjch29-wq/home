import tkinter as tk
import importlib.util
import os

spec = importlib.util.spec_from_file_location("unified", "JJCHSITPMI-V2-Unified.py")
unified = importlib.util.module_from_spec(spec)
spec.loader.exec_module(unified)

root = tk.Tk()
app = unified.PMIReportApp(root)

def do_test():
    def find_entry(w):
        if w.winfo_class() in ('Entry', 'TEntry'):
            return w
        for c in w.winfo_children():
            res = find_entry(c)
            if res: return res
        return None
    
    entry = find_entry(root)
    if not entry:
        print("FAIL: No entry found")
        root.destroy()
        return

    print(f"Testing drag on {entry.winfo_class()} widget: {entry}")
    entry.focus_force()
    entry.delete(0, tk.END)
    entry.insert(0, "ABCDEFGH")
    entry.update()
    
    # Simulate realistic click and drag
    entry.event_generate("<Enter>", x=5, y=5)
    entry.event_generate("<Button-1>", x=5, y=5)
    for i in range(5, 50, 5):
        entry.event_generate("<B1-Motion>", x=i, y=5)
    entry.event_generate("<ButtonRelease-1>", x=50, y=5)
    
    has_sel = entry.select_present()
    print("Has selection after B1-Motion?:", has_sel)
    if has_sel:
        print("Selected:", entry.get()[entry.index("sel.first"):entry.index("sel.last")])
        
    root.destroy()

root.after(1500, do_test)
root.mainloop()
