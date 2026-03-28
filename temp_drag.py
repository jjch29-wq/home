import tkinter as tk
from tkinter import ttk

def run():
    r = tk.Tk()
    
    # Simulate UI
    frame = ttk.Frame(r)
    frame.pack(fill='both', expand=True, padx=20, pady=20)
    
    e1 = ttk.Entry(frame, width=30)
    e1.pack(pady=10)
    e1.insert(0, 'Drag me to select text')
    
    e2 = ttk.Entry(frame, width=30)
    e2.pack(pady=10)
    e2.insert(0, 'Target')

    def _on_click_drop_focus(event):
        try:
            widget = event.widget
            if widget and hasattr(widget, 'winfo_class'):
                if widget.winfo_class() not in ('Treeview', 'Text', 'Listbox', 'Scrollbar', 'TScrollbar', 'Entry', 'TEntry'):
                    r.focus_set()
        except:
            pass
            
    r.bind_all("<Button-1>", _on_click_drop_focus, add='+')
    
    # simulate the popup
    popup = tk.Menu(r, tearoff=0)
    popup.add_command(label="Copy", command=lambda: print("Copy"))
    
    def _show(event):
        event.widget.focus_set()
        popup.tk_popup(event.x_root, event.y_root)
        
    r.bind_class("TEntry", "<Button-3>", _show)
    
    r.after(3000, lambda: print("App started, test drag..."))
    r.after(10000, r.destroy)
    r.mainloop()

run()
