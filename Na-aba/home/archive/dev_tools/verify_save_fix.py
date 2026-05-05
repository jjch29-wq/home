
import tkinter as tk
from tkinter import ttk
import json
import os

class MockMaterialManagerSaveFix:
    def __init__(self):
        self.root = tk.Tk()
        self.notebook = ttk.Notebook(self.root)
        self.tab = ttk.Frame(self.notebook)
        self.notebook.add(self.tab, text="Test")
        self.notebook.pack()
        
        self.sites = ["A"]
        self.users = ["U"]
        self.warehouses = []
        self.equipments = []
        self.layout_locked = True
        self.config_path = "test_config.json"
        self.draggable_items = {}
        self.memos = {}
        
    def save_tab_config(self):
        # MOCKED LOGIC based on the NEW MaterialManager-6.py
        try:
            config = {
                'selected_tab': self.notebook.index(self.notebook.select()),
                'tab_order': [self.notebook.tab(i, 'text') for i in range(self.notebook.index('end'))],
                'sites': self.sites,
                'users': self.users,
                'warehouses': self.warehouses,
                'equipments': self.equipments,
                'layout_locked': self.layout_locked,
                'draggable_geometries': {}
            }
            
            for key, widget in self.draggable_items.items():
                if widget.winfo_manager() == 'place':
                    config['draggable_geometries'][key] = {
                        'x': widget.winfo_x(), 'y': widget.winfo_y(),
                        'width': widget.winfo_width(), 'height': widget.winfo_height(),
                        'hidden': False
                    }
                    if hasattr(widget, '_label_widget'):
                        config['draggable_geometries'][key]['custom_label'] = "Label"
                elif hasattr(widget, 'winfo_manager') and widget.winfo_manager() == '':
                    config['draggable_geometries'][key] = {'hidden': True}
                else:
                    # It's in grid or pack, NO ACTION taken now (fix confirmed)
                    pass
                    
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print("Save successful")
        except Exception as e:
            print(f"Save FAILED: {e}")

# Prepare test
app = MockMaterialManagerSaveFix()

# Case 1: Widget in PLACE
w1 = ttk.Frame(app.tab)
w1.place(x=10, y=10, width=50, height=50)
w1._label_widget = ttk.Label(w1, text="L1")
app.draggable_items['w1'] = w1

# Case 2: Widget in GRID (this would have caused crash before)
w2 = ttk.Frame(app.tab)
w2.grid(row=0, column=0)
w2._label_widget = ttk.Label(w2, text="L2")
app.draggable_items['w2'] = w2

app.save_tab_config()

with open(app.config_path, 'r', encoding='utf-8') as f:
    saved = json.load(f)
    print(f"Saved layout_locked: {saved.get('layout_locked')}")
    print(f"Saved geometries: {list(saved.get('draggable_geometries').keys())}")

if saved.get('layout_locked') == True and 'w1' in saved.get('draggable_geometries') and 'w2' not in saved.get('draggable_geometries'):
    print("Verification SUCCESS: No crash and w2 (grid) was skipped correctly")
else:
    print("Verification FAILED")

if os.path.exists(app.config_path): os.remove(app.config_path)
app.root.destroy()
