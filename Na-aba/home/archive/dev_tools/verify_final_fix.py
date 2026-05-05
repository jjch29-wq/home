
import tkinter as tk
from tkinter import ttk
import sys
import os

# Mock the parts of MaterialManager with the FINAL implementation
class MockMaterialManagerFinal:
    def __init__(self):
        self.users = ["User 1", "User 2"]
        self.sites = ["Site A"]
        self.draggable_items = {}
        self.root = tk.Tk()
        self.entry_inner_frame = ttk.Frame(self.root)
        self.entry_inner_frame.pack()
        
    def create_draggable_container(self, parent, label_text, widget_class, config_key, manage_list_key=None, **widget_kwargs):
        container = ttk.Frame(parent)
        lbl = ttk.Label(container, text=label_text)
        lbl.pack(side='left')
        widget = widget_class(container, **widget_kwargs)
        widget.pack(side='left')
        
        self.draggable_items[config_key] = container
        container._label_widget = lbl
        container._widget = widget
        container._manage_list_key = manage_list_key
        return container, widget

    def refresh_ui_for_list_change(self, config_key):
        list_map = {'users': self.users, 'sites': self.sites}
        if config_key not in list_map: return
        sorted_vals = sorted(list_map[config_key])
        
        for key, container in self.draggable_items.items():
            m_key = getattr(container, '_manage_list_key', None)
            if not m_key:
                if hasattr(container, '_label_widget'):
                    lbl_text = container._label_widget.cget('text').lower()
                    if config_key == 'users' and any(x in lbl_text for x in ['작업자', '담당자', 'user', 'worker']):
                        m_key = 'users'
                        container._manage_list_key = 'users'

            if m_key == config_key:
                if hasattr(container, '_widget'):
                    container._widget['values'] = sorted_vals
                    # print(f"Widget {key} updated")

    def load_tab_config_mock(self):
        # 1. Restore data
        self.users = ["Worker A", "Worker B"] # New data from config/db
        
        # 2. Restore standard widgets (Mocking setup_daily_usage_tab call)
        self.cb_daily_user = ttk.Combobox(self.root) # This represents the internal ref
        
        # 3. Restore Draggable Items
        # Imagine we restore a clone from config that has NO manage_list_key
        cont, w = self.create_draggable_container(self.entry_inner_frame, "작업자1", ttk.Combobox, 'clone_123', manage_list_key=None)
        
        # 4. THE FINAL FIX: Broad refresh
        for l_key in ['users', 'sites']:
            self.refresh_ui_for_list_change(l_key)
            
        print(f"Clone values after load: {w.cget('values')}")
        return w

app = MockMaterialManagerFinal()
cloned_cb = app.load_tab_config_mock()

if str(cloned_cb.cget('values')) == str(tuple(sorted(["Worker A", "Worker B"]))):
    print("Final Verification SUCCESS: cloned widget populated after load")
else:
    print("Final Verification FAILED")

app.root.destroy()
