
import tkinter as tk
from tkinter import ttk
import sys
import os

# Mock the parts of MaterialManager with the NEW implementation
class MockMaterialManagerNEW:
    def __init__(self):
        self.users = ["User 1", "User 2"]
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
        
        # NEW LOGIC
        self.draggable_items[config_key] = container
        container._label_widget = lbl
        container._widget = widget
        container._widget_class = widget_class
        container._widget_kwargs = widget_kwargs
        container._manage_list_key = manage_list_key
        
        return container, widget

    def clone_widget(self, key):
        orig = self.draggable_items.get(key)
        new_key = f"clone_test"
        label_text = orig._label_widget.cget('text')
        cont, w = self.create_draggable_container(
            self.entry_inner_frame, 
            label_text, 
            orig._widget_class, 
            new_key, 
            manage_list_key=getattr(orig, '_manage_list_key', None), # Pass manage_list_key
            **orig._widget_kwargs
        )
        print(f"Cloned widget values during creation: {w.cget('values')}")
        return w

    def refresh_ui_for_list_change(self, config_key):
        # NEW implementation in MaterialManager-6.py
        list_map = {
            'users': self.users
        }
        if config_key not in list_map: return
        current_vals = list_map[config_key]
        sorted_vals = sorted(current_vals)
        
        if config_key == 'users':
            if hasattr(self, 'cb_daily_user'): 
                self.cb_daily_user['values'] = sorted_vals
                print(f"Main widget updated to: {self.cb_daily_user['values']}")

        # 2. Update ALL draggable widgets (clones) that depend on this list
        for key, container in self.draggable_items.items():
            if hasattr(container, '_manage_list_key') and container._manage_list_key == config_key:
                if hasattr(container, '_widget'):
                    container._widget['values'] = sorted_vals
                    print(f"Widget {key} updated to: {container._widget['values']}")

app = MockMaterialManagerNEW()
# Original widget
container, app.cb_daily_user = app.create_draggable_container(
    app.entry_inner_frame, "Worker:", ttk.Combobox, 'user_box', manage_list_key='users', values=app.users
)

# Clone the widget
cloned_cb = app.clone_widget('user_box')

# Change user list
app.users = ["User 1", "User 2", "User 3"]
app.refresh_ui_for_list_change('users')

# Check if clone updated
print(f"Cloned widget values after update: {cloned_cb.cget('values')}")

if str(cloned_cb.cget('values')) == str(tuple(sorted(app.users))):
    print("Verification SUCCESS: cloned widget updated correctly")
else:
    print("Verification FAILED: cloned widget did NOT update")

app.root.destroy()
