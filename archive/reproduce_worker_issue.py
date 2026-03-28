
import tkinter as tk
from tkinter import ttk
import sys
import os

# Mock the parts of MaterialManager we need
class MockMaterialManager:
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
        self.draggable_items[config_key] = container
        container._label_widget = lbl
        container._widget_class = widget_class
        container._widget_kwargs = widget_kwargs
        # Missing in current code: container._manage_list_key = manage_list_key
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
            **orig._widget_kwargs
        )
        print(f"Cloned widget values: {w.cget('values')}")
        return w

    def refresh_ui_for_list_change(self, config_key):
        # Current implementation in MaterialManager-6.py
        if config_key == 'users':
            if hasattr(self, 'cb_daily_user'): 
                self.cb_daily_user['values'] = sorted(self.users)
                print(f"Main widget updated to: {self.cb_daily_user['values']}")

app = MockMaterialManager()
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

if str(cloned_cb.cget('values')) != str(tuple(sorted(app.users))):
    print("reproduction successful: cloned widget did NOT update")
else:
    print("reproduction failed: cloned widget updated")

app.root.destroy()
