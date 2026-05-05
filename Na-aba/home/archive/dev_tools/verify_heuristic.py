
import tkinter as tk
from tkinter import ttk
import sys
import os

# Mock the parts of MaterialManager with the HEURISTIC implementation
class MockMaterialManagerHeuristic:
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
        container._widget = widget
        container._manage_list_key = manage_list_key
        return container, widget

    def refresh_ui_for_list_change(self, config_key):
        list_map = {'users': self.users}
        if config_key not in list_map: return
        sorted_vals = sorted(list_map[config_key])
        
        # 2. Update ALL draggable widgets (clones) that depend on this list
        for key, container in self.draggable_items.items():
            # Apply HEURISTIC
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
                    print(f"Widget {key} (Label: {container._label_widget.cget('text')}) updated to: {container._widget['values']}")

app = MockMaterialManagerHeuristic()
# Original widget
app.create_draggable_container(app.entry_inner_frame, "담당자:", ttk.Combobox, 'user_box', manage_list_key='users', values=app.users)

# Create a "Corrupted" clone (missing manage_list_key but has label)
# This simulates what's in the user's config
container, cb = app.create_draggable_container(app.entry_inner_frame, "작업자1", ttk.Combobox, 'clone_broken', manage_list_key=None, values=[])

print(f"Broken clone values before update: {cb.cget('values')}")

# Change user list
app.users = ["User 1", "User 2", "User 3"]
app.refresh_ui_for_list_change('users')

# Check if clone updated via heuristic
print(f"Broken clone values after update: {cb.cget('values')}")

if str(cb.cget('values')) == str(tuple(sorted(app.users))):
    print("Heuristic verification SUCCESS: cloned widget updated via label")
else:
    print("Heuristic verification FAILED")

app.root.destroy()
