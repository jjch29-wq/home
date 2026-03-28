
import tkinter as tk
from tkinter import ttk

class MockMaterialManagerLock:
    def __init__(self):
        self.layout_locked = False
        self.root = tk.Tk()
        self.widget = ttk.Frame(self.root)
        
    def on_drag_start(self, event):
        if self.layout_locked:
            return "break"
        return None

    def on_resize_start(self, event):
        if self.layout_locked:
            return "break"
        return None

    def on_mouse_motion(self, event):
        if self.layout_locked:
            return "break"
        return None

app = MockMaterialManagerLock()
event = tk.Event()

# Test Unlocked
app.layout_locked = False
assert app.on_drag_start(event) is None, "Unlocked drag should not return break"
assert app.on_resize_start(event) is None, "Unlocked resize should not return break"
assert app.on_mouse_motion(event) is None, "Unlocked motion should not return break"

# Test Locked
app.layout_locked = True
assert app.on_drag_start(event) == "break", "Locked drag should return break"
assert app.on_resize_start(event) == "break", "Locked resize should return break"
assert app.on_mouse_motion(event) == "break", "Locked motion should return break"

print("Layout Lock Logic Verification SUCCESS")
app.root.destroy()
