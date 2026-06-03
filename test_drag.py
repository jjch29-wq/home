import sys
import matplotlib
matplotlib.use('Agg')
sys.path.append('.')
import importlib
paut = importlib.import_module("PAUT 개선각 위치계산")

app = paut.App()
# Simulate a drag down to z=16
app.selected_defect_idx = 0
app.orig_z_start = 5.0
app.orig_z_end = 9.0
app.orig_cz = 7.0
app.mouse_offset_z = 0.0
app.mouse_offset_x = 0.0
app.active_ax = app.ax_side
app.drag_mode = "translate"
app.dragging = True

class MockEvent:
    def __init__(self, x, y, inaxes):
        self.xdata = x
        self.ydata = y
        self.inaxes = inaxes

print("BEFORE DRAG:")
print(app.defects[0])
print("Side patches:", len(app.ax_side.patches))

app.on_motion(MockEvent(5.0, 16.0, app.ax_side))

print("AFTER DRAG:")
print(app.defects[0])
print("Side patches:", len(app.ax_side.patches))
for p in app.ax_side.patches:
    print(p)
