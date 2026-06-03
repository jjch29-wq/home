import sys
import matplotlib
matplotlib.use('Agg')
sys.path.append('.')
import importlib
paut = importlib.import_module("PAUT 개선각 위치계산")

app = paut.App()
print("INITIAL:", app.defects[0]["width"])

# Set selected idx
app.selected_defect_idx = 0

# Set shape to Ellipse
app.shape_var.set("타원형(Ellipse)")

# Change width entry
app.entries["defect_width"].delete(0, "end")
app.entries["defect_width"].insert(0, "15.0")

# Apply
app.apply_defect_properties()

print("AFTER APPLY:", app.defects[0]["width"])
print("Shape:", app.defects[0]["shape"])

# Find Ellipse patches in Side View
for p in app.ax_side.patches:
    if isinstance(p, matplotlib.patches.Ellipse):
        print("Side View Ellipse:", p.width, p.height)
