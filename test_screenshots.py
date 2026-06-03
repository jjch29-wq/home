import sys
import matplotlib
matplotlib.use('Agg')
sys.path.append('.')
import importlib
paut = importlib.import_module("PAUT 개선각 위치계산")

app = paut.App()
app.selected_defect_idx = 0
app.shape_var.set("타원형(Ellipse)")

# Before width change
app.entries["defect_width"].delete(0, "end")
app.entries["defect_width"].insert(0, "2.0")
app.apply_defect_properties()
app.fig.savefig("test_before.png")

# After width change
app.entries["defect_width"].delete(0, "end")
app.entries["defect_width"].insert(0, "15.0")
app.apply_defect_properties()
app.fig.savefig("test_after.png")

print("Generated test_before.png and test_after.png")
