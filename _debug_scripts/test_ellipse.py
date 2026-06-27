import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig, ax = plt.subplots()
ax.set_xlim(-20, 20)
ax.set_ylim(20, -5) # inverted Y, like side view

# defect_width = 2
e1 = patches.Ellipse((0, 5), width=5, height=2, angle=37.5, color='red', alpha=0.5)
ax.add_patch(e1)

# defect_width = 10
e2 = patches.Ellipse((10, 5), width=5, height=10, angle=37.5, color='blue', alpha=0.5)
ax.add_patch(e2)

plt.savefig("test_ellipse.png")
