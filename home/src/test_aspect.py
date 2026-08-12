import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import math

fig, (ax_side, ax_top) = plt.subplots(2, 1, figsize=(8, 8))

half_width = 305
specimen_length = 320

# Simulate ax_side
ax_side.set_xlim(-half_width, half_width)
ax_side.set_xscale('symlog', linthresh=30, linscale=0.56)
ax_side.xaxis.set_major_formatter(ScalarFormatter())
ax_side.set_xticks([-half_width, -100, -30, 0, 30, 100, half_width])
ax_side.set_ylim(30.88, -10)
ax_side.axvspan(12.5, half_width, color='pink', alpha=0.3)
ax_side.axvspan(-half_width, -12.5, color='pink', alpha=0.3)
ax_side.set_aspect('equal', adjustable='datalim') # Try with True

# Simulate ax_top
ax_top.set_xlim(-half_width, half_width)
ax_top.set_xscale('symlog', linthresh=30, linscale=0.56)
ax_top.xaxis.set_major_formatter(ScalarFormatter())
ax_top.set_xticks([-half_width, -100, -30, 0, 30, 100, half_width])
ax_top.set_ylim(0, specimen_length)
ax_top.axvspan(12.5, half_width, color='pink', alpha=0.3)
ax_top.axvspan(-half_width, -12.5, color='pink', alpha=0.3)
ax_top.set_aspect('equal', adjustable='datalim') # Try with True

plt.tight_layout()
fig.savefig('test_aspect.png')
