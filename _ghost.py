
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

# ghost decay lines for the cover: lines only, no area fill -> no visible seam
m = np.arange(0, 12)
series = [( [100, 61, 48, 41, 36, 33, 30, 28, 26, 25, 24, 23], 0.62, 2.2),
          ( [100, 78, 70, 65, 62, 60, 58, 57, 56, 55, 54, 53], 0.30, 1.6) ]
fig, ax = plt.subplots(figsize=(7.6, 2.4), dpi=220)
fig.patch.set_alpha(0)
ax.set_facecolor("none")
for y, a, lw in series:
    ax.plot(m, y, color="#3FBAC3", lw=lw, alpha=a, solid_capstyle="round")
ax.set_xlim(-0.2, 11.2)
ax.set_ylim(6, 106)
for sp in ax.spines.values():
    sp.set_visible(False)
ax.set_xticks([]); ax.set_yticks([])
fig.subplots_adjust(0, 0, 1, 1)
fig.savefig(r"E:\Cari Work\Alora\Visualizations\cover_ghost.png", transparent=True, dpi=220)
print("ok")
