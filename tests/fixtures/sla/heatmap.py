import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import PercentFormatter
import numpy as np

data = xl("Analysis!B26:H31", headers=True)
values = data.iloc[:, 1:].to_numpy(dtype=float)
paper, ink, muted = "#FFFFFF", "#142D46", "#516578"
fig, ax = plt.subplots(figsize=(4.32, 2.58), dpi=180, facecolor=paper)
fig.subplots_adjust(left=.22, right=.84, bottom=.16, top=.71)
fig.text(.06, .90, "Where breaches concentrate", fontsize=13, weight="bold", color=ink)
fig.text(.06, .80, "First-response breach rate by team", fontsize=8.5, color=muted)

palette = LinearSegmentedColormap.from_list("risk", ["#E8F5F1", "#FFC984", "#A22B44"])
palette.set_bad("#E9EEF3")
image = ax.imshow(np.ma.masked_invalid(values), cmap=palette, vmin=0, vmax=1, aspect="auto")
ax.set_xticks(range(6), ["Jan", "Feb", "Mar", "Apr", "May", "Jun"])
ax.set_yticks(range(len(data)), data["Team"].tolist())
ax.set_xticks(np.arange(-.5, 6, 1), minor=True)
ax.set_yticks(np.arange(-.5, len(data), 1), minor=True)
ax.grid(which="minor", color=paper, linewidth=1.2)
for (y, x), value in np.ndenumerate(values):
    ax.text(x, y, "–" if np.isnan(value) else f"{value:.0%}", ha="center", va="center",
            color=paper if value >= .70 else ink, fontsize=7.5, weight="medium")
ax.tick_params(axis="both", which="both", length=0, labelsize=7.5, colors=muted, pad=5)
for spine in ax.spines.values():
    spine.set_visible(False)
scale = fig.colorbar(image, cax=fig.add_axes([.87, .16, .022, .55]), ticks=[0, .5, 1],
                     format=PercentFormatter(1, decimals=0))
scale.outline.set_visible(False)
scale.ax.tick_params(length=0, labelsize=7, colors=muted)
fig
