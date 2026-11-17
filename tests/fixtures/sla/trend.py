import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator, PercentFormatter
import numpy as np

data = xl("Analysis!B16:E22", headers=True)
x = np.arange(len(data))
labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
paper, ink, muted, teal = "#FFFFFF", "#142D46", "#516578", "#087F81"
fig, ax = plt.subplots(figsize=(4.32, 2.58), dpi=180, facecolor=paper)
fig.subplots_adjust(left=.13, right=.87, bottom=.18, top=.72)
fig.text(.06, .90, "Volume & response breaches", fontsize=13, weight="bold", color=ink)
fig.text(.06, .80, "Opened tickets", fontsize=8.5, color="#3E6284")
fig.text(.39, .80, "●  Breach rate", fontsize=8.5, color=teal)

ax.bar(x, data["Tickets"], width=.59, color="#C8DAEC", zorder=2)
ax.set_xticks(x, labels)
ax.set_xlim(-.56, len(labels) - .44)
ax.set_ylim(0, max(1, data["Tickets"].max() * 1.18))
ax.yaxis.set_major_locator(MaxNLocator(4, integer=True))
ax.grid(axis="y", color="#E2EBF2", linewidth=.75)
rate = ax.twinx()
rate.plot(x, data["Breach rate"].astype(float), color=teal, linewidth=2, marker="o", markersize=4)
rate.set_ylim(0, 1)
rate.set_yticks([0, .25, .5, .75, 1])
rate.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
for axis in (ax, rate):
    axis.tick_params(axis="both", length=0, labelsize=8, colors=muted, pad=5)
    for spine in axis.spines.values():
        spine.set_visible(False)
fig
