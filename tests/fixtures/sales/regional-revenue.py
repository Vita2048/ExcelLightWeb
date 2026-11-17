import matplotlib.pyplot as plt
regional = xl("Charts!A3:D8", headers=True)
figure, axes = plt.subplots(figsize=(8, 4.8), dpi=120, facecolor="white")
axes.bar(regional["Region"], regional["Total Revenue"], color="#4F81BD")
axes.set(title="Revenue by Region", xlabel="Region", ylabel="Revenue ($)")
axes.set_ylim(bottom=0)
axes.set_axisbelow(True)
axes.grid(axis="y", color="#dddddd")
axes.title.set_fontweight("bold")
figure.tight_layout()
print("Regional revenue:", regional.to_json(orient="records"))
figure
