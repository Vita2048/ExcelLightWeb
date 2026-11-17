import pandas as pd
import matplotlib.pyplot as plt
sales = xl("SalesTable[#Data]", headers=True).dropna(how="all")
products = xl("ProductsTable[#Data]", headers=True).dropna(how="all")
customers = xl("CustomersTable[#Data]", headers=True).dropna(how="all")
for label, frame in [("Sales", sales), ("Products", products), ("Customers", customers)]:
    if frame.isna().any().any():
        raise ValueError("Incomplete " + label + " record; fill missing fields or clear the whole record")
for frame, key in [(sales, "OrderID"), (products, "ProductID"), (customers, "CustomerID")]:
    if frame[key].isna().any() or frame[key].duplicated().any():
        raise ValueError("Missing or duplicate " + key)
if not sales["ProductID"].isin(products["ProductID"]).all() or not sales["CustomerID"].isin(customers["CustomerID"]).all():
    raise ValueError("Sales contains an unknown product/customer key")
if not sales["Region"].isin(["North", "South", "East", "West", "Central"]).all():
    raise ValueError("Sales contains a missing or unknown region")
joined = sales.merge(products, on="ProductID", how="left", validate="many_to_one")
joined["Revenue"] = (joined["UnitPrice"] * joined["Quantity"]).round(2)
joined["Month"] = pd.to_datetime(joined["OrderDate"]).dt.to_period("M").astype(str)
monthly = joined.groupby("Month")["Revenue"].sum().sort_index()
print("Monthly revenue:", monthly.to_json())
labels = pd.to_datetime(monthly.index).strftime("%B %Y")
figure, axes = plt.subplots(figsize=(8, 4.8), dpi=120, facecolor="white")
axes.bar(labels, monthly.values, color="#4F81BD")
axes.set(title="Monthly Sales Revenue", xlabel="Month", ylabel="Revenue ($)")
axes.set_ylim(bottom=0)
axes.set_axisbelow(True)
axes.grid(axis="y", color="#dddddd")
axes.title.set_fontweight("bold")
figure.tight_layout()
figure
