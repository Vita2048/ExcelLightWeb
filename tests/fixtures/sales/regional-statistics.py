import pandas as pd
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
order = ["North", "South", "East", "West", "Central"]
regional = joined.groupby("Region").agg(
    **{"Total Orders": ("OrderID", "nunique"), "Units Sold": ("Quantity", "sum"),
       "Total Revenue": ("Revenue", "sum")}).reindex(order, fill_value=0).reset_index()
total = pd.DataFrame([{"Region": "Total", "Total Orders": sales["OrderID"].nunique(),
                       "Units Sold": regional["Units Sold"].sum(),
                       "Total Revenue": regional["Total Revenue"].sum().round(2)}])
pd.concat([regional, total], ignore_index=True)
