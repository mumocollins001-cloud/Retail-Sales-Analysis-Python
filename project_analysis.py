import pandas as pd

# Load dataset
df = pd.read_csv(
    r"C:\Users\ADMIN\Desktop\Retail_Sales_Project\data\superstore.csv",
    encoding="latin1"
)

# DATA CLEANING
df["Order Date"] = pd.to_datetime(df["Order Date"])
df["Ship Date"] = pd.to_datetime(df["Ship Date"])
df["Postal Code"] = df["Postal Code"].astype(str)

# EXPLORATORY DATA ANALYSIS (EDA)
print("Total Sales:")
print(df["Sales"].sum())

print("\nTotal Profit:")
print(df["Profit"].sum())

print("\nSales by Category:")
print(df.groupby("Category")["Sales"].sum())

print("\nProfit by Category:")
print(df.groupby("Category")["Profit"].sum())
category_summary = df.groupby("Category")[["Sales", "Profit"]].sum()

category_summary["Profit Margin %"] = (
    category_summary["Profit"] / category_summary["Sales"]
) * 100

print(category_summary)
subcategory_profit = df.groupby("Sub-Category")["Profit"].sum()

print("\nProfit by Sub-Category:")
print(subcategory_profit.sort_values(ascending=False))
print("\nAverage Discount by Sub-Category:")
print(
    df.groupby("Sub-Category")["Discount"]
      .mean()
      .sort_values(ascending=False)
)
# Create Year-Month column
df["YearMonth"] = df["Order Date"].dt.to_period("M")

monthly_sales = df.groupby("YearMonth")["Sales"].sum()

print("\nMonthly Sales:")
print(monthly_sales)
import matplotlib.pyplot as plt
region_summary = df.groupby("Region")[["Sales", "Profit"]].sum()

region_summary["Profit Margin %"] = (
    region_summary["Profit"] / region_summary["Sales"]
) * 100

print(region_summary)
top_products = (
    df.groupby("Product Name")["Sales"]
      .sum()
      .sort_values(ascending=False)
      .head(10)
)

print("\nTop 10 Products by Sales:")
print(top_products)
top_profit_products = (
    df.groupby("Product Name")["Profit"]
      .sum()
      .sort_values(ascending=False)
      .head(10)
)

print("\nTop 10 Products by Profit:")
print(top_profit_products)
top_customers = (
    df.groupby("Customer Name")["Sales"]
      .sum()
      .sort_values(ascending=False)
      .head(10)
)

print("\nTop 10 Customers by Sales:")
print(top_customers)
top_profit_customers = (
    df.groupby("Customer Name")["Profit"]
      .sum()
      .sort_values(ascending=False)
      .head(10)
)

print("\nTop 10 Customers by Profit:")
print(top_profit_customers)

# ===================================
# EXPORT RESULTS
# ===================================

import os

print(os.getcwd())
top_profit_customers.to_excel(
    "top_profit_customers.xlsx"
)

print("\nExcel file exported successfully.")

# ===================================
# VISUALIZATIONS
# ===================================

import matplotlib.pyplot as plt

top_profit_customers.plot(kind="bar")

plt.title("Top 10 Customers by Profit")
plt.ylabel("Profit")
plt.tight_layout()

plt.savefig("top_profit_customers.png")

plt.show()
monthly_sales.plot(figsize=(12,5))

plt.title("Monthly Sales Trend")
plt.ylabel("Sales")
plt.tight_layout()

plt.savefig("monthly_sales_trend.png")

plt.show()
state_sales = (
    df.groupby("State")["Sales"]
      .sum()
      .sort_values(ascending=False)
      .head(10)
)

print("\nTop 10 States by Sales:")
print(state_sales)
state_sales.plot(
    kind="barh",
    figsize=(10,6)
)

plt.title("Top 10 States by Sales")
plt.xlabel("Sales")
plt.tight_layout()

plt.savefig("top_states_sales.png")

plt.show()
df["YearMonth"] = df["YearMonth"].astype(str)

# ===================================
# SQL ANALYSIS
# ===================================

from pandasql import sqldf

query = """
SELECT
    Category,
    SUM(Sales) AS Total_Sales
FROM df
GROUP BY Category
ORDER BY Total_Sales DESC
"""
print(df["YearMonth"].dtype)
print(df["YearMonth"].head())

result = sqldf(query)

print(result)
print(df["YearMonth"].dtype)
query = """
SELECT
    Region,
    ROUND(SUM(Sales), 2) AS Total_Sales,
    ROUND(SUM(Profit), 2) AS Total_Profit
FROM df
GROUP BY Region
ORDER BY Total_Profit DESC
"""

result = sqldf(query)

print(result)
query = """
SELECT
    Region,
    ROUND(AVG(Discount),3) AS Avg_Discount
FROM df
GROUP BY Region
ORDER BY Avg_Discount DESC
"""

result = sqldf(query)

print(result)
query = """
SELECT
    Region,
    ROUND(
        (SUM(Profit) / SUM(Sales)) * 100,
        2
    ) AS Profit_Margin
FROM df
GROUP BY Region
ORDER BY Profit_Margin DESC
"""

region_profit = sqldf(query)

print(region_profit)
query = """
SELECT
    Region,
    Category,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
GROUP BY Region, Category
ORDER BY Region, Total_Profit
"""

result = sqldf(query)

print(result)
query = """
SELECT
    Sub_Category,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM (
    SELECT
        Region,
        REPLACE([Sub-Category], '-', '_') AS Sub_Category,
        Sales,
        Profit
    FROM df
)
WHERE Region = 'Central'
AND Sub_Category IN ('Tables','Bookcases')
GROUP BY Sub_Category
ORDER BY Total_Profit
"""

result = sqldf(query)

print(result)

print(result)
print(result.columns)
category_query = """
SELECT
    Category,
    SUM(Sales) AS Total_Sales
FROM df
GROUP BY Category
ORDER BY Total_Sales DESC
"""

category_sales = sqldf(category_query)

print(category_sales)
category_sales.plot(
    x="Category",
    y="Total_Sales",
    kind="bar"
)

plt.title("Sales by Category")
plt.ylabel("Sales")
plt.tight_layout()

plt.savefig("sales_by_category.png")
plt.show()
category_sales = sqldf(category_query)

region_profit.plot(
    x="Region",
    y="Profit_Margin",
    kind="bar"
)

plt.title("Profit Margin by Region")
plt.ylabel("Profit Margin (%)")
plt.tight_layout()
plt.savefig("profit_margin_region.png")
plt.show()
loss_query = """
SELECT
    [Sub-Category] AS Sub_Category,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
GROUP BY [Sub-Category]
HAVING SUM(Profit) < 0
ORDER BY Total_Profit
"""

loss_result = sqldf(loss_query)

print(loss_result)
loss_result.plot(
    x="Sub_Category",
    y="Total_Profit",
    kind="bar"
)

plt.title("Loss-Making Sub-Categories")
plt.ylabel("Profit")
plt.tight_layout()
plt.savefig("loss_making_subcategories.png")
plt.show()
# ===================================
# REPORT VALIDATION CHECKS
# ===================================
print("\n=== OVERALL PERFORMANCE ===")

print("Total Sales:")
print(round(df["Sales"].sum(), 2))

print("\nTotal Profit:")
print(round(df["Profit"].sum(), 2))
category_sales_query = """
SELECT
    Category,
    ROUND(SUM(Sales),2) AS Total_Sales
FROM df
GROUP BY Category
ORDER BY Total_Sales DESC
"""

category_sales = sqldf(category_sales_query)

print("\n=== SALES BY CATEGORY ===")
print(category_sales)
monthly_sales = (
    df.groupby("YearMonth")["Sales"]
      .sum()
)

print("\n=== FIRST 5 MONTHS ===")
print(monthly_sales.head())

print("\n=== LAST 5 MONTHS ===")
print(monthly_sales.tail())
region_query = """
SELECT
    Region,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
GROUP BY Region
ORDER BY Total_Profit DESC
"""

region_summary = sqldf(region_query)

print("\n=== REGION PERFORMANCE ===")
print(region_summary)
profit_margin_query = """
SELECT
    Region,
    ROUND(
        (SUM(Profit) * 100.0 / SUM(Sales)),
        2
    ) AS Profit_Margin
FROM df
GROUP BY Region
ORDER BY Profit_Margin DESC
"""

profit_margin = sqldf(profit_margin_query)

print("\n=== PROFIT MARGIN BY REGION ===")
print(profit_margin)
discount_query = """
SELECT
    Region,
    ROUND(AVG(Discount),3) AS Avg_Discount
FROM df
GROUP BY Region
ORDER BY Avg_Discount DESC
"""

discount_summary = sqldf(discount_query)

print("\n=== DISCOUNT BY REGION ===")
print(discount_summary)
loss_query = """
SELECT
    REPLACE([Sub-Category],'-','_') AS Sub_Category,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
GROUP BY [Sub-Category]
HAVING SUM(Profit) < 0
ORDER BY Total_Profit
"""

loss_subcategories = sqldf(loss_query)

print("\n=== LOSS MAKING SUB-CATEGORIES ===")
print(loss_subcategories)
top_customer_query = """
SELECT
    [Customer Name],
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
GROUP BY [Customer Name]
ORDER BY Total_Profit DESC
LIMIT 10
"""

top_profit_customers = sqldf(top_customer_query)

print("\n=== TOP PROFIT CUSTOMERS ===")
print(top_profit_customers)
furniture_query = """
SELECT
    Region,
    Category,
    ROUND(SUM(Sales),2) AS Total_Sales,
    ROUND(SUM(Profit),2) AS Total_Profit
FROM df
WHERE Category = 'Furniture'
GROUP BY Region, Category
"""

furniture_check = sqldf(furniture_query)

print("\n=== FURNITURE PERFORMANCE BY REGION ===")
print(furniture_check)