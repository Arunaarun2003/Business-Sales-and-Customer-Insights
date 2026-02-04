# =========================================================
# Step 1: Import Libraries
# =========================================================
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


# =========================================================
# Step 2: Read and Inspect Data (FIRST TABLE)
# =========================================================

df = pd.read_excel("C:\\Users\\Aruna\\Downloads\\archive (1)\\Sales Dataset.xlsx", engine='openpyxl')
print(df.head())
print(df.info())
print(df.describe())


# =========================================================
# Step 3: Clean the Data
# =========================================================

# Remove duplicate rows
df.drop_duplicates(inplace=True)

# Handle missing values
df['Amount'].fillna(0, inplace=True)
df['Profit'].fillna(0, inplace=True)
df['Quantity'].fillna(0, inplace=True)

# Remove rows where customer name is missing
df.dropna(subset=['CustomerName'], inplace=True)

# Convert Order Date to datetime
df['Order Date'] = pd.to_datetime(df['Order Date'])


# # =========================================================
# # Step 4: Exploratory Data Analysis (EDA)
# # =========================================================

# a) Overall Sales Trend (Monthly)
monthly_sales = df.groupby(df['Order Date'].dt.to_period('M'))['Amount'].sum()
monthly_sales.plot(kind='line', figsize=(10,5), title='Monthly Sales Trend')
plt.ylabel("Total Sales Amount")
plt.show()


# b) Top 10 Sub-Categories by Sales
top_subcategories = (
    df.groupby('Sub-Category')['Amount']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)

sns.barplot(x=top_subcategories.values, y=top_subcategories.index)
plt.title("Top 10 Sub-Categories by Sales")
plt.xlabel("Sales Amount")
plt.show()


# c) Sales by State
state_sales = (
    df.groupby('State')['Amount']
    .sum()
    .sort_values(ascending=False)
)

state_sales.plot(kind='bar', figsize=(12,6), title='Sales by State')
plt.ylabel("Sales Amount")
plt.show()



# d) Top 10 Customers by Sales
top_customers = (
    df.groupby('CustomerName')['Amount']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)

sns.barplot(x=top_customers.values, y=top_customers.index)
plt.title("Top 10 Customers by Sales")
plt.xlabel("Sales Amount")
plt.show()


# e) Correlation between Amount, Profit, and Quantity
sns.heatmap(
    df[['Amount', 'Profit', 'Quantity']].corr(),
    annot=True,
    cmap='coolwarm'
)
plt.title("Correlation Analysis")
plt.show()


# =========================================================
# Step 5: Pivot Table (Sales by State and Sub-Category)
# =========================================================
pivot_table = df.pivot_table(
    index='State',
    columns='Sub-Category',
    values='Amount',
    aggfunc='sum'
)

print(pivot_table)
pivot_table.plot(kind='bar', figsize=(12, 6))
plt.show()


