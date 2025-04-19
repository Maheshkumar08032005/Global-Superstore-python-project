import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

sns.set(style="whitegrid")

# Load dataset
df = pd.read_excel(r"E:\Python Report\Global Superstore for python (Ca-2).xlsx")

# Dataset Overview
print("\nDataset Overview\n", df)
print("\nHead\n", df.head())
print("\nTail\n", df.tail())
print("\nSummary\n", df.describe())
print("\nInfo\n")
df.info()
print("\nColumns\n", df.columns)
print("\nShape\n", df.shape)
print("\nNull Values\n", df.isnull().sum())


# Objective 1: Monthly Sales Trend
df['Order Date'] = pd.to_datetime(df['Order Date'])
monthly_sales = df.groupby(df['Order Date'].dt.to_period('M'))['Sales'].sum().to_timestamp()
top4 = monthly_sales.nlargest(4)
plt.figure(figsize=(10, 5))
plt.plot(monthly_sales, label='Monthly Sales', color='steelblue')
plt.scatter(top4.index, top4.values, color='green', s=100, label='Top 4 Peaks')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.legend()
plt.tight_layout()
plt.show()



# Objective 2: Top 10 Most Profitable Products
top_products = df.groupby('Product Name')['Profit'].sum().nlargest(10).reset_index()
plt.figure(figsize=(10, 6))
sns.barplot(data=top_products, x='Profit', y='Product Name', color='teal')
plt.title("Top 10 Most Profitable Products")
plt.tight_layout()
plt.show()



# Objective 3: Average Sales by Ship Mode
print("\nObjective 3: Average Sales by Ship Mode")
avg_sales_shipmode = df.groupby('Ship Mode')['Sales'].mean().reset_index()

plt.figure(figsize=(8, 5))
plt.bar(avg_sales_shipmode['Ship Mode'], avg_sales_shipmode['Sales'], color='purple', edgecolor='black')
plt.title("Average Sales by Ship Mode")
plt.xlabel("Ship Mode")
plt.ylabel("Average Sales")
plt.tight_layout()
plt.show()



# Objective 4: Sales and Profit by Region
region_data = df.groupby('Region')[['Sales', 'Profit']].sum().reset_index()
x = np.arange(len(region_data))
w = 0.35
plt.figure(figsize=(10, 6))
plt.bar(x - w/2, region_data['Sales'], w, label='Sales', color='orange')
plt.bar(x + w/2, region_data['Profit'], w, label='Profit', color='blue')
plt.xticks(x, region_data['Region'], rotation=30, ha='right')
plt.title("Sales and Profit by Region")
plt.legend()
plt.tight_layout()
plt.show()


# Objective 5: Total Profit by Region
print("\nObjective 5: Total Profit by Region")
df.groupby('Region')['Profit'].sum().plot(marker='o', linestyle='-', color='blue', figsize=(8,5))
plt.title("Profit by Region")
plt.xlabel("Region")
plt.ylabel("Total Profit")
plt.grid(True)
plt.tight_layout()
plt.show()

