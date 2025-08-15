import pandas as pd
import matplotlib.pyplot as plt

# Load data
df = pd.read_csv("C:/Users/Suhani/Downloads/sales_data_sample.csv", encoding='latin1')

# Show first few rows
print(df.head())
# Convert order date to datetime
df['ORDERDATE'] = pd.to_datetime(df['ORDERDATE'])

# Check for missing values
print(df.isnull().sum())
# Drop missing values (if any)
df.dropna(inplace=True)

# Monthly sales trend
monthly_sales = df.groupby(df['ORDERDATE'].dt.to_period('M'))['SALES'].sum()
monthly_sales.plot(kind='line', figsize=(10, 5), title='Monthly Sales Trend')
plt.xlabel('Month')
plt.ylabel('Sales')
plt.grid(True)
plt.tight_layout()
plt.show()

top_products = df.groupby('PRODUCTCODE')['SALES'].sum().sort_values(ascending=False).head(10)
print(top_products)

top_products.plot(kind='bar', title='Top 10 Products by Sales', figsize=(10, 5), color='skyblue')
plt.ylabel('Total Sales')
plt.xlabel('Product Code')
plt.tight_layout()
plt.show()

total_revenue = df['SALES'].sum()
avg_order_value = df.groupby('ORDERNUMBER')['SALES'].sum().mean()
total_orders = df['ORDERNUMBER'].nunique()

print(f"Total Revenue: ${total_revenue:,.2f}")
print(f"Average Order Value: ${avg_order_value:,.2f}")
print(f"Total Orders: {total_orders}")

df.to_excel("sales_analysis_output.xlsx", index=False)
print("✅ Cleaned dataset exported as Excel file.")


import os

# Export the cleaned DataFrame to Excel
df.to_excel("sales_analysis_output.xlsx", index=False)

# Print the full path so you know where the file is saved
print("✅ Excel file saved at:")
print(os.path.abspath("sales_analysis_output.xlsx"))
