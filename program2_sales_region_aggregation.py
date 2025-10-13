import pandas as pd
import random

regions = ['North', 'South', 'East', 'West']
products = ['Shoes', 'Bags', 'Shirts', 'Accessories']

# Fake sales data
data = {
    'Region': [random.choice(regions) for _ in range(30)],
    'Product': [random.choice(products) for _ in range(30)],
    'Sales': [round(random.uniform(1000, 10000), 2) for _ in range(30)]
}

df = pd.DataFrame(data)

# --- Transformation for analysis ---
region_summary = df.groupby('Region')['Sales'].agg(['mean', 'sum', 'count']).sort_values(by='sum', ascending=False)
product_pivot = df.pivot_table(values='Sales', index='Region', columns='Product', aggfunc='sum', fill_value=0)

print("Regional Summary:\n", region_summary, "\n")
print("Pivot Table:\n", product_pivot)