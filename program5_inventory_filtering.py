import pandas as pd
import numpy as np
import random

# Fake inventory data
items = ['Necklace', 'Ring', 'Bracelet', 'Earrings']
data = {
    'Item': [random.choice(items) for _ in range(25)],
    'Stock_Quantity': [random.randint(0, 200) for _ in range(25)],
    'Price': [round(random.uniform(100, 5000), 2) for _ in range(25)]
}

df = pd.DataFrame(data)

# --- Transformation ---
df['Stock_Value'] = df['Stock_Quantity'] * df['Price']
df['Normalized_Price'] = (df['Price'] - df['Price'].min()) / (df['Price'].max() - df['Price'].min())
df_filtered = df[df['Stock_Quantity'] < 50].sort_values(by='Stock_Quantity')

print("Normalized & Filtered Inventory:\n", df_filtered)