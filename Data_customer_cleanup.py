from faker import Faker
import pandas as pd
import random
import numpy as np

fake = Faker()

# Generate fake customer data
data = []
for _ in range(20):
    data.append({
        'Customer_ID': fake.uuid4(),
        'Name': fake.name(),
        'Age': random.randint(18, 60),
        'City': fake.city(),
        'Purchase_Amount': random.choice([None, round(random.uniform(500, 5000), 2)]),
        'Purchase_Date': fake.date_this_year()
    })

df = pd.DataFrame(data)

# --- Data Transformation ---
df['Purchase_Amount'].fillna(df['Purchase_Amount'].mean(), inplace=True)
df['Age_Group'] = pd.cut(df['Age'], bins=[0, 25, 40, 60], labels=['Youth', 'Adult', 'Senior'])
df['Month'] = pd.to_datetime(df['Purchase_Date']).dt.month_name()

print(df.head())