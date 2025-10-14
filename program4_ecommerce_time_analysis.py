import pandas as pd
import random
from faker import Faker

fake = Faker()

# Generate fake order data
data = []
for _ in range(20):
    order_time = fake.date_time_this_year()
    data.append({
        'Order_ID': fake.uuid4(),
        'Order_Timestamp': order_time,
        'Customer': fake.name(),
        'Order_Value': round(random.uniform(100, 2000), 2)
    })

df = pd.DataFrame(data)

# --- Transformations ---
df['Order_Timestamp'] = pd.to_datetime(df['Order_Timestamp'])
df['Day_of_Week'] = df['Order_Timestamp'].dt.day_name()
df['Hour'] = df['Order_Timestamp'].dt.hour
df['Order_Period'] = pd.cut(df['Hour'], bins=[0, 12, 18, 24], labels=['Morning', 'Afternoon', 'Night'], right=False)

print(df.head())