from faker import Faker
import pandas as pd
import random

fake = Faker()

# Generate employee data
data = []
for _ in range(15):
    data.append({
        'Employee_ID': fake.uuid4(),
        'Name': fake.name(),
        'Department': random.choice(['HR', 'Sales', 'IT', 'Finance']),
        'Base_Salary': random.randint(30000, 90000)
    })

df = pd.DataFrame(data)

# --- Transformations ---
df['Bonus'] = df['Department'].map({'HR': 0.05, 'Sales': 0.10, 'IT': 0.08, 'Finance': 0.07})
df['Total_Salary'] = df['Base_Salary'] * (1 + df['Bonus'])
df['Salary_Level'] = pd.cut(df['Total_Salary'], bins=[0, 50000, 80000, 100000], labels=['Low', 'Medium', 'High'])

print(df)