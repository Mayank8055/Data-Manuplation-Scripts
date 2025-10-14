import pandas as pd
import random
import re

# Generate fake product review data
products = ['T-Shirt', 'Jeans', 'Shoes', 'Watch', 'Sunglasses']
reviews = [
    "Absolutely love it! Great quality and price.",
    "Terrible, not worth the money.",
    "Good product but delivery was late.",
    "Excellent fit, would buy again!",
    "Worst experience ever.",
    "Average quality, expected better.",
    "Perfect, exactly as shown in pictures.",
    "Not bad, but could be improved."
]

data = {
    'Product': [random.choice(products) for _ in range(25)],
    'Review_Text': [random.choice(reviews) for _ in range(25)],
    'Rating': [random.randint(1, 5) for _ in range(25)]
}

df = pd.DataFrame(data)

# --- Transformations ---
# Clean text
df['Clean_Review'] = df['Review_Text'].apply(lambda x: re.sub(r'[^a-zA-Z\s]', '', x.lower()))

# Sentiment label based on rating
df['Sentiment'] = df['Rating'].map({1: 'Negative', 2: 'Negative', 3: 'Neutral', 4: 'Positive', 5: 'Positive'})

# Aggregation for analysis
sentiment_summary = df.groupby('Product')['Sentiment'].value_counts().unstack(fill_value=0)

print("Cleaned Reviews:\n", df[['Product', 'Clean_Review', 'Rating', 'Sentiment']].head(), "\n")
print("Sentiment Summary by Product:\n", sentiment_summary)