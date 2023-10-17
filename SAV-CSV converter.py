import pyreadstat
import pandas as pd

# Specify the path to your .sav file
sav_file_path = 'Project BHT – W1-Aug-23_Pilot Data.sav'

# Read the .sav file
df, meta = pyreadstat.read_sav(sav_file_path)

# Specify the path for the CSV file
csv_file_path = 'Project BHT – W1-Aug-23_Pilot Data.csv'

# Save the data to a CSV file
df.to_csv(csv_file_path, index=False)

print(f"Conversion complete. CSV file saved at: {csv_file_path}")
