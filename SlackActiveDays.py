import pandas as pd
from datetime import datetime
import os

# Define file paths using your provided paths
input_file = r"\Raw\Vodafone Oman Member Analytics_3m.csv"
output_folder = r"\report"
output_file = os.path.join(output_folder, 'analyzed_report_days_active.xlsx')

# Create the output folder if it does not exist
os.makedirs(output_folder, exist_ok=True)

# Read the CSV file
df = pd.read_csv(input_file)

# Convert the "Account created (UTC)" column to datetime using the correct format
df["Account created (UTC)"] = pd.to_datetime(df["Account created (UTC)"], format="%b %d, %Y")

# Calculate "Days Alive" as the difference between today's date and the account creation date
today = datetime.today()
df["Days Alive"] = (today - df["Account created (UTC)"]).dt.days

# Cap "Days Alive" at 60 if the value is greater than 60
df["Days Alive"] = df["Days Alive"].apply(lambda x: x if x <= 60 else 60)

# Calculate "Rank" using the formula: (Days Alive - Days active)
df["Rank"] = df["Days Alive"] - df["Days active"]

# Calculate "Days Active in Percentage" as ((Days Alive - Rank) / Days Alive)
# Since (Days Alive - Rank) is equal to Days active, this is equivalent to Days active / Days Alive.
# We include a check to avoid division by zero.
df["Days Active in Percentage"] = df.apply(
    lambda row: (row["Days Alive"] - row["Rank"]) / row["Days Alive"] if row["Days Alive"] != 0 else 0, axis=1
)

# Save the analyzed DataFrame to an Excel file in the specified folder
df.to_excel(output_file, index=False)

print(f"Analyzed report saved to: {output_file}")
