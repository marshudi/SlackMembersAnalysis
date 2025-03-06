import pandas as pd
import os

# -------------------------------
# Define file paths
# -------------------------------
# Make sure to include the file extension (e.g., .xlsx) for the primary file
primary_file = r"\report\vodafone_active_modified.xlsx"  # Update with your actual file path and extension

# Similarly, include the proper extension for the secondary file
secondary_file = r"\report\analyzed_report_days_active.xlsx"  # Update as needed

# Output file path
output_folder = r"\report"
os.makedirs(output_folder, exist_ok=True)
output_file = os.path.join(output_folder, 'joined_users_report.xlsx')

# -------------------------------
# Load the primary file sheets
# -------------------------------
# Read only the two sheets of interest
df_vodafone = pd.read_excel(primary_file, sheet_name="Vodafone_Active")
df_other = pd.read_excel(primary_file, sheet_name="Other_Billing_Active")

# Combine the two dataframes (they have the same structure)
df_primary = pd.concat([df_vodafone, df_other], ignore_index=True)
# Ensure the join key is trimmed of extra spaces (if needed)
df_primary["fullname"] = df_primary["fullname"].str.strip()

# -------------------------------
# Load the secondary file
# -------------------------------
df_secondary = pd.read_excel(secondary_file)
# Ensure the join key is trimmed of extra spaces
df_secondary["Name"] = df_secondary["Name"].str.strip()

# -------------------------------
# Join the dataframes on the name keys
# -------------------------------
# Perform an inner join so only matching records are returned.
df_joined = pd.merge(
    df_primary, 
    df_secondary[["Name", "Days active", "Messages posted", "Days Alive", "Rank", "Days Active in Percentage"]],
    left_on="fullname", 
    right_on="Name", 
    how="inner"
)

# Optionally, drop the duplicate join key column ("Name")
df_joined.drop(columns=["Name"], inplace=True)

# -------------------------------
# Identify records in secondary that did NOT match any primary record
# -------------------------------
df_not_found = df_secondary[~df_secondary["Name"].isin(df_primary["fullname"])].copy()

# Add a note column to label these rows as not found
df_not_found["Note"] = "User not found"

# -------------------------------
# Write the results to an Excel file with multiple sheets
# -------------------------------
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_joined.to_excel(writer, sheet_name="Joined_Users", index=False)
    df_not_found.to_excel(writer, sheet_name="Not_Found", index=False)

print(f"Joined report saved to: {output_file}")
