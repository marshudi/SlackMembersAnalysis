import pandas as pd
import os

# Define full paths for files
report_folder = "report"
raw_folder = "Raw"

# Define file names
main_file = os.path.join(report_folder, "Slack_User_Analysis_report.xlsx")
compare_file = os.path.join(raw_folder, "NCteam.xlsx")

# Load the main Excel file
xls = pd.ExcelFile(main_file)
print("Available Sheets in Main File:", xls.sheet_names)

# Ensure the correct sheet name exists
sheet_name = "Joined_Users"  # Change if needed
if sheet_name not in xls.sheet_names:
    raise ValueError(f"❌ Sheet '{sheet_name}' not found in {main_file}. Available sheets: {xls.sheet_names}")

# Read the main Excel file (Joined_Users sheet)
main_df = pd.read_excel(main_file, sheet_name=sheet_name)

# Read the comparison Excel file
compare_df = pd.read_excel(compare_file)

# **Fix column names by stripping spaces**
compare_df.columns = compare_df.columns.str.strip()

# Ensure required columns exist
if "email" not in compare_df.columns or "left" not in compare_df.columns:
    raise ValueError("❌ 'email' or 'left' column not found in NCteam.xlsx after cleaning column names.")

# Convert emails to lowercase for case-insensitive matching
main_df["email"] = main_df["email"].astype(str).str.lower()
compare_df["email"] = compare_df["email"].astype(str).str.lower()
compare_df["left"] = compare_df["left"].astype(str).str.lower()  # Normalize "left" column

# Filter the comparison data to only include rows where 'left' = 'yes'
filtered_compare_df = compare_df[compare_df["left"] == "yes"]

# Find matching emails where 'left' = 'yes'
matching_emails = main_df[main_df["email"].isin(filtered_compare_df["email"])]

# Export the matching records to the report folder
output_file = os.path.join(report_folder, "matching_users_vendors.xlsx")
matching_emails.to_excel(output_file, index=False)

print(f"✅ Matching users exported to {output_file}")
