import pandas as pd
import os

# Define the input file name (assumed to be in the same folder as the script)
input_file = r"\Raw\slack-vodafoneoman-members.csv"

# Define the output folder path
output_folder = r"\report"

# Ensure the output directory exists
os.makedirs(output_folder, exist_ok=True)

# Read the CSV file into a DataFrame
df = pd.read_csv(input_file)

# Convert email domain to lowercase and store in "domain" column
df["domain"] = df["email"].apply(lambda x: x.split("@")[-1].lower())

# Optionally, also convert status to lowercase for case-insensitive grouping/filtering
df["status"] = df["status"].str.lower()

# ---------------------------
# 1. Create output_workbook.xlsx:
# ---------------------------
output_excel = os.path.join(output_folder, "output_workbook.xlsx")

def get_unique_sheet_name(base_name, used_names):
    sheet_name = base_name
    counter = 1
    while sheet_name.lower() in used_names:
        sheet_name = f"{base_name}_{counter}"
        counter += 1
    used_names.add(sheet_name.lower())
    return sheet_name

used_sheet_names = set()
with pd.ExcelWriter(output_excel) as writer:
    unique_domains = df["domain"].unique()
    for domain in unique_domains:
        domain_df = df[df["domain"] == domain]
        base_sheet_name = domain.replace(".", "_").replace("@", "_")[:31]
        sheet_name = get_unique_sheet_name(base_sheet_name, used_sheet_names)
        domain_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Exported data for domain '{domain}' to sheet '{sheet_name}'.")
    
    status_summary = df.groupby(["domain", "status"]).size().unstack(fill_value=0)
    status_summary_sheet = get_unique_sheet_name("Status_Summary", used_sheet_names)
    status_summary.to_excel(writer, sheet_name=status_summary_sheet)
    print(f"Exported status summary to sheet '{status_summary_sheet}'.")
    
    # Note: since we converted status to lowercase, use "member" in the filter.
    member_df = df[df["status"] == "member"]
    billing_summary = member_df.groupby(["domain", "billing-active"]).size().unstack(fill_value=0)
    billing_sheet = get_unique_sheet_name("Member_Billing", used_sheet_names)
    billing_summary.to_excel(writer, sheet_name=billing_sheet)
    print(f"Exported member billing summary to sheet '{billing_sheet}'.")

print(f"Data analysis complete. Excel workbook saved as '{output_excel}'.")


# ---------------------------
# 2. Create domain_summary.xlsx:
# ---------------------------
summary_file = os.path.join(output_folder, "domain_summary.xlsx")
pivot_summary = pd.pivot_table(df, index=["domain", "status"], 
                               columns="billing-active", 
                               aggfunc="size", fill_value=0)
pivot_summary = pivot_summary.reset_index()
pivot_summary.columns.name = None
pivot_summary.to_excel(summary_file, index=False)
print(f"Domain summary saved in '{summary_file}'.")


# ---------------------------
# 3. Create vodafone_active_modified.xlsx:
# ---------------------------
output_vodafone_file = os.path.join(output_folder, "vodafone_active_modified.xlsx")

# Filter rows where billing-active equals 1
billing_active_df = df[df["billing-active"] == 1]
vodafone_active_df = billing_active_df[billing_active_df["domain"] == "vodafone.om"]
other_billing_active_df = billing_active_df[billing_active_df["domain"] != "vodafone.om"]
billing_active_summary = billing_active_df.groupby("domain").size().reset_index(name="Active_Count")

with pd.ExcelWriter(output_vodafone_file) as writer:
    vodafone_active_df.to_excel(writer, sheet_name="Vodafone_Active", index=False)
    print("Exported Vodafone active records to sheet 'Vodafone_Active'.")
    
    other_billing_active_df.to_excel(writer, sheet_name="Other_Billing_Active", index=False)
    print("Exported other billing-active records to sheet 'Other_Billing_Active'.")
    
    billing_active_summary.to_excel(writer, sheet_name="Billing_Active_Summary", index=False)
    print("Exported billing-active summary to sheet 'Billing_Active_Summary'.")

print(f"Vodafone active modified workbook saved as '{output_vodafone_file}'.")
