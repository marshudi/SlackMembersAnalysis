# Slack Members Analysis

## Overview
This project automates the analysis of Slack member data by processing two reports exported from Slack. The scripts generate structured insights, summarizing member status, billing information, and activity levels.

## Input Files
The analysis requires **two reports** downloaded from Slack:

1. **Slack Members Report** (`slack-vodafoneoman-members.xlsx` or `.csv`)
   - Contains user details such as `username, email, status, billing-active, has-2fa, has-sso, userid, fullname, displayname, expiration-timestamp`.

2. **Active Users Report** (`Vodafone Oman Member Analytics_3m.csv`)
   - Contains user activity details such as `Name, What I Do, Account type, Account created (UTC), Days active, Messages posted`.

## Scripts and Workflow
The workflow consists of four Python scripts, each performing a specific analysis task.

### 1. `SlackMembersAnalysis.py`
**Purpose**: Processes the Slack members report to categorize users by domain, summarize their status, and filter billing-active members.

- **Outputs**:
  - `output_workbook.xlsx` → Sheets by domain, status summary, billing-active summary.
  - `domain_summary.xlsx` → Pivot table summary of domain vs. status vs. billing-active.
  - `vodafone_active_modified.xlsx` → Segregates Vodafone active users from other billing-active users.

### 2. `SlackActiveDays.py`
**Purpose**: Processes the active user report, calculates account age, and measures user engagement levels.

- **Outputs**:
  - `analyzed_report_days_active.xlsx` → Adds `Days Alive`, `Rank`, and `Days Active in Percentage` to the dataset.

### 3. `CompareNames.py`
**Purpose**: Merges the data from both reports based on user names and identifies unmatched users.

- **Outputs**:
  - `Slack_User_Analysis_report.xlsx` → Contains `Joined_Users` (matched records) and `Not_Found` (unmatched users).

### 4. `CompareEmailFromVendors.py`
**Purpose**: Compares Slack member data with an external vendor list to find matching users who have left the company.

- **Functionality:**
  - Loads `Slack_User_Analysis_report.xlsx` from `report/`.
  - Reads vendor data from `NCteam.xlsx` in `Raw/`.
  - Filters the vendor list to only consider users where `left = 'yes'`.
  - Finds matching emails in the Slack dataset.
  - Exports results to `matching_users_vendors.xlsx` in `report/`.

- **Outputs**:
  - `matching_users_vendors.xlsx` → Contains Slack users who are also in the vendor list with `left = 'yes'`.

## Installation & Requirements
Ensure you have Python 3.x installed along with the required dependencies:

```bash
pip install pandas openpyxl xlsxwriter
```

## Running the Analysis
1. **Ensure the input files are placed correctly.**
2. **Run the scripts in the following order:**
   ```bash
   python SlackMembersAnalysis.py
   python SlackActiveDays.py
   python CompareNames.py
   python CompareEmailFromVendors.py
   ```
3. **Check the `report` folder** for the generated Excel files.

## Output Files & Summary
| File Name | Description |
|-----------|-------------|
| `output_workbook.xlsx` | Categorized member details by domain and billing summary. |
| `domain_summary.xlsx` | Pivot summary of domain vs. status vs. billing-active. |
| `vodafone_active_modified.xlsx` | Filtered active Vodafone users vs. other billing-active users. |
| `analyzed_report_days_active.xlsx` | Adds engagement metrics (Days Alive, Rank, etc.). |
| `Slack_User_Analysis_report.xlsx` | Merged user data with matched and unmatched users. |
| `matching_users_vendors.xlsx` | Slack users who match vendor data where `left = 'yes'`. |

## Notes
- Ensure the sheet names are properly formatted, as long domain names may be truncated.
- Adjust file paths in the scripts if necessary.

**Happy Analyzing!** 🚀

