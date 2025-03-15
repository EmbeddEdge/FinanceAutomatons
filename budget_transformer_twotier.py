import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# Define the file paths
csv_file = 'your_csv_file.csv'
spreadsheet = 'your_spreadsheet.xlsx'

#Each month is on a specific Column so the column for the corresponding month must be defined
column_index = 3 #Column D for February

# Load the CSV file into a DataFrame
columns = ['Transaction Date', 'Parent Category', 'Money In', 'Money Out', 'Fee', 'Category']
transactions_df = pd.read_csv(csv_file, usecols = columns)

#Merge Money Out and Fees into one column
#Convert columns to numeric, replacing any non-numeric values with 0
transactions_df['Money Out'] = pd.to_numeric(transactions_df['Money Out'], errors='coerce').fillna(0)
transactions_df['Fee'] = pd.to_numeric(transactions_df['Fee'], errors='coerce').fillna(0)
#Add the columns together
transactions_df['Money Out'] = transactions_df['Money Out'] + transactions_df['Fee']
transactions_df.drop(columns=['Fee'], inplace=True)
print(transactions_df)

# Define Aggregate values by category
category_totals = transactions_df.groupby('Category').agg({
    'Money In': 'sum',
    'Money Out': 'sum'
}).fillna(0)
category_totals_O = (category_totals['Money In'].fillna(0).astype(float) + category_totals['Money Out'].fillna(0).astype(float))
print(category_totals_O)

# Define category mapping
# The left hand side is the category names from the Input CSV and the right hand side are names for the output spreadsheet
category_mapping = {
    'Transfer': 'Refunds and paybacks',
    'Interest': 'Side hustle income and interest',
    'Other Income': 'Side hustle income and interest',
    '': 'Emergencies',
    'Rent': 'Rent',
    '': 'Subscriptions',
    'Fees': 'Bank fees',
    'Medical Aid': 'Medical aid',
    'Groceries': 'Groceries',
    '': 'Insurance',
    'Licence': 'Transport',
    'Parking': 'Transport',
    'Vehicle Maintenance': 'Transport',
    'Public Transport': 'Transport',
    'Fuel': 'Transport',
    'Electricity': 'Electricity',
    '': 'Water',
    'Internet': 'Internet',
    '': 'Messing around/Sus',
    'Cash Deposit': 'Side hustle income and interest',
    'Cash Withdrawal': 'Cash',
    'Education': 'Education',
    'Other Communication': 'Communication',
    'Cellphone': 'Communication',
    'Donations': 'Donations to JW',
    'Home Improvements': 'Shopping',
    'Alcohol': 'Eating out & takeaways',
    'Takeaways': 'Eating out & takeaways',
    'Restaurants': 'Eating out & takeaways',
    'Housekeeping': 'Personal & Entertainment',
    'Movies': 'Personal & Entertainment',
    'Digital Subscriptions': 'Personal & Entertainment',
    'Online Store': 'Personal & Entertainment',
    'Sport & Hobbies': 'Personal & Entertainment',
    'Gadgets': 'Personal & Entertainment',
    'Software/Games': 'Personal & Entertainment',
    'DonationsFN': 'Family',
    'Personal Care': 'Personal care',
    'Garden': 'Personal care',
    '': 'Adventures and holidays',
    'Savings': 'Transfers between accounts',
    'Transfer': 'Transfers between accounts',
}

# Load the Excel spreadsheet
workbook = load_workbook(spreadsheet)
sheet = workbook["Monthly Spending"]

# Initialize a dictionary to accumulate values for each Excel category
excel_totals = {}

# Accumulate values for each Excel category
for csv_category, total_value in category_totals_O.items():
    excel_category = category_mapping.get(csv_category)
    if excel_category:
        excel_totals[excel_category] = excel_totals.get(excel_category, 0) + abs(total_value)

# Update Excel with accumulated totals
for row in sheet.iter_rows():
    category = row[0].value
    if category in excel_totals:
        row[column_index].value = excel_totals[category]
        #print(f"Updated {category} to {excel_totals[category]}")

# Save the updated Excel file
workbook.save(spreadsheet)
