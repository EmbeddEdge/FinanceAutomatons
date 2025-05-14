# Personal Finance Automation Scripts

This repository contains a collection of Python scripts designed to automate personal finance tasks. These scripts accept CSV inputs and input the data into a budget tracking spreadsheet, making it easier to manage and track your finances.

## Features

- **CSV Input**: The scripts are designed to accept CSV files as input, allowing you to easily import your financial data.
- **Budget Tracking**: The scripts automatically update a budget tracking spreadsheet with the imported data, providing an overview of your income, expenses, and savings.
- **Automation**: By automating the process of inputting financial data, these scripts save you time and effort, allowing you to focus on other important tasks.

## Getting Started

### Quick demo

To see how the scripts work you can use the sample files provided. 
First check the 'Money Dashboard 2025' spreadsheet file and confirm that the 'Monthly Spending' sheet has no entries.
Second have a look at the 'SampleBankExport' CSV file and check the numbers and the categories.
Run the 'budget_transformer.py' script
Check the spreadsheet again and compare the now filled in 'Monthly Spending' sheet as a summation of the line items from the CSV files.

### How to use this script

To modify this program for your usage

1. Download your own financial data in a CSV format, and modify it to replicate the sample file format
2. Check the budget tracking spreadsheet to view the updated financial information.
3. Updated the script to match the file names if you change the names to fit your own files. Still keeping the same format.
4. Run the script.
5. Optionally you can check the files manually to verify. 



## Contributing

Contributions are welcome! If you have any ideas, suggestions, or bug reports, please open an issue or submit a pull request. 

## Long-term Goals

The idea is to create a full set of scripts that can automate as much as possible in the process of personal finance or for small businesses. 
The ideal output would be to generate monthly reports with or without suggestions that will help a person spend better in the next month.
The ideal workflow is to remove as much manual entry and error-checking as possible. 
As close to real-time is another goal so that the data can be updated weekly or possibly even on a daily basis.

## Current Status

- **Manual Data Acquisition**: The CSV data needs to be manually downloaded from online banking or the banking app. It would be better if we could automate this.
- **Manual Categorisation**: There is some auto-categorization that happens on the banking app but it still takes time for this to be checked. 