"""
This python script will read your excel file and mask name references to protect the named information.
Name your files correctly and save the file as "All_Ticket_last_6_Months_output.xlsx", in xlsx format.
Also, the sheet_name has to be 'Sheet1'
Requires, pandas, openpyxl modules for this to work.
"""

import pandas as pd
import re

# Function to mask the name
def mask_name(name):
    words = name.split()
    masked_words = ['*' * len(word) for word in words]
    return ' '.join(masked_words)

# Read the Excel file with masked emails
file_path = 'C:\\Python312\\Scripts\\All_Ticket_last_6_Months_output.xlsx'
sheet_name = 'Sheet1'  # Change this to your sheet name if different

# Read the Excel file
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Define a regex pattern to find names (assuming names have at least two words)
name_pattern = r'\b[A-Z][a-z]+\s[A-Z][a-z]+\b'

# Replace names in all columns
for column in data.columns:
    data[column] = data[column].astype(str).apply(lambda x: re.sub(name_pattern, lambda y: mask_name(y.group()), x))

# Save the modified data to a new Excel file
output_file = 'C:\\Python312\\Scripts\\All_Ticket_last_6_Months_output_names_removed.xlsx'
data.to_excel(output_file, index=False)
