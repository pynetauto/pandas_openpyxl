"""
This python script will read your excel file and mask all email alias to protect the email information.
Name your files correctly and save the file as "All_Ticket_last_6_Months.xlsx", in xlsx format.
Also, the sheet_name has to be 'Sheet1'
Requires, pandas, openpyxl modules for this to work.
"""

import pandas as pd
import re

# Function to mask the email
def mask_email(email):
    if '@' in email:
        username, domain = email.split('@')
        username = username[0] + '*' * (len(username) - 1)
        domain_parts = domain.split('.')
        masked_domain = domain_parts[0][0] + '*' * (len(domain_parts[0]) - 1) + '.' + domain_parts[1]
        return f"{username}@{masked_domain}"
    else:
        return email

# Read the Excel file
file_path = 'C:\\Python312\\Scripts\\All_Ticket_last_6_Months.xlsx'
sheet_name = 'Sheet1'  # Change this to your sheet name if different

# Read the Excel file
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Define a regex pattern to find emails
email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Replace emails in all columns
for column in data.columns:
    data[column] = data[column].astype(str).apply(lambda x: re.sub(email_pattern, lambda y: mask_email(y.group()), x))

# Save the modified data to a new Excel file
output_file = 'C:\\Python312\\Scripts\\All_Ticket_last_6_Months_output.xlsx'
data.to_excel(output_file, index=False)
