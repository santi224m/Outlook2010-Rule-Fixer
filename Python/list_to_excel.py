from openpyxl import load_workbook

# Filepath constants
EXCEL_FILEPATH = "../output/Rules.xlsx"
RECIPIENTS_TXT_FILEPATH = "../Lists/recipients.txt"
FOLDERS_TXT_FILEPATH = "../Lists/folders.txt"

# Extract domain from email address
def getEmailDomain(email):
  return email[email.index('@') + 1 :]

# Extract email addresses from recipients.txt and store in recipients array
recipients_file = open(RECIPIENTS_TXT_FILEPATH, 'r')
recipients = list(set(recipients_file.read().split()))
recipients.sort(key=getEmailDomain) # Sort addresses by domain

# Extract folder names from folders.txt and store in folders array
folders_file = open(FOLDERS_TXT_FILEPATH, 'r')
folders = list(set(folders_file.read().split()))
folders.sort() # Sort folders in alphabetical order

# Load excel file
excel_book = load_workbook(EXCEL_FILEPATH)
excel_sheet = excel_book["Sheet1"]

# Insert recipients to excel file
recipients_index = 2 # Start at 2 because there is a header at cell 1
for recipient in recipients:
  excel_sheet[f"A{str(recipients_index)}"] = recipient
  recipients_index += 1

# Insert folder names into excel file
folders_index = 2 # Start at 2 because there is a header at cell 1
for folder in folders:
  excel_sheet[f"G{folders_index}"] = folder
  folders_index += 1

# Save excel file
excel_book.save(EXCEL_FILEPATH)