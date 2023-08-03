from openpyxl import load_workbook

EXCEL_FILEPATH = "../output/Rules.xlsx"
DICTIONARY_FILEPATH = "../output/dictionary.txt"

def append_to_dict_entry(entryStr):
  with open(DICTIONARY_FILEPATH, "a") as dict_file:
    dict_file.write(entryStr + "\n")

# Load excel workbook
excel_book = load_workbook(EXCEL_FILEPATH)
excel_sheet = excel_book["Sheet1"]

# Output VBA dictionary definition for each recipient using excel file
row = 2
while True:
  # Create cells for current row
  cell_recipient = 'A' + str(row)
  cell_folder = 'B' + str(row)
  # Continue until empty cell_recipient encountered
  if excel_sheet[cell_recipient].value is not None:
    # VBA dictionary name: dictAddressToFolder
    # Entry format: dictAddressToFolder.Add "email address", "folder name"
    VBA_dict_entry = f"dictAddressToFolder.Add \"{excel_sheet[cell_recipient].value}\", \"{excel_sheet[cell_folder].value}\""
    append_to_dict_entry(VBA_dict_entry)
    row += 1
  else:
    break