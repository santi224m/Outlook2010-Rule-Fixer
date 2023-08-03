# Outlook 2010 Rule Fixer

The scripts in this repository can be used to more quickly fix an Outlook rules error that occurs when moving an Outlook PST file from one computer to another. The error occurs for "Move messages from someone to a folder" rules. When the PST file is moved to another computer, Outlook cannnot find the folder on the new computer, so the rule has no specified folder for which to move the messages. In order to fix this manually, you have to reselect the folder for each rule that has this error. If you have hundreds of these rules, this can take a very long time. The scripts in this repository attempt to make this process faster by letting you specify the destination folder for each email address in Excel.

## Overview

The scripts should be used in the following order:

1. Import ```PrintFolder.bas```, ```PrintRecepients.bas```, and ```UpdateFolder.bas``` into Outlook. These three files are found in the ```/VBA``` directory
2. Run the ```PrintRecepients``` module in Outlook and copy the output to ```/Lists/recipients.txt```
3. Rune the ```PrintFolders``` module in Outlook and copy the output to ```/Lists/folders.txt```
4. Run ```Python/list_to_excel.py```. This will insert all the email addresses and folder names into ```output/Rules.xlsx```. Now you can specify a destination folder for each email address. The email addresses are sorted by domain to allow to you quickly set the same folder for all emails with the same domain.
5. Fill out colum B for each email address with the destination folder name
6. Run ```Python/excel_to_VBA_dict.py``` to create a dictionary that maps a folder to each email address using the ```Rules.xlsx``` excel file from the previous step. This script will save the definitions to ```output/dictionary.txt```
7. Copy and past the text in ```output/dictionary.txt``` to ```VBA/UpdateFolder.bas``` right under the line ```Set dictAddressToFolder = CreateObject("Scripting.Dictionary")```
8. Run ```VBA/UpdateFolder.bas``` to set the destination folder for each rule that does not have a specified distination folder. After running the script, all rules that had an error due to an unspecified folder should now be working.

Note: ```VBA/PrintFolders``` only prints folders that are located within the main inbox. If your destination folders are located outside the Inbox folder, modify the script to find these instead. You will also need to modify ```VBA/UpdateFolder``` to locate folders outside the Inbox folder.
