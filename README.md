# PS-Outlook-User-Report
Creates a spreadsheet report of users based on Outlook address book. I made this in a pinch when a higher up requested a report of users.

This script exports the Global Address List from Outlook to an Excel file, creating an "Outlook Addressbook Report" with the current date in the file name.

## Prerequisites

- PowerShell
- Microsoft Outlook installed and configured
- ImportExcel PowerShell module

## Installation

1. Install the ImportExcel PowerShell module if you haven't already:

   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   
## Usage
Save the script as Export-OutlookAddressBook.ps1 in your desired directory.

Open a PowerShell window and navigate to the directory where the script is saved.

Run the script:
.\Export-OutlookAddressBook.ps1
