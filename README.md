# PS-Outlook-User-Report
Creates a spreadsheet report of users based on Outlook address book. I made this in a pinch when a higher up requested a report of users.

This script exports the Global Address List from Outlook to an Excel file, creating an "Outlook Addressbook Report" with the current date in the file name.

## Prerequisites

- PowerShell
- Microsoft Outlook installed and configured
- ImportExcel PowerShell module

## Installation

1. Install the ImportExcel PowerShell module if you haven't already:

   Install-Module ImportExcel -Scope CurrentUser
   
   Below is the GitHub documentation in markdown format for the script. You can copy and paste this into your GitHub project's README.md file:

markdown
Copy code
# Outlook Address Book Exporter

This PowerShell script exports the Global Address List (GAL) from Microsoft Outlook to an Excel file, creating an "Outlook Addressbook Report" with the current date in the file name.

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

The script will create an Excel file named Outlook Addressbook Report(DD.MM.YYYY).xlsx in the same directory as the script. The file contains a "Main" worksheet with the Name, Site, and Phone Number columns populated with the respective information for each entry in the GAL.
