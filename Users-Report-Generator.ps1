# Required module: ImportExcel
# Install it for the current user if not already installed
Install-Module ImportExcel -Scope CurrentUser

# Get the script path, directory, and set the output file name with date
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$date = Get-Date -Format d
$date = $date.Replace("/", ".")
$finaldir = $dir + "\Outlook Addressbook Report(" + $date + ").xlsx"

# Initialize index for array manipulation
$index = 0

# Function to enumerate Global Address List (GAL) from Outlook
function enumerate-GAL {
    [Microsoft.Office.Interop.Outlook.Application] $outlook = New-Object -ComObject Outlook.Application
    $entries = $outlook.Session.GetGlobalAddressList().AddressEntries
    foreach ($entry in $entries) {
        $entry2 = $entry.getExchangeUser()
        Write-Output $entry2
    }
}

# Get the users from GAL and create a new Outlook instance
$users = enumerate-GAL
$outlook = New-Object -ComObject Outlook.Application

# Create an empty CSV template
$Main = ConvertFrom-Csv @"
Name,Site,Phone Number
"@

# Iterate through users and populate the CSV template
foreach ($email in $users."PrimarySMTPAddress") {
    $Main += [PSCustomObject]@{
        'Name'          = $users."Name"[$index]
        'Site'          = $users."OfficeLocation"[$index]
        'Phone Number'  = $users."BusinessTelephoneNumber"[$index]
    }
    $index++
}

# Export the populated CSV template to an Excel file
$Main | Export-Excel $finaldir -AutoSize -BoldTopRow -WorkSheetName "Main"
