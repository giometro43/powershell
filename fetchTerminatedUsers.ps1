<# SYNOPSIS: 
    Based on a xls file, compare the users in the file to the users in AD, 

    identify if the users are still active or terminated, and output the results to the console and a CSV file.
#>

# Here is what a begginner would need to use to make this script: 
# 1. Install the AzureAD module if not already installed
#    Install-Module -Name AzureAD
# 2. Connect to Azure AD
#    Connect-AzureAD
# 3. Make sure you have the necessary permissions to read user information in Azure AD
# 4. Update the path to the input Excel file and output CSV file as needed
# 5. Install the ImportExcel module if not already installed
#    Install-Module -Name ImportExcel
# 6. Import the ImportExcel module
#    Import-Module ImportExcel


Import-Module AzureAD
Import-Module ImportExcel

#
# Path to the input Excel file
$inputFilePath = "C:\Users\gcastillo\Documents\Employee_Master_File_Inquiry_20250728112310230.xlsx"

# a csv file to output the results
$outputFilePath = "C:\Users\gcastillo\Documents\TerminatedUsersReport.csv"

#connect to azure ad
Connect-AzureAD

# Import the Excel file
$usersFromExcel = Import-Excel -Path $inputFilePath
if ($null -eq $usersFromExcel) {
    Write-Error "Failed to import data from the Excel file. Please check the file path and format."
    exit
}
Write-Host "Imported $($usersFromExcel.Count) users from the Excel file."
# Prepare an array to hold the results
$results = @()
# Loop through each user in the Excel file, the row names are Last Name, First Name, job title
foreach ($user in $usersFromExcel) {
    $firstName = $user.'First Name'
    $lastName = $user.'Last Name'
    $jobTitle = $user.'Job Title'

    # Search for the user in Azure AD
    $aadUser = Get-AzureADUser -All $true | Where-Object { $_.GivenName -eq $firstName -and $_.Surname -eq $lastName }

    if ($null -ne $aadUser) {
        # User found in Azure AD
        $isAccountEnabled = $aadUser.AccountEnabled
        if ($isAccountEnabled) {
            $status = "Active"
        } else {
            $status = "Terminated"
        }
    } else {
        # User not found in Azure AD
        $status = "Not Found"
    }

    # Create a result object
    $result = [PSCustomObject]@{
        'First Name'  = $firstName
        'Last Name'   = $lastName
        'Job Title'   = $jobTitle
        'Status'      = $status
    }
    # Add the result to the results array
    $results += $result 
}

# Output the results to the console
$results | Format-Table -AutoSize
# Export the results to a CSV file
$results | Export-Csv -Path $outputFilePath -NoTypeInformation
Write-Host "Results exported to $outputFilePath"
# Disconnect from Azure AD
Disconnect-AzureAD
Write-Host "Disconnected from Azure AD"
# End of script
# Note: Make sure to run this script in an environment where you have the necessary permissions and modules installed.


#confirm if their Job Title is the same as the one in AD, if not, output the first and last name of users and their job title in AD vs the one in the excel file
foreach ($user in $usersFromExcel) {
    $firstName = $user.'First Name'
    $lastName = $user.'Last Name'
    $jobTitle = $user.'Job Title'

    # Search for the user in Azure AD
    $aadUser = Get-AzureADUser -All $true | Where-Object { $_.GivenName -eq $firstName -and $_.Surname -eq $lastName }

    if ($null -ne $aadUser) {
        # User found in Azure AD
        $aadJobTitle = $aadUser.JobTitle
        if ($aadJobTitle -ne $jobTitle) {
            Write-Host "Job title mismatch for $firstName $lastName: Excel Job Title = '$jobTitle', AD Job Title = '$aadJobTitle'"
        }
    }
}

# End of script
# Note: Make sure to run this script in an environment where you have the necessary permissions and modules installed.
#confirm if their Job Title is the same as the one in AD, if not, output the first and last name of users and their job title in AD vs the one in the excel file
foreach ($user in $usersFromExcel) {
    $firstName = $user.'First Name'
    $lastName = $user.'Last Name'
    $jobTitle = $user.'Job Title'

    # Search for the user in Azure AD
    $aadUser = Get-AzureADUser -All $true | Where-Object { $_.GivenName -eq $firstName -and $_.Surname -eq $lastName }

    if ($null -ne $aadUser) {
        # User found in Azure AD
        $aadJobTitle = $aadUser.JobTitle
        if ($aadJobTitle -ne $jobTitle) {
            Write-Host "Job title mismatch for $firstName $lastName: Excel Job Title = '$jobTitle', AD Job Title = '$aadJobTitle'"
        }
    }
}
# Here is what a begginner would need to use to make this script:
# 1. Install the AzureAD module if not already installed
#    Install-Module -Name AzureAD
# 2. Connect to Azure AD
#    Connect-AzureAD
# 3. Make sure you have the necessary permissions to read user information in Azure AD
# 4. Update the path to the input Excel file and output CSV file as needed
# 5. Install the ImportExcel module if not already installed
#    Install-Module -Name ImportExcel
# 6. Import the ImportExcel module

