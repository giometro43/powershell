Import-Module ImportExcel
Import-Module ActiveDirectory

# Input / output paths
$xlsPath = "C:\Users\gcastillo\Documents\Employee_Master_File_Inquiry_20250728112310230.xlsx"
$outputPath = "C:\Users\gcastillo\Documents\TerminatedUsersReport.csv"

Write-Host "Testing AD connectivity..."
try {
    $testUser = Get-ADUser -Filter * -ResultSetSize 1 | Select-Object SamAccountName
    Write-Host "OK: Able to contact AD. Example user:" $testUser.SamAccountName
} catch {
    Write-Host "ERROR: Could not contact Active Directory. Exiting."
    exit
}

# Import Excel data
$employees = Import-Excel -Path $xlsPath

$results = @()

foreach ($emp in $employees) {
    $first = $emp.'First Name'
    $last = $emp.'Last Name'
    $jobTitleFromXLS = $emp.'Job Class Code Desc'

    # Skip missing names
    if ([string]::IsNullOrWhiteSpace($first) -or [string]::IsNullOrWhiteSpace($last)) {
        $results += [PSCustomObject]@{
            FirstName      = $first
            LastName       = $last
            CSV_JobTitle   = $jobTitleFromXLS
            AD_JobTitle    = "Skipped - Missing Name"
            Terminated     = "N/A"
            JobTitle_Check = "N/A"
            AD_Query       = "Skipped"
        }
        continue
    }

    # Build SamAccountName (first initial + last name, lowercase)
    $username = ($first.Substring(0,1) + $last).ToLower()

    Write-Host "Checking AD for $username..."

    # Query AD by username directly
    $adUser = Get-ADUser -Identity $username -Properties Enabled, Title -ErrorAction SilentlyContinue

    if ($adUser) {
        $terminated = if (-not $adUser.Enabled) { "Yes" } else { "No" }
        $jobTitleMismatch = if ($adUser.Title -ne $jobTitleFromXLS) { "Mismatch" } else { "Match" }

        $results += [PSCustomObject]@{
            FirstName      = $first
            LastName       = $last
            CSV_JobTitle   = $jobTitleFromXLS
            AD_JobTitle    = $adUser.Title
            Terminated     = $terminated
            JobTitle_Check = $jobTitleMismatch
            AD_Query       = $username
        }
    }
    else {
        $results += [PSCustomObject]@{
            FirstName      = $first
            LastName       = $last
            CSV_JobTitle   = $jobTitleFromXLS
            AD_JobTitle    = "Not Found"
            Terminated     = "N/A"
            JobTitle_Check = "User Not Found"
            AD_Query       = $username
        }
    }
}

# Export results
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "Check complete. Results saved to $outputPath"
