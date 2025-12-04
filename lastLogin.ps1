# ============================
# Configuration
# ============================
$HostList = "C:\Users\YOURUSERNAME\Desktop\win10hostnames.txt"

$Days60   = 60
$Days180  = 180

$Cutoff60  = (Get-Date).AddDays(-$Days60)
$Cutoff180 = (Get-Date).AddDays(-$Days180)

Import-Module ActiveDirectory

# ============================
# Load Hostnames
# ============================
$Computers = Get-Content $HostList
$Results   = @()

foreach ($Computer in $Computers) {

    # Try to read AD object
    $ADComputer = Get-ADComputer -Identity $Computer -Properties lastLogonTimestamp -ErrorAction SilentlyContinue

    if ($ADComputer) {
        $LastLogon = [DateTime]::FromFileTime($ADComputer.lastLogonTimestamp)

        $Results += [PSCustomObject]@{
            ComputerName = $Computer
            LastLogon    = $LastLogon
        }
    }
    else {
        $Results += [PSCustomObject]@{
            ComputerName = $Computer
            LastLogon    = $null
        }
    }
}

# ============================
# Categorize Systems
# ============================
$Active60       = $Results | Where-Object { $_.LastLogon -ne $null -and $_.LastLogon -ge $Cutoff60 }
$Inactive60     = $Results | Where-Object { $_.LastLogon -ne $null -and $_.LastLogon -lt $Cutoff60 }
$Inactive180    = $Results | Where-Object { $_.LastLogon -ne $null -and $_.LastLogon -lt $Cutoff180 }
$NotFound       = $Results | Where-Object { $_.LastLogon -eq $null }

# ============================
# Summary Counts
# ============================
$ActiveCount       = $Active60.Count
$InactiveCount     = $Inactive60.Count
$Inactive180Count  = $Inactive180.Count
$NotFoundCount     = $NotFound.Count

Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
Write-Host "Active (Last 60 Days):     $ActiveCount"
Write-Host "Inactive (Over 60 Days):   $InactiveCount"
Write-Host "Inactive (Over 6 Months):  $Inactive180Count"
Write-Host "Not Found in AD:           $NotFoundCount"
Write-Host "=====================`n" -ForegroundColor Cyan

# ============================
# Console Output Tables
# ============================
Write-Host "`n===== LOGGED IN WITHIN LAST 60 DAYS =====" -ForegroundColor Green
$Active60 | Sort-Object LastLogon -Descending | Format-Table -AutoSize

Write-Host "`n===== NOT LOGGED IN WITHIN 60 DAYS =====" -ForegroundColor Yellow
$Inactive60 | Sort-Object LastLogon | Format-Table -AutoSize

Write-Host "`n===== NOT LOGGED IN WITHIN 6 MONTHS =====" -ForegroundColor Red
$Inactive180 | Sort-Object LastLogon | Format-Table -AutoSize

Write-Host "`n===== NOT FOUND IN AD =====" -ForegroundColor Magenta
$NotFound | Format-Table -AutoSize

# ============================
# Build CSV Rows
# ============================
$FinalExport = @()

# Category: Active < 60 days
foreach ($item in $Active60) {
    $FinalExport += [PSCustomObject]@{
        ComputerName = $item.ComputerName
        LastLogon    = $item.LastLogon
        Category     = "Active_60_Days"
        CategoryCount = ""
    }
}

# Category: Inactive > 60 days
foreach ($item in $Inactive60) {
    $FinalExport += [PSCustomObject]@{
        ComputerName = $item.ComputerName
        LastLogon    = $item.LastLogon
        Category     = "Inactive_Over_60_Days"
        CategoryCount = ""
    }
}

# Category: Inactive > 180 days
foreach ($item in $Inactive180) {
    $FinalExport += [PSCustomObject]@{
        ComputerName = $item.ComputerName
        LastLogon    = $item.LastLogon
        Category     = "Inactive_Over_180_Days"
        CategoryCount = ""
    }
}

# Category: Not Found in AD
foreach ($item in $NotFound) {
    $FinalExport += [PSCustomObject]@{
        ComputerName = $item.ComputerName
        LastLogon    = ""
        Category     = "Not_Found"
        CategoryCount = ""
    }
}

# ============================
# Add Summary Rows at Bottom
# ============================
$FinalExport += [PSCustomObject]@{
    ComputerName   = "TOTAL_Active_60_Days"
    LastLogon      = ""
    Category       = "Active_60_Days"
    CategoryCount  = $ActiveCount
}

$FinalExport += [PSCustomObject]@{
    ComputerName   = "TOTAL_Inactive_Over_60_Days"
    LastLogon      = ""
    Category       = "Inactive_Over_60_Days"
    CategoryCount  = $InactiveCount
}

$FinalExport += [PSCustomObject]@{
    ComputerName   = "TOTAL_Inactive_Over_180_Days"
    LastLogon      = ""
    Category       = "Inactive_Over_180_Days"
    CategoryCount  = $Inactive180Count
}

$FinalExport += [PSCustomObject]@{
    ComputerName   = "TOTAL_Not_Found"
    LastLogon      = ""
    Category       = "Not_Found"
    CategoryCount  = $NotFoundCount
}

# ============================
# Export CSV
# ============================
$FinalExport | Export-Csv "C:\Users\YOURUSERNAME\Desktop\LoginReportw10.csv" -NoTypeInformation

