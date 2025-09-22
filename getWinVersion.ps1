Import-Module ActiveDirectory

# Get all computers
$Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion

# Define EOS keywords (now includes Windows 10)
$EOS_OS = @(
    "Windows XP",
    "Windows 7",
    "Windows 8",
    "Windows 8.1",
    "Windows 10",
    "Windows Vista",
    "Server 2003",
    "Server 2008",
    "Server 2012"   # EOS as of Oct 2023
)

# Build regex pattern for EOS detection
$EOS_Pattern = ($EOS_OS -join "|")

# Separate by OS presence
$Unknown_OS = $Computers | Where-Object { [string]::IsNullOrWhiteSpace($_.OperatingSystem) }
$Known_OS   = $Computers | Where-Object { -not [string]::IsNullOrWhiteSpace($_.OperatingSystem) }

# Split servers vs workstations
$Servers = $Known_OS | Where-Object { $_.OperatingSystem -like "*Server*" }
$Workstations = $Known_OS | Where-Object { $_.OperatingSystem -notlike "*Server*" }

# EOS vs Supported logic (fixed)
$Servers_EOS        = $Servers      | Where-Object { $_.OperatingSystem -match $EOS_Pattern }
$Servers_Supported  = $Servers      | Where-Object { $_.OperatingSystem -notmatch $EOS_Pattern }

$Workstations_EOS        = $Workstations | Where-Object { $_.OperatingSystem -match $EOS_Pattern }
$Workstations_Win11      = $Workstations | Where-Object { $_.OperatingSystem -like "*Windows 11*" }
$Workstations_Supported  = $Workstations | Where-Object { ($_.OperatingSystem -notmatch $EOS_Pattern) -and ($_.OperatingSystem -notlike "*Windows 11*") }

# ---- Console Output ----
Write-Host "=== Workstations (EOS) ===" -ForegroundColor Red
$Workstations_EOS | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize
Write-Host "Count (EOS Workstations): $($Workstations_EOS.Count)`n" -ForegroundColor Yellow

Write-Host "=== Workstations (Windows 11 - Supported) ===" -ForegroundColor Green
$Workstations_Win11 | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize
Write-Host "Count (Win11 Workstations): $($Workstations_Win11.Count)`n" -ForegroundColor Yellow

Write-Host "=== Workstations (Other Supported) ===" -ForegroundColor Cyan
$Workstations_Supported | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize
Write-Host "Count (Supported Workstations): $($Workstations_Supported.Count)`n" -ForegroundColor Yellow

Write-Host "`n==================="

Write-Host "=== Servers (EOS) ===" -ForegroundColor Red
$Servers_EOS | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize
Write-Host "Count (EOS Servers): $($Servers_EOS.Count)`n" -ForegroundColor Yellow

Write-Host "=== Servers (Supported) ===" -ForegroundColor Green
$Servers_Supported | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize
Write-Host "Count (Supported Servers): $($Servers_Supported.Count)`n" -ForegroundColor Yellow

Write-Host "`n==================="

Write-Host "=== Unknown OS (No Attribute) ===" -ForegroundColor Magenta
$Unknown_OS | Select-Object Name | Format-Table -AutoSize
Write-Host "Count (Unknown OS): $($Unknown_OS.Count)`n" -ForegroundColor Yellow

# ---- CSV Export ----
$ExportData = @()

$ExportData += $Workstations_EOS       | Select-Object Name, OperatingSystem, OperatingSystemVersion,@{Name="Category";Expression={"Workstation_EOS"}}
$ExportData += $Workstations_Win11     | Select-Object Name, OperatingSystem, OperatingSystemVersion,@{Name="Category";Expression={"Workstation_Win11"}}
$ExportData += $Workstations_Supported | Select-Object Name, OperatingSystem, OperatingSystemVersion,@{Name="Category";Expression={"Workstation_Supported"}}
$ExportData += $Servers_EOS            | Select-Object Name, OperatingSystem, OperatingSystemVersion,@{Name="Category";Expression={"Server_EOS"}}
$ExportData += $Servers_Supported      | Select-Object Name, OperatingSystem, OperatingSystemVersion,@{Name="Category";Expression={"Server_Supported"}}
$ExportData += $Unknown_OS             | Select-Object Name,@{Name="OperatingSystem";Expression={"N/A"}},@{Name="OperatingSystemVersion";Expression={"N/A"}},@{Name="Category";Expression={"Unknown_OS"}}

$ExportData | Export-Csv "C:\Users\gcastillo\Desktop\ComputerOS_Report.csv" -NoTypeInformation -Encoding UTF8

Write-Host "`nResults exported to ComputerOS_Report.csv" -ForegroundColor Green


