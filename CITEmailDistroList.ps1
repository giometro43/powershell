<#
.SYNOPSIS
This script connects to Exchange Online, filters distribution groups to find those related to the CIT department, retrieves their members, and exports the information to a CSV file.

.DESCRIPTION
The script performs the following steps:
1. Connects to Exchange Online.
2. Filters distribution groups to get only those related to the CIT department by looking for groups with names containing "CIT -" or "CIT ".
3. Iterates through each filtered distribution group to retrieve its members.
4. Collects relevant information (Group Name, Group Email, Member Name, Member Email Address, and Recipient Type) about each member.
5. Exports the collected information to a specified CSV file.
#>

$CSVFilePath = "C:\Users\giovannicastillo\OneDrive - City of Torrance\Desktop\CSV Files\DL-Members.csv"

Try {
    Connect-ExchangeOnline -ShowBanner:$False

   $Result = @()
    # Filter to get only CIT department's distribution groups, specifically "CIT " and "CIT-"
    $DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {($_.Name -like "CIT -*" -or $_.Name -like "CIT *") -and $_.Name -notlike "*City*"}
    $GroupsCount = $DistributionGroups.Count
    $Counter = 1

    foreach ($Group in $DistributionGroups) {
        Write-Progress -Activity "Processing Distribution List: $($Group.DisplayName)" -Status "$Counter out of $GroupsCount completed" -PercentComplete (($Counter / $GroupsCount) * 100)

        $Members = Get-DistributionGroupMember -Identity $Group.Name -ResultSize Unlimited -ErrorAction SilentlyContinue
        foreach ($Member in $Members) {
            $Result += New-Object PSObject -property @{
                GroupName = $Group.Name
                GroupEmail = $Group.PrimarySmtpAddress
                Member = $Member.Name
                EmailAddress = $Member.PrimarySMTPAddress
                RecipientType = $Member.RecipientType
            }
        }
        $Counter++
    }

    $Result | Export-CSV $CSVFilePath -NoTypeInformation -Encoding UTF8
} Catch {
    Write-Host -f Red "Error: $($_.Exception.Message)"
}
