# Takes Organization CSV file and connects to exchange online and filters through site list for only ones containing CIT in their url
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
