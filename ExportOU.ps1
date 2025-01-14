<#  - Exports OU computers in domain directory, exports csv to given users desktop
    - check for Active Directory module installed on Powershell
    - If module present, continue
    - If No module, import
    - Get AD Computers in requsted OU
    - Export to user desktop
#>
Write-Output "Please add your computer username (without email) to export to Desktop"
$UsernameforPath = Read-Host
$Path = "C:\Users\$UsernameforPath\Desktop"
cd $Path

$ModuleCheck = $false
do{ 
    if(Get-Module -Name ActiveDirectory){
    #continues if module is present
    $ModuleCheck = $true
    Write-Output "AD Module Present. "    
    }
    else{
    #installs AD module
    Install-Module -Name ActiveDirectory
    Import-Module ActiveDirectory
    Write-Output "AD Module absent. Installing and retrying."
    }
}until($ModuleCheck)

#after module install, get OU computers
Get-ADComputer -SearchBase 'OU=Computers,OU=HWP,DC=hollywoodparkca,DC=com'-SearchScope Subtree -Filter * | Select-Object Name | Export-Csv -Path ".\Computers_in_OUs.csv" -NoTypeInformation
Write-Host "CSV Saved to desktop with Computers in the 'hollywoodparkca.com/HWP/Computers' OU "
