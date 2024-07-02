# Install and Import required modules
Install-Module -Name AzureAD -AllowClobber -Force
Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
Install-Module -Name MSOnline -AllowClobber -Force

Import-Module AzureAD
Import-Module ExchangeOnlineManagement
Import-Module MSOnline

# Connect to Azure AD, Microsoft Online and Exchange
Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline

# Ask for the terminated user's UPN and the requesting person's UPN
$terminatedUserUPN = Read-Host -Prompt 'Enter the UPN of the terminated user'
$requestingUserUPN = Read-Host -Prompt 'Enter the UPN of the person who needs access to the mailbox'

# Restore the deleted user if necessary
$deletedUser = Get-MsolUser -ReturnDeletedUsers | Where-Object { $_.UserPrincipalName -eq $terminatedUserUPN }
if ($deletedUser) {
    Restore-MsolUser -UserPrincipalName $terminatedUserUPN
    Write-Host "Deleted Azure AD user $($terminatedUserUPN) has been restored."
    Start-Sleep -Seconds 45 # First 30-second delay
    # delay for sync
}



# Attempt mailbox operations with error handling
try {
    Set-Mailbox -Identity $terminatedUserUPN -Type Shared -ErrorAction Stop
    Add-MailboxPermission -Identity $terminatedUserUPN -User $requestingUserUPN -AccessRights FullAccess -ErrorAction Stop
    Write-Host "Mailbox for user $terminatedUserUPN has been converted to Shared. Permissions granted to $requestingUserUPN."

    # Permanently delete the Azure AD user account using MSOnline module
    Remove-MsolUser -UserPrincipalName $terminatedUserUPN -Force
    Write-Host "Azure AD user account for $terminatedUserUPN has been permanently deleted using MSOnline module."

    # Schedule disabling of the shared mailbox after 30 days
    $trigger = New-JobTrigger -Once -At (Get-Date).AddDays(30)
    Register-ScheduledJob -Name "DisableSharedMailbox" -Trigger $trigger -ScriptBlock {
        param($upn)
        Set-Mailbox -Identity $upn -Type Regular
    } -ArgumentList $terminatedUserUPN
    Write-Host "Shared mailbox for user $terminatedUserUPN is scheduled to be disabled on $((Get-Date).AddDays(30))"
} catch {
    Write-Host "An error occurred: $_"
}

