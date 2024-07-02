# Define SharePoint Online Admin Center URL
$adminUrl = "https://torrancecagov-admin.sharepoint.com"
$siteUrl = "https://torrancecagov.sharepoint.com/sites/CIT_DEPT"

try {
    # Import certificate information for PnP PowerShell
    $cert = Import-Csv "C:\Scripts\SharePoint-CATV\PnPPowerShellIDs.csv"
    $AzureTenant = "torrancecagov.onmicrosoft.com"

    # Connect to the CIT_DEPT SharePoint site with certificate-based login
    Write-Host "Connecting to CIT_DEPT site..."
    Connect-PnPOnline -Url $siteUrl -ClientId $cert.ClientId -Tenant $AzureTenant -CertificateBase64Encoded $cert.Base64Encoded
} catch {
    Write-Host "Error connecting to CIT_DEPT site: $_"
    Read-Host -Prompt "Press Enter to exit"
    exit
}

try {
    # Name of the SharePoint list
    $listName = "DeptSitesList"

    # Check if the list exists, if not create it
    Write-Host "Checking if list exists..."
    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        Write-Host "List does not exist, creating list..."
        $list = New-PnPList -Title $listName -Template GenericList
        Add-PnPField -List $listName -DisplayName "Site URL" -InternalName "SiteURL" -Type URL -AddToDefaultView
        Add-PnPField -List $listName -DisplayName "Owner" -InternalName "SiteOwner" -Type Text -AddToDefaultView
        Write-Host "List and fields created."
    } else {
        Write-Host "List already exists."
    }
} catch {
    Write-Host "Error checking or creating list: $_"
    Read-Host -Prompt "Press Enter to exit"
    exit
}

$siteData = @()

try {
    # Connect to SharePoint Online Admin Center with certificate-based login
    Write-Host "Connecting to SharePoint Online Admin Center..."
    Connect-PnPOnline -Url $adminUrl -ClientId $cert.ClientId -Tenant $AzureTenant -CertificateBase64Encoded $cert.Base64Encoded

    # Retrieve all SharePoint Sites ending with "_DEPT"
    Write-Host "Retrieving DEPT sites..."
    $deptSites = Get-PnPTenantSite | Where-Object { $_.Url -like "*_DEPT" }

    # Check if any DEPT sites were found
    if ($deptSites.Count -eq 0) {
        Write-Host "No DEPT sites found!"
        Read-Host -Prompt "Press Enter to exit"
        exit
    } else {
        Write-Host "Found $($deptSites.Count) DEPT sites."
    }
} catch {
    Write-Host "Error retrieving DEPT sites: $_"
    Read-Host -Prompt "Press Enter to exit"
    exit
}

foreach ($site in $deptSites) {
    try {
        # Connect to each individual DEPT site with certificate-based login
        Write-Host "Connecting to site: $($site.Url)"
        Connect-PnPOnline -Url $site.Url -ClientId $cert.ClientId -Tenant $AzureTenant -CertificateBase64Encoded $cert.Base64Encoded

        # Retrieve the Owners group and its members
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
        if ($ownersGroup) {
            $owners = Get-PnPGroupMember -Identity $ownersGroup.Id
            foreach ($owner in $owners) {
                $siteData += [PSCustomObject]@{
                    "SiteUrl" = $site.Url
                    "OwnerEmail" = $owner.Email
                }
            }
        } else {
            Write-Host "Owners group not found for site: $($site.Url)"
        }
    } catch {
        Write-Host "Error processing site $($site.Url): $_"
    } finally {
        # Disconnect from the site to manage connections
        Disconnect-PnPOnline
    }
}

try {
    # Update the list with the site data
    Write-Host "Updating the list with site data..."
    foreach ($data in $siteData) {
        # Connect to the CIT_DEPT site with certificate-based login to manage the list
        Connect-PnPOnline -Url $siteUrl -ClientId $cert.ClientId -Tenant $AzureTenant -CertificateBase64Encoded $cert.Base64Encoded

        # Add or update list items as needed
        Write-Host "Adding item for site: $($data.SiteUrl) with owner: $($data.OwnerEmail)"
        $addItemResult = Add-PnPListItem -List $listName -Values @{
            "SiteURL" = $data.SiteUrl
            "SiteOwner" = $data.OwnerEmail
        }
        if ($addItemResult) {
            Write-Host "Item added with ID: $($addItemResult.Id)"
        } else {
            Write-Host "Failed to add item."
        }

        # Disconnect from the site after updating the list
        Disconnect-PnPOnline
    }

    # Print the collected site data in the terminal
    Write-Host "Collected site data:"
    $siteData | Format-Table -Property SiteUrl, OwnerEmail -AutoSize
} catch {
    Write-Host "Error updating the list: $_"
} finally {
    # Ensure the script pauses at the end
    Read-Host -Prompt "Press Enter to exit"
}
