# Import the PnP PowerShell module
Import-Module PnP.PowerShell

# Connect to SharePoint Online
$tenantAdminUrl = "https://m365x29955680.sharepoint.com"
Connect-PnPOnline -Url $tenantAdminUrl -Interactive

# Get all site collections
$siteCollections = Get-PnPTenantSite

# Initialize an array to store the results
$results = @()

# Exclude certain libraries
$excludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images", "Site Collection Documents", "Site Collection Images","Style Library")

# Iterate through each site collection
foreach ($site in $siteCollections) {
    try {
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -Interactive
        
        # Get the document library settings
        $documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -notin $excludedLists -and $_.Hidden -eq $false } # BaseTemplate 101 is for document libraries
        
        foreach ($documentLibrary in $documentLibraries) {
            # Store the successful result
            $results += [PSCustomObject]@{
                SiteTitle                       = $site.Title
                SiteUrl                         = $site.Url
                LibraryName                     = $documentLibrary.Title
                EnableVersioning                = $documentLibrary.EnableVersioning
                EnableMinorVersions             = $documentLibrary.EnableMinorVersions
                MajorVersionLimit               = $documentLibrary.MajorVersionLimit
                MajorWithMinorVersionsLimit     = $documentLibrary.MajorWithMinorVersionsLimit
                Notes                           = $null
            }
        }
    } catch {
        # Store the failed result
        $results += [PSCustomObject]@{
            SiteTitle                       = $site.Title
            SiteUrl                         = $site.Url
            LibraryName                     = $null
            EnableVersioning                = $null
            EnableMinorVersions             = $null
            MajorVersionLimit               = $null
            MajorWithMinorVersionsLimit     = $null
            Notes                           = "$_"
        }

        Write-Host "Error accessing site: $($site.Url) - $_"
    }
}

# Display the results
$results | Format-Table -AutoSize

# Optionally, export results to CSV
$results | Export-Csv -Path "VersionLimitsReport.csv" -NoTypeInformation

# Disconnect from SharePoint Online
Disconnect-PnPOnline
