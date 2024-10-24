# Created by: CompuNet Inc
# Authors: Andy Giesen <agiesen@compunet.biz>
# Last Modified October 23, 2024

# PowerShell 7.x is required
# PnP module must be installed and up-to-date
# Install   : Install-Module -Name PnP.PowerShell
# Update    : Update-Module -Name PnP.PowerShell

# Config Parameters
$tenantAdminUrl = "https://m365x97415188-admin.sharepoint.com"  # Change this to your admin center URL
$clientId = "c00e64bc-4761-4c36-b358-3119fea350e5"              # Change this to the PnP module application ID
$enableVersioning = $false                                      # Enable versioning if not enabled
$maxVersions = 100                                              # Maximum number of major versions
$minorVersionsLimit = 10                                        # Maximum number of minor versions

# Create a stopwatch instance to track the runtime
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Get all site collections
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId
$siteCollections = Get-PnPTenantSite

# Iterate through site collections
foreach ($site in $siteCollections){
    Write-Host "Processing Site: $($site.Title) ($($site.Url))" -ForegroundColor Blue

    # Set document library counter to 0
    $i = 0

    try{
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId
        
        # Get document libraries
        $documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -notin $excludedLists -and $_.Hidden -eq $false } # BaseTemplate 101 is for document libraries
        
        # Iterate through document libraries
        foreach ($documentLibrary in $documentLibraries){
            # Display the current runtime
            $elapsedTime = $stopwatch.Elapsed
            Write-Host "Current Runtime: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s" -ForegroundColor DarkGray

            $i++
            Write-Host "$i/$($documentLibraries.Count): Processing Document Library: $($site.Title)/$($documentLibrary.Title)" -ForegroundColor Magenta

            if ($enableVersioning){
                # Enable versioning if not already enabled
                if (-not $list.EnableVersioning) {
                    Write-Host "Enabling versioning for list: $($list.Title)"
                    Set-PnPList -Identity $list.Title -EnableVersioning $true
                }
            }

            # Check if the document library supports versioning
            if ($documentLibrary.EnableVersioning) {
                Write-Host "Updating versioning settings for document library: $($documentLibrary.Title)" -ForegroundColor Yellow

                # Set the major and minor version limits
                Set-PnPList -Identity $documentLibrary.Title `
                    -MajorVersions $maxVersions `
                    -MinorVersions $minorVersionsLimit
            } else {
                Write-Host "Skipping document library: Versioning not enabled on $($site.Title)/$($documentLibrary.Title)" -ForegroundColor Red
            }
        }
    }catch{
        Write-Host "Error accessing site: $($site.Url) - $_" -ForegroundColor Red
    }
}

# Stop the stopwatch and display the total runtime
$stopwatch.Stop()
$finalTime = $stopwatch.Elapsed
Write-Host "Total Runtime: $($finalTime.Hours)h $($finalTime.Minutes)m $($finalTime.Seconds)s" -ForegroundColor DarkGray

# Disconnect from SharePoint Online
Disconnect-PnPOnline