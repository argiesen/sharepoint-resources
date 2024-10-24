# Created by: CompuNet Inc
# Authors: Andy Giesen <agiesen@compunet.biz>
# Last Modified October 23, 2024

# PowerShell 7.x is required
# PnP module must be installed and up-to-date
# Install   : Install-Module -Name PnP.PowerShell
# Update    : Update-Module -Name PnP.PowerShell

# Config Parameters
$tenantAdminUrl = "https://m365x97415188-admin.sharepoint.com"  # Change this to your admin center URL
$clientId       = "c00e64bc-4761-4c36-b358-3119fea350e5"        # Change this to the PnP module application ID
$VersionsToKeep = 5                                             # Change to desired amount. Does not include current version

# Create a stopwatch instance to track the runtime
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Get all site collections
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId
$siteCollections = Get-PnPTenantSite | `
    Where-Object Url -notmatch "-my.|/appcatalog|/contenttypehub|/search|/personal/|/my/" | `   # Exclude system and OneDrive sites
    Where-Object Template -notmatch "GROUP#0|STS#-1"                                            # Exclude Teams group sites and system templates

# Exclude certain libraries
$excludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images", 
    "Site Collection Documents", "Site Collection Images","Style Library")

# Iterate through site collections
foreach ($site in $siteCollections){
    Write-Host "Processing Site: $($site.Title) ($($site.Url))" -ForegroundColor Blue

    # Set document library counter to 0
    $i = 0

    try{
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId
    
        # Get the context to collect file versions
        $context = Get-PnPContext
    
        # Get document libraries
        $documentLibraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Title -notin $excludedLists -and $_.Hidden -eq $false }
    
        # Iterate through document libraries
        foreach ($documentLibrary in $documentLibraries){
            # Display the current runtime
            $elapsedTime = $stopwatch.Elapsed
            Write-Host "Current Runtime: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s" -ForegroundColor DarkGray

            $i++
            Write-Host "$i/$($documentLibraries.Count): Processing Document Library: $($site.Title)/$($documentLibrary.Title)" -ForegroundColor Magenta
    
            # Get All Items from the List - Exclude 'Folder' List Items
            $listItems = Get-PnPListItem -List $documentLibrary -PageSize 5000 | Where-Object { $_.FileSystemObjectType -eq "File" }
    
            # Loop through each file
            foreach ($item in $listItems){
                # Get File Versions
                $file = $item.File
                $versions = $file.Versions
                $context.Load($file)
                $context.Load($versions)
                $context.ExecuteQuery()
    
                Write-Host "`tScanning File: $($file.Name)" -ForegroundColor Yellow
                $versionsCount = $versions.Count
                $versionsToDelete = $versionsCount - $VersionsToKeep
                If($versionsToDelete -gt 0){
                    Write-Host "`t Total Number of Versions of the File: $versionsCount" -ForegroundColor Cyan
                    $versionCounter = 0

                    # Delete versions
                    For($i=0; $i -lt $versionsToDelete; $i++){
                        If($versions[$versionCounter].IsCurrentVersion){
                        $versionCounter++
                        Write-Host "`t`t Retaining Current Major Version: $($versions[$versionCounter].VersionLabel)" -ForegroundColor Magenta
                        Continue
                        }
                        Write-Host "`t Deleting Version: $($versions[$versionCounter].VersionLabel)"  -ForegroundColor Cyan
                        $versions[$versionCounter].DeleteObject()
                        "{0} : {1} : {2}/{3} (v{4})" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel | Out-File VersionDeletion.log -Append
                    }

                    $context.ExecuteQuery()
                    Write-Host "`t Version History is cleaned for the File: $($file.Name)" -ForegroundColor Green
                }
            }
        }
    }catch{
        "{0} : {1} : {2}/{3} (v{4})" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel, $_.Exception.Message | Out-File VersionDeletionError.log -Append
        Write-Host "Error Cleaning up Version History: $($_.Exception.Message)"  -ForegroundColor Red
    }
}

# Stop the stopwatch and display the total runtime
$stopwatch.Stop()
$finalTime = $stopwatch.Elapsed
Write-Host "Total Runtime: $($finalTime.Hours)h $($finalTime.Minutes)m $($finalTime.Seconds)s" -ForegroundColor DarkGray

Write-Host "Deletion logs have been saved to VersionDeletion.log" -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline
