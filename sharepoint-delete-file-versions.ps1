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

# Retry parameters
$maxRetryAttempts = 3
$retryDelaySeconds = 3

# Create a stopwatch instance to track the runtime
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Get all site collections
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId
$siteCollections = Get-PnPTenantSite

# Exclude certain libraries
$excludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images", "Site Collection Documents", "Site Collection Images","Style Library", "User Photos")

# Iterate through site collections
foreach ($site in $siteCollections){
    Write-Host "Processing site: $($site.Title) ($($site.Url))" -ForegroundColor Blue

    # Set document library counter to 0
    $i = 0

    try {
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
            Write-Host "Current runtime: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s" -ForegroundColor DarkGray

            $i++
            Write-Host "$i/$($documentLibraries.Count): Processing document library: $($site.Title)/$($documentLibrary.Title)" -ForegroundColor Magenta
    
            # Get All Items from the List - Exclude 'Folder' List Items
            $listItems = Get-PnPListItem -List $documentLibrary -PageSize 5000 | Where-Object { $_.FileSystemObjectType -eq "File" }
    
            # Loop through each file
            foreach ($item in $listItems){
                # Get File Versions
                $file = $item.File
                $versions = $file.Versions
                $context.Load($file)
                $context.Load($versions)
                #$context.ExecuteQuery()

                $retryAttempts = 0
                $success = $false

                while (-not $success -and $retryAttempts -lt $maxRetryAttempts){
                    try {
                        $context.ExecuteQuery()
                        $success = $true
                    } catch {
                        $retryAttempts++
                        if ($retryAttempts -lt $maxRetryAttempts) {
                            Write-Host "Error encountered (Attempt $retryAttempts/$maxRetryAttempts). Retrying in $retryDelaySeconds seconds..." -ForegroundColor Yellow
                            Start-Sleep -Seconds $retryDelaySeconds
                        } else {
                            Write-Host "Max retry attempts reached. Error: $($_.Exception.Message)" -ForegroundColor Red
                            "{0} : {1} : {2}/{3} (v{4}) : {5}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel, $_.Exception.Message | Out-File VersionDeletionError.log -Append
                        }
                    }
                }

                if ($success){
                    Write-Host "`tScanning file: $($file.Name)" -ForegroundColor Yellow
                    $versionsCount = $versions.Count
                    $versionsToDelete = $versionsCount - $VersionsToKeep
                    if ($versionsToDelete -gt 0){
                        Write-Host "`tTotal number of versions of the file: $versionsCount" -ForegroundColor Cyan
                        $versionCounter = 0

                        # Delete versions
                        for ($i=0; $i -lt $versionsToDelete; $i++){
                            if ($versions[$versionCounter].IsCurrentVersion){
                                $versionCounter++
                                Write-Host "`t`tRetaining current major version: $($versions[$versionCounter].VersionLabel)" -ForegroundColor Magenta
                                Continue
                            }
                            Write-Host "`tDeleting version: $($versions[$versionCounter].VersionLabel)"  -ForegroundColor Cyan
                            $versions[$versionCounter].DeleteObject()
                            "{0} : {1} : {2}/{3} (v{4})" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel | Out-File VersionDeletion.log -Append
                        }

                        #$context.ExecuteQuery()
                        #Write-Host "`tVersion history is cleaned for the file: $($file.Name)" -ForegroundColor Green

                        # Retry ExecuteQuery for version deletions
                        $retryAttempts = 0
                        $success = $false

                        while (-not $success -and $retryAttempts -lt $maxRetryAttempts){
                            try {
                                $context.ExecuteQuery()
                                $success = $true
                            } catch {
                                $retryAttempts++
                                if ($retryAttempts -lt $maxRetryAttempts) {
                                    Write-Host "Error encountered while deleting versions (Attempt $retryAttempts/$maxRetryAttempts). Retrying in $retryDelaySeconds seconds..." -ForegroundColor Yellow
                                    Start-Sleep -Seconds $retryDelaySeconds
                                } else {
                                    Write-Host "Max retry attempts reached while deleting versions. Error: $($_.Exception.Message)" -ForegroundColor Red
                                    "{0} : {1} : {2}/{3} : ERROR : {4}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $_.Exception.Message | Out-File VersionDeletionError.log -Append
                                }
                            }
                        }

                        if ($success){
                            Write-Host "`t Version History is cleaned for the File: $($file.Name)" -ForegroundColor Green
                        }
                    }
                }
            }
        }
    } catch {
        "{0} : {1} : {2}/{3} (v{4}) : {5}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel, $_.Exception.Message | Out-File VersionDeletionError.log -Append
        Write-Host "Error cleaning up version history: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Stop the stopwatch and display the total runtime
$stopwatch.Stop()
$finalTime = $stopwatch.Elapsed
Write-Host "Total runtime: $($finalTime.Hours)h $($finalTime.Minutes)m $($finalTime.Seconds)s" -ForegroundColor DarkGray

Write-Host "Deletion logs have been saved to VersionDeletion.log" -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline
