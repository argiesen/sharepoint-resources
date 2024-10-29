# Created by: CompuNet Inc
# Authors: Andy Giesen <agiesen@compunet.biz>
# Last Modified October 29, 2024

# PowerShell 7.x is required
# PnP module must be installed and up-to-date
# Install   : Install-Module -Name PnP.PowerShell
# Update    : Update-Module -Name PnP.PowerShell

# Config Parameters
$tenantAdminUrl     = "https://m365x97415188-admin.sharepoint.com"  # Change this to your admin center URL
$clientId           = "c00e64bc-4761-4c36-b358-3119fea350e5"        # Change this to the PnP module application ID
$errorLogFileName   = "RecycleBinEmptyError.log"

# Create a stopwatch instance to track the runtime
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Get all site collections (excluding personal OneDrive sites)
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId
$sites = Get-PnPTenantSite -Filter "Url -notlike '-my.sharepoint.com'"

foreach ($site in $sites){
    Write-Host "$($site.Title): Processing site: $($site.Url)" -ForegroundColor Blue
    
    # Display the current runtime
    $elapsedTime = $stopwatch.Elapsed
    Write-Host "$($site.Title): Current runtime: $($elapsedTime.Hours)h $($elapsedTime.Minutes)m $($elapsedTime.Seconds)s" -ForegroundColor DarkGray

    try {
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId

        # First-stage Recycle Bin: Empty all items
        $firstStageItems = Get-PnPRecycleBinItem -FirstStage
        if ($firstStageItems.Count -gt 0) {
            Write-Host "$($site.Title): Emptying $($firstStageItems.Count) items from first-stage recycle bin..." -ForegroundColor Yellow
            $firstStageItems | Clear-PnPRecycleBinItem -Force
        } else {
            Write-Host "$($site.Title): First-stage recycle bin is already empty." -ForegroundColor Green
        }

        # Second-stage Recycle Bin: Empty all items
        $secondStageItems = Get-PnPRecycleBinItem -SecondStage
        if ($secondStageItems.Count -gt 0) {
            Write-Host "$($site.Title): Emptying $($secondStageItems.Count) items second-stage recycle bin..." -ForegroundColor Yellow
            $secondStageItems | Clear-PnPRecycleBinItem -Force
        } else {
            Write-Host "$($site.Title): Second-stage recycle bin is already empty." -ForegroundColor Green
        }
        
        Write-Host "$($site.Title): Finished processing $($site.Url)" -ForegroundColor Green
    } catch {
        "{0} : {1} : {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $site.Url, $_.Exception.Message | Out-File $errorLogFileName -Append
        Write-Host "$($site.Title): Error emptying recycle bin: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Stop the stopwatch and display the total runtime
$stopwatch.Stop()
$finalTime = $stopwatch.Elapsed
Write-Host "Total runtime: $($finalTime.Hours)h $($finalTime.Minutes)m $($finalTime.Seconds)s" -ForegroundColor DarkGray

Write-Host "Logs have been saved to $errorLogFileName" -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline
