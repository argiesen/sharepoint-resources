# Import the PnP PowerShell module
Import-Module PnP.PowerShell

# Connect to SharePoint Online
$tenantAdminUrl = "M365x97415188.sharepoint.com"
$clientId = "c00e64bc-4761-4c36-b358-3119fea350e5"
Connect-PnPOnline -Url $tenantAdminUrl -Interactive -ClientId $clientId

# Get all site collections
$siteCollections = Get-PnPTenantSite

# Initialize an array to store the results
$results = @()

# Exclude certain libraries
$excludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images", "Site Collection Documents", "Site Collection Images","Style Library")

# Iterate through each site collection
foreach ($site in $siteCollections){
    Write-Host "Processing site $($site.Title)" -ForegroundColor Yellow

    # Set document library counter to 0
    $i = 0

    try {
        # Connect to the site collection
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientId

        # Get the document libraries
        $documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -notin $excludedLists -and $_.Hidden -eq $false } # BaseTemplate 101 is for document libraries

        foreach ($documentLibrary in $documentLibraries){
            $i++
            $Error.Clear()

            Write-Host "$i/$($documentLibraries.Count): Processing document library $($site.Title)/$($documentLibrary.Title)... " -ForegroundColor Yellow -NoNewline

            # Get all files in the document library
            $files = Get-PnPListItem -List $documentLibrary.Title | Where-Object { $_.FileSystemObjectType -eq "File" }

            foreach ($file in $files){
                # Get the version count for each file
                $fileVersions = Get-PnPFileVersion -Url $file.FieldValues.FileRef
                
                # Store the result
                $results += [PSCustomObject]@{
                    SiteUrl         = $site.Url
                    LibraryName     = $documentLibrary.Title
                    FileName        = $file.FieldValues.FileLeafRef
                    FileUrl         = $file.FieldValues.FileRef
                    VersionCount    = $fileVersions.Count
                }
            }

            if ($Error){
                Write-Host "Error: $($Error.Exception.Message)" -ForegroundColor Red
            }else{
                Write-Host "Completed" -ForegroundColor Green
            }
        }
    }catch{
        Write-Host "Error accessing site: $($site.Url) - $_"
    }
}

# Export results to CSV
$outputFile = "FileVersionReport-$(Get-Date -Format yyyy.MM.dd).csv"
$results | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "Output saved to $outputFile" -ForegroundColor Green

# Disconnect from SharePoint Online
Disconnect-PnPOnline
