# PnP module must be installed
# Config Parameters
$SiteURL = "https://m365x29955680.sharepoint.com/sites/ContosoBrand"
$VersionsToKeep = 5 # Does not include current version

Try {
    # Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive
 
    # Get the Context
    $context = Get-PnPContext
 
    # Exclude certain libraries
    $excludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images",
                            "Site Collection Documents", "Site Collection Images","Style Library")
 
    # Get All document libraries
    $documentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Title -notin $excludedLists -and $_.Hidden -eq $false}
 
    # Iterate through each document library
    ForEach($documentLibrary in $documentLibraries){
        Write-Host "Processing Document Library:"$documentLibrary.Title -f Magenta
 
        # Get All Items from the List - Exclude 'Folder' List Items
        $listItems = Get-PnPListItem -List $documentLibrary -PageSize 2000 | Where-Object {$_.FileSystemObjectType -eq "File"}
 
        # Loop through each file
        ForEach ($item in $listItems){
            # Get File Versions
            $file = $item.File
            $versions = $file.Versions
            $context.Load($file)
            $context.Load($versions)
            $context.ExecuteQuery()
  
            Write-Host -f Yellow "`tScanning File:"$file.Name
            $versionsCount = $versions.Count
            $versionsToDelete = $versionsCount - $VersionsToKeep
            If($versionsToDelete -gt 0){
                Write-Host -f Cyan "`t Total Number of Versions of the File:" $versionsCount
                $versionCounter = 0

                # Delete versions
                For($i=0; $i -lt $versionsToDelete; $i++){
                    If($versions[$versionCounter].IsCurrentVersion){
                       $versionCounter++
                       Write-Host -f Magenta "`t`t Retaining Current Major Version:"$versions[$versionCounter].VersionLabel
                       Continue
                    }
                    Write-Host -f Cyan "`t Deleting Version:" $versions[$versionCounter].VersionLabel
                    $versions[$versionCounter].DeleteObject()
                    "{0} : {1} : {2}/{3} (v{4})" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $SiteURL, $documentLibrary.Title, $file.Name, $versions[$versionCounter].VersionLabel | Out-File VersionDeletion.log -Append
                }

                $context.ExecuteQuery()
                Write-Host -f Green "`t Version History is cleaned for the File:"$file.Name
            }
        }
    }
} Catch {
    Write-Host -f Red "Error Cleaning up Version History!" $_.Exception.Message
}