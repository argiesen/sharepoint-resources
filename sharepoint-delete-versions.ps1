# PnP module must be installed
# Config Parameters
$SiteURL = "https://m365x29955680.sharepoint.com/sites/ContosoBrand"
$VersionsToKeep = 5 # Does not include current version

Try {
    # Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive
 
    # Get the Context
    $Ctx = Get-PnPContext
 
    # Exclude certain libraries
    $ExcludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images",
                            "Site Collection Documents", "Site Collection Images","Style Library")
 
    # Get All document libraries
    $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Title -notin $ExcludedLists -and $_.Hidden -eq $false}
 
    # Iterate through each document library
    ForEach($Library in $DocumentLibraries){
        Write-Host "Processing Document Library:"$Library.Title -f Magenta
 
        # Get All Items from the List - Exclude 'Folder' List Items
        $ListItems = Get-PnPListItem -List $Library -PageSize 2000 | Where-Object {$_.FileSystemObjectType -eq "File"}
 
        # Loop through each file
        ForEach ($Item in $ListItems){
            # Get File Versions
            $File = $Item.File
            $Versions = $File.Versions
            $Ctx.Load($File)
            $Ctx.Load($Versions)
            $Ctx.ExecuteQuery()
  
            Write-Host -f Yellow "`tScanning File:"$File.Name
            $VersionsCount = $Versions.Count
            $VersionsToDelete = $VersionsCount - $VersionsToKeep
            If($VersionsToDelete -gt 0){
                Write-Host -f Cyan "`t Total Number of Versions of the File:" $VersionsCount
                $VersionCounter = 0

                # Delete versions
                For($i=0; $i -lt $VersionsToDelete; $i++){
                    If($Versions[$VersionCounter].IsCurrentVersion){
                       $VersionCounter++
                       Write-Host -f Magenta "`t`t Retaining Current Major Version:"$Versions[$VersionCounter].VersionLabel
                       Continue
                    }
                    Write-Host -f Cyan "`t Deleting Version:" $Versions[$VersionCounter].VersionLabel
                    $Versions[$VersionCounter].DeleteObject()
                    "{0} : {1} : {2}/{3} (v{4})" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $SiteURL, $Library.Title, $File.Name, $Versions[$VersionCounter].VersionLabel | Out-File VersionDeletion.log -Append
                }

                $Ctx.ExecuteQuery()
                Write-Host -f Green "`t Version History is cleaned for the File:"$File.Name
            }
        }
    }
} Catch {
    Write-Host -f Red "Error Cleaning up Version History!" $_.Exception.Message
}