$AdminCenterURL = "https://<tenanturl>-admin.sharepoint.com"
$FolderName = "Home Drive Content"

#Get All OneDrive sites from CSV
$OneDriveSites = Import-Csv "C:\MigrationRuns\PilotGroupUsers.csv"

#Get Credentials to connect
$Credentials = Get-Credential

Try {
    #Connect to Admin Center
    Connect-PnPOnline -Url $AdminCenterURL -Credential $Credentials

    #Iterate through Each OneDrive
    foreach($Site in $OneDriveSites){  
        Try {
            Write-Host -f Yellow "Ensuring Folder $FolderName in $($Site.OneDriveURL)" -NoNewline
            
            #Connect to OneDrive site
            Connect-PnPOnline -Url $Site.OneDriveURL -Credential $Credentials -ErrorAction Stop
            
            #Ensure folder in SharePoint Online using powershell
            $NewFolder = Resolve-PnPFolder -SiteRelativePath "Documents/$FolderName" -ErrorAction Stop
            Write-Host -f Green " Done!"
        } Catch {
            Write-Host "`tError: $($_.Exception.Message)" -Foregroundcolor Red
        }
    }
} Catch {
    Write-Host "Error: $($_.Exception.Message)" -Foregroundcolor Red
}
