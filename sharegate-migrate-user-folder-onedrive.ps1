# https://documentation.sharegate.com/hc/en-us/articles/115000640548-Connect-Site

Import-Module Sharegate

#Get All OneDrive sites from CSV
$OneDriveSites = Import-Csv "C:\MigrationRuns\PilotGroupUsers.csv"

#Get Credentials to connect
$Credentials = Get-Credential
#$dstsiteConnection = Connect-Site -Url https://<tenanturl>-admin.sharepoint.com/ -Browser

$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
$propertyTemplate = New-PropertyTemplate -AuthorsAndTimestamps -VersionHistory -VersionLimit 5 -CheckInAs Publish

Set-Variable dstSite, dstList
foreach ($Site in $OneDriveSites) {
    Clear-Variable dstSite
    Clear-Variable dstList
    $dstSite = Connect-Site -Url $Site.OneDriveURL -Credential $Credentials
    $dstList = Get-List -Name Documents -Site $dstSite
    Import-Document -SourceFolder "$($Site.Path)\$($Site.Directory)" -DestinationList $dstList -DestinationFolder $Site.OneDriveFolder -CopySettings $copysettings -Template $propertyTemplate
}
