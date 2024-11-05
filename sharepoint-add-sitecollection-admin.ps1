#Parameters
$TenantAdminURL = "https://m365x97415188-admin.sharepoint.com"
$clientId       = "c00e64bc-4761-4c36-b358-3119fea350e5"    
$SiteCollAdmin  = "fry@M365x97415188.onmicrosoft.com"

$CsvPath        = "Sites.csv"

#Connect to Admin Center
Connect-PnPOnline -Url $TenantAdminURL -Interactive -ClientId $clientId

#Get All Site collections and Iterate through
$SiteCollections = Import-Csv $CsvPath
#$SiteCollections = Get-PnPTenantSite
foreach($Site in $SiteCollections){ 
    #Add Site collection Admin
    Set-PnPTenantSite -Url $Site.Url -Owners $SiteCollAdmin
    Write-host "Added Site Collection Administrator to $($Site.URL)"
}

