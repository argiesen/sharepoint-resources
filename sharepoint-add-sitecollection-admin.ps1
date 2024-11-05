#Parameters
$TenantAdminURL = "https://<tenant>-admin.sharepoint.com"
$clientId       = "c00e64bc-4761-4c36-b358-3119fea350e5"    
$SiteCollAdmin  = "user@domain.com"

#Connect to Admin Center
Connect-PnPOnline -Url $TenantAdminURL -Interactive -ClientId

#Get All Site collections and Iterate through
$SiteCollections = Import-Csv Sites.csv
foreach($Site in $SiteCollections){ 
    #Add Site collection Admin
    Set-PnPTenantSite -Url $Site.Url -Owners $SiteCollAdmin
    Write-host "Added Site Collection Administrator to $($Site.URL)"
}

