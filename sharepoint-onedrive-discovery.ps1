$EntraList = Import-Csv "exportUsers.csv"
$File1 = Import-Csv "UserHomeDirectories.csv"
$File2 = Import-Csv "HomeDriveOneDriveMap.csv"
$Output = @()

# FolderPath = 1: Directory, 2: FolderName

#File1
foreach ($record in $File2){
    $FolderPath = "\\san\H$\" + $record.FolderName
    $Output += $record | Select-Object @{l='FullName';e={$_.FULLNAME}},@{l='JobTitle';e={$_.JOBTITLE}},@{l='Department';e={$_.Department}},@{l='OneDriveURL';e={$_.OneDriveURL}},@{l='FolderPath';e={$FolderPath}}
}

#File2
foreach ($record in $File1){
    if ($record.FullName -eq ""){
        $Output += $record | Select-Object @{l='FullName';e={$_.FULLNAME}},@{l='JobTitle';e={$_.JOBTITLE}},@{l='Department';e={$_.Department}},@{l='OneDriveURL';e={$_.OneDriveURL}},@{l='FolderPath';e={$_.Directory}}
    }elseif(!($File2.FullName -contains $record.FullName)){
        $Output += $record | Select-Object @{l='FullName';e={$_.FULLNAME}},@{l='JobTitle';e={$_.JOBTITLE}},@{l='Department';e={$_.Department}},@{l='OneDriveURL';e={$_.OneDriveURL}},@{l='FolderPath';e={$_.Directory}}
    }
}

foreach ($record in $Output | Where-Object FullName -ne ""){
    $match = $null
    $match = $EntraList | Where-Object displayName -eq $record.FullName
    $record.JobTitle = $match.JobTitle
    $record.Department = $match.Department
}

# Get additional info from AD
$Output = Import-Csv Output.csv
foreach ($record in $Output | Where-Object FullName -eq ""){
    $ADOutput = $null
    $ADOutput = Get-AdUser -Identity (Split-Path $record.FolderPath -Leaf) -Properties * | Select-Object displayName,title,department
    if ($ADOutput){
        $record.FullName = $ADOutput.displayName
        $record.JobTitle = $ADOutput.title
        $record.Department = $ADOutput.department
    }
}

$Output | Export-Csv -NoTypeInformation OutputFinal.csv


# Get additional info from AD
$Output = Import-Csv "Home Directory Paths for Active Users.csv"
foreach ($record in $Output){
    $ADOutput = $null
    $ADOutput = Get-AdUser -Identity $record.User -Properties * | Select-Object title,department
    if ($ADOutput){
        $record.JobTitle = $ADOutput.title
        $record.Department = $ADOutput.department
    }
}

$Output | Export-Csv -NoTypeInformation OutputFinal.csv



# Overlake
# Get additional info from AD
$Output = Import-Csv File01_FolderSizing.csv
foreach ($record in $Output | Where-Object DisplayName -eq ""){
    $adOutput = $null
    $adOutput = Get-AdUser -Identity (Split-Path $record.UserFolderPath -Leaf) -Properties * | Select-Object displayName,samAccountName,userPrincipalName,title,department
    if ($adOutput){
        $record.DisplayName = $adOutput.displayName
        $record.SamAccountName = $adOutput.samAccountName
        $record.UserPrincipalName = $adOutput.userPrincipalName
        $record.JobTitle = $adOutput.title
        $record.Department = $adOutput.department
    }
}


# DisplayName,SAMAccountName,UserPrincipalName,JobTitle,Department,OneDriveURL,UserFolderPath,UserFolderSizeGB,OneDriveFolder


Connect-PnPOnline overlakehospital.sharepoint.com -Interactive

$Output = Import-Csv File01_FolderSizing_01.csv
foreach ($record in $Output | Where-Object userPrincipalName -ne ""){
    $record.OneDriveURL = (Get-PnPUserProfileProperty -Account $record.userPrincipalName).PersonalUrl
}


# Resume
foreach ($record in ($Output | where-object {($_.UserPrincipalName -ne "" -and $_.UserPrincipalName -ne $null) -and $_.OneDriveURL -eq ""})){
    $record.OneDriveURL = (Get-PnPUserProfileProperty -Account $record.userPrincipalName).PersonalUrl
}
