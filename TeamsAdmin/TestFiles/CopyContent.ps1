
# Write to the Azure Functions log stream.
Write-Host "TeamsAdminContent starter.."

$listItemId = 5
$clientId = "153767e6-5de3-4c6b-89a3-d7299fc3aeb4"
$clientSecret = "8Kv1qIE1vK.j.__WdZSqpuQ~2o.hApSMt0"
$adminSiteUrl = "https://norgesvel.sharepoint.com/sites/Prosjekt"
$listName = "ProsjektoversiktProvisjonering"
#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus" = "Kopierer innhold"; } -Connection $adminConn
$templateSiteUrl = $listItem["avMal"]
$spoUrl = $listItem["avSharePointUrl"]
#################################

#################################
#Copy content
Write-Host "Copy content" 
$templateConn = Connect-PnPOnline -url $templateSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$sourceContentSite = Get-PnPWeb -Connection $templateConn 
$sourceContentSiteRelUrl = $sourceContentSite.ServerRelativeUrl
$spoConn = Connect-PnPOnline -url $spoUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$destContentSite = Get-PnPWeb -Connection $spoConn 
$destContentSiteRelUrl = $destContentSite.ServerRelativeUrl 
Copy-PnPFile -SourceUrl "$sourceContentSiteRelUrl/Delte%20dokumenter/General" -TargetUrl "$destContentSiteRelUrl/Delte%20dokumenter" -Overwrite -Force
#################################
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus" = "Opprettet"; } -Connection $adminConn

Disconnect-PnPOnline -Connection $adminConn
Disconnect-PnPOnline -Connection $spoConn
$body = "Content copied - OK"
