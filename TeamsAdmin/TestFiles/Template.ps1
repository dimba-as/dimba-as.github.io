Write-Host "TeamsAdminTemplate starter.."

$listItemId = 3

$clientId = "153767e6-5de3-4c6b-89a3-d7299fc3aeb4"
$clientSecret = "8Kv1qIE1vK.j.__WdZSqpuQ~2o.hApSMt0"
$adminSiteUrl = "https://norgesvel.sharepoint.com/sites/Prosjekt"
$libraryName = "Dokumenter"
$listName = "ProsjektoversiktProvisjonering"

#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Legger på mal";} -Connection $adminConn
$Title = $listItem["Title"]
$templateSiteUrl = $listItem["avMal"]
$spoUrl = $listItem["avSharePointUrl"]
#################################

#################################
#Templating
Write-Host "Get template from $templateSiteUrl" 
$templateConn = Connect-PnPOnline -url $templateSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$template = Get-PnPSiteTemplate -OutputInstance -Handlers Lists,Navigation,Features,PageContents,Pages -IncludeAllPages -Connection $templateConn
Write-Host "Apply template to $Title" 
$spoConn = Connect-PnPOnline -url $spoUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#$destContentSite = Get-PnPWeb -Connection $spoConn
Invoke-PnPSiteTemplate -InputInstance $template -ClearNavigation -Connection $spoConn
#################################

#################################
#Update list item
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn
#################################
Disconnect-PnPOnline -Connection $adminConn
Disconnect-PnPOnline -Connection $spoConn
$body = "Template applied - OK " 
# Associate values to output bindings by calling 'Push-OutputBinding'.
