

$clientId = "ac4e43f9-520c-427d-8349-24cbb9fa8ff5"
$clientSecret = "~Pilzjy.gCmci-wd6BuI8P6.P.5t28mh-o"

$sourceConn = Connect-PnPOnline -Url https://dimbaas.sharepoint.com/sites/aviado.no77 -Interactive #-ClientSecret $clientSecret -ClientId $clientId -ReturnConnection

Get-PnpWeb -Connection $sourceConn

#################################
#Templating
Write-Host "Get template from $templateSiteUrl" 
$template = Get-PnPSiteTemplate -OutputInstance -Handlers Lists,Navigation,PageContents,Pages -IncludeAllPages -Connection $sourceConn

$destConn = Connect-PnPOnline -Url https://dimbaas.sharepoint.com/sites/elevateit.no2  -Interactive -ReturnConnection #-ClientSecret $clientSecret -ClientId $clientId 
Get-PnPWeb -Connection $destConn
Write-Host "Apply template to $Title" 
Invoke-PnPSiteTemplate -InputInstance $template -ClearNavigation -Connection $destConn
#################################

Disconnect-PnPOnline