using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TeamsAdminContent starter.."

$listItemId = $Request.Query.listItemId


if (-not $listItemId) {
    $body = "listItemId is required"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::NotAcceptable
        Body = $body
    })
    return
}
$clientId  = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminId' -AsPlainText
$clientSecret = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminSecret' -AsPlainText
$adminSiteUrl = "https://dimbaas.sharepoint.com/sites/aviadoas"
$listName = "Anbud"
    
#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
Set-PnPListItem -List Teams -Identity $listItemId  -Values @{"avStatus"="Kopierer innhold";} -Connection $adminConn
$templateSiteUrl = $listItem["avTemplateSiteUrl"]
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
Set-PnPListItem -List Teams -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn
$body = "Content copied - OK"
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
