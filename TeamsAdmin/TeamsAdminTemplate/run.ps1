using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TeamsAdminTemplate starter.."

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
$libraryName = "Dokumenter"
$listName = "Anbud"
#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Legger på mal";} -Connection $adminConn
$Title = $listItem["Title"]
$templateSiteUrl = $listItem["avTemplateSiteUrl"]
$spoUrl = $listItem["avSharePointUrl"]
#################################

#################################
#Templating
Write-Host "Get template from $templateSiteUrl" 
$templateConn = Connect-PnPOnline -url $templateSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$template = Get-PnPSiteTemplate -OutputInstance -Handlers Lists,Navigation,PageContents,Pages -IncludeAllPages -ListsToExtract $libraryName -Connection $templateConn
Write-Host "Apply template to $Title" 
$spoConn = Connect-PnPOnline -url $spoUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
Invoke-PnPSiteTemplate -InputInstance $template -ClearNavigation -Connection $spoConn
#################################

#################################
#Update list item
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn
#################################

$body = "Template applied - OK"
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
