using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TeamsAdminDefaultMetadata starter.."

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
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Legger på default metadata";} -Connection $adminConn
$Title = $listItem["Title"]
$templateSiteUrl = $listItem["avTemplateSiteUrl"]
$destinationSPUrl = $listItem["avSharePointUrl"]
#################################

#################################
#Set metadata
$templateConn = Connect-PnPOnline -url $templateSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$templateLibrary = Get-PnpList -Identity $libraryName -Connection $templateConn

$defaultValues = Get-PnPDefaultColumnValues -List $templateLibrary -Connection $templateConn

$destConn = Connect-PnPOnline -url $destinationSPUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection

$destLibrary = Get-PnpList -Identity $libraryName -Connection $destConn

foreach($defaultVal in $defaultValues)
{
    $val = $defaultVal.Value
    if($val -like "{*}")
    {
        if($val -eq "{SiteTitle}")
        {
            try
            {
                Set-PnPDefaultColumnValues -List $destLibrary -Field $defaultVal.Field -Value $Title -Folder $defaultVal.Path -Connection $destConn 
            }
            catch{}
        } 
    }
    else{
        try
        {
            Set-PnPDefaultColumnValues -List $destLibrary -Field $defaultVal.Field -Value $val -Folder $defaultVal.Path -Connection $destConn 
        }
        catch{}
    }
    
}
#################################
#Update list item
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn
#################################

$body = "Default metdata set - OK"
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
