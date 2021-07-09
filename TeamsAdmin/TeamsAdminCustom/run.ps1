using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TeamsAdminCustom starter...."

$listItemId = $Request.Query.listItemId

if (-not $listItemId) {
  $body = "listItemId is required"
  Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
      StatusCode = [HttpStatusCode]::NotAcceptable
      Body       = $body
    })
  return
}

# try {
  $clientId  = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminId' -AsPlainText
  $clientSecret = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminSecret' -AsPlainText
$adminSiteUrl = Get-ChildItem env:AdminSiteUrl
$adminSiteUrl = $adminSiteUrl.value #"https://norgesvel.sharepoint.com/sites/Prosjekt"

#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listName = Get-ChildItem env:ProjectList
$listName = $listName.value #"ProsjektoversiktProvisjonering"
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn


#################################
$spoUrl = $listItem["avSharePointUrl"]
$spoConn = Connect-PnPOnline -url $spoUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
Install-PnPApp -Identity 0b76954b-6ef8-4144-bf3b-7c95adbd10b6 -Connection $spoConn -ErrorAction Ignore

$body = "Custom - OK "


Disconnect-PnPOnline -Connection $adminConn
Disconnect-PnPOnline -Connection $spoConn
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body       = $body
  })
