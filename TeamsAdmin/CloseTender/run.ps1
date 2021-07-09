using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "Invite starter...."

$listItemId = $Request.Body.listItemId
$listId = $Request.Body.listId
$siteUrl = $Request.Body.siteUrl

if (-not $listItemId) {
  $body = "listItemId is required"
  Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
      StatusCode = [HttpStatusCode]::NotAcceptable
      Body       = $body
    })
  return
}

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

$sourceConn = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection -WarningAction Ignore
  

$listItem = Get-PnPListItem -List $listId -Id $listItemId -Connection $sourceConn
$projectSiteUrl = $listItem["avSharePointUrl"]
$destConn = Connect-PnPOnline -Url $projectSiteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection -WarningAction Ignore

$listName = "Anbud"
#$listItems = Get-PnPListItem -List $listName -Connection $destConn 
Set-PnPList -Identity $listName -ClearSubscopes -BreakRoleInheritance:$false -Connection $destConn 
<#
Get-PnpUser -Connection $destConn
ForEach($ListItem in $listItems)
{
    #Check if the Item has unique permissions
    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property "HasUniqueRoleAssignments" -Connection $destConn
    If($HasUniquePermissions)
    {       
      Set-PnPListItemPermission -List $ListName -Identity $ListItem.ID -InheritPermissions -Connection $destConn

    }
}
#>
$visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -Connection $destConn 
$members = Get-PnPGroupMember -Group $visitorGroup -Connection $destConn 
foreach($member in $members)
{
    Remove-PnPGroupMember -Group $visitorGroup -LoginName $member.LoginName -Connection $destConn
    $user = Get-PnpUser -Identity $member.LoginName -Connection $destConn 
    Remove-PnPUser -Identity $user -Connection $destConn -Force
    
}
Set-PnPListItem -List $listId -Identity $listItemId -Values @{"avAnbudsavslutningUtfort"=$true;"avAnbudsstatus"="Lukket"} -Connection $sourceConn

Disconnect-PnPOnline -Connection $destConn
Disconnect-PnPOnline -Connection $sourceConn
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body       = $body
  })
