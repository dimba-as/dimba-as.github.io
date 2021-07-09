
# Write to the Azure Functions log stream.
Write-Host "Invite starter...."

$listItemId = "115"
$listId = "38d17a23-4784-4e7a-9eac-5e1eef0af140"
$siteUrl = "https://dimbaas.sharepoint.com/sites/aviadoas/"

if (-not $listItemId) {
  $body = "listItemId is required"
  Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
      StatusCode = [HttpStatusCode]::NotAcceptable
      Body       = $body
    })
  return
}

# try {
$clientId  = "ab4a01e3-1900-4063-acb0-341f0127bddb" #Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminId' -AsPlainText
$clientSecret = "pL6.TUm4374cda2sU-B71O3_xC-m4ucq_x" #Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminSecret' -AsPlainText

$sourceConn = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection
$newLibName = "Anbud"
   
$listItem = Get-PnPListItem -List $listId  -Id $listItemId -Connection $sourceConn
$user = $listItem["avPersonEmail"]
Write-Host $listItem
$company = $listItem["avSelskap"]
$projectSiteUrl = $listItem["avAnbud_x003A_SharePointUrl"].LookupValue
$destConn = Connect-PnPOnline -Url $projectSiteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection
$visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -Connection $destConn
Add-PnPGroupMember -LoginName $user -Group $visitorGroup -Connection $destConn
$list = Get-PnPList -Identity $newLibName
if($null -eq $list){
    New-PnPList -Title $newLibName -Template DocumentLibrary -Url $newLibName -EnableVersioning -OnQuickLaunch -Connection $destConn
}
$folderIdentity = $newLibName +"/"+ $company
$roleDefs = Get-PnPRoleDefinition -Connection $destConn
    foreach($r in $roleDefs)
    {
        if($r.RoleTypeKind -eq "Editor")
        {
            $roleDef = $r
        }


    }
$folder = Get-PnPFolder -Url $folderIdentity -Connection $destConn -ErrorAction SilentlyContinue
if($null -eq $folder){
    
    $folder = Resolve-PnPFolder -SiteRelativePath $folderIdentity -Connection $destConn
    $g = Get-PnPGroup -AssociatedMemberGroup
    Set-PnPFolderPermission -List $newLibName -Identity $folderIdentity -Group $g -AddRole $roleDef.Name -ClearExisting -Connection $destConn 
}

$roleDefs = Get-PnPRoleDefinition -Connection $destConn
foreach($r in $roleDefs)
{
    if($r.RoleTypeKind -eq "Editor")
    {
        $roleDef = $r
    }


}
Set-PnPFolderPermission -List $newLibName -Identity $folderIdentity -User $user -AddRole $roleDef.Name -Connection $destConn 
Set-PnPListItem -List $listId -Identity $listItem.Id -Values @{"avAnbudsForsporselUtfort"=$true} -Connection $sourceConn

Disconnect-PnPOnline -Connection $destConn
Disconnect-PnPOnline -Connection $sourceConn
# Associate values to output bindings by calling 'Push-OutputBinding'.
