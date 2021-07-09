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

# try {
$clientId  = "ab4a01e3-1900-4063-acb0-341f0127bddb" #Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminId' -AsPlainText
$clientSecret = "pL6.TUm4374cda2sU-B71O3_xC-m4ucq_x" #Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminSecret' -AsPlainText

$sourceConn = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection -WarningAction Ignore
$newLibName = "Anbud"
   
$listItem = Get-PnPListItem -List $listId  -Id $listItemId -Connection $sourceConn
$user = $listItem["avPersonEmail"]

$company = $listItem["avSelskap"]
$projectSiteUrl = $listItem["avAnbud_x003A_SharePointUrl"].LookupValue
$destConn = Connect-PnPOnline -Url $projectSiteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection -WarningAction Ignore
$visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -Connection $destConn
Add-PnPGroupMember -LoginName $user -Group $visitorGroup -Connection $destConn
$list = Get-PnPList -Identity $newLibName
if($null -eq $list){
    New-PnPList -Title $newLibName -Template DocumentLibrary -Url $newLibName -EnableVersioning -OnQuickLaunch -Connection $destConn
}
$folderIdentity = $newLibName +"/"+ $company
$roleDefs = Get-PnPRoleDefinition -Connection $destConn
if($null -ne $roleDefs){
    foreach($r in $roleDefs)
    {
        Write-Host $r.RoleTypeKind
        if($r.RoleTypeKind -eq "Editor")
        {
            $roleDef = $r
            Write-Host $roleDef
            $roleDefName = $roleDef.Name
        }
    }
}
if($null -eq $roleDef)
{
    $roleDefName = "Contribute"
}
$folder = Get-PnPFolder -Url $folderIdentity -Connection $destConn -ErrorAction SilentlyContinue
if($null -eq $folder){
    $folder = Resolve-PnPFolder -SiteRelativePath $folderIdentity -Connection $destConn
    $g = Get-PnPGroup -AssociatedMemberGroup
    Set-PnPFolderPermission -List $newLibName -Identity $folderIdentity -Group $g -AddRole $roleDef.Name -ClearExisting -Connection $destConn 
}
Write-Host $roleDefName
Write-Host "Set-PnPFolderPermission"
Set-PnPFolderPermission -List $newLibName -Identity $folderIdentity -User $user -AddRole  $roleDefName -Connection $destConn 
Write-Host "Set-PnPListItem"
Write-Host $listId
Write-Host $listItemId

#$sourceConn = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -ClientSecret $clientSecret -ReturnConnection -WarningAction Ignore
$updatedListItem = Set-PnPListItem -List $listId -Identity $listItemId -Values @{"avAnbudsForsporselUtfort"=$true;"avAnbudsForesporsel"="Godkjent"} -Connection $sourceConn

Write-Host "Disconnect"
Disconnect-PnPOnline -Connection $destConn
Disconnect-PnPOnline -Connection $sourceConn
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body       = $body
  })
