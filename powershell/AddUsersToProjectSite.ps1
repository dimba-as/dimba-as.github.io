$listItemId = 78
$listId = ""
Connect-PnPOnline -Url https://dimbaas.sharepoint.com -Interactive

$searchResult = Submit-PnPSearchQuery -Query "ContentType:Foresporsel"

foreach($row in $searchResult.ResultRows)
{
    $searchListWeb = $row["SPWebUrl"]
    $newLibName = "Anbud"
    Connect-PnPOnline -Url  $searchListWeb -Interactive
    $listId = $row["ListId"]
    $uniqueId = $row["UniqueId"]
    $item = Get-PnPListItem -List $listId -UniqueId $uniqueId -Fields Id
    $listItem = Get-PnPListItem -List $listId -Id $item.Id
    $user = $listItem["avPersonEmail"]
    $company = $listItem["avSelskap"]
    $projectSiteUrl = $listItem["avAnbud_x003A_SharePointUrl"].LookupValue
    $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup
    Connect-PnPOnline -Url $projectSiteUrl -Interactive
    Add-PnPGroupMember -LoginName $user -Group $visitorGroup 
    $list = Get-PnPList -Identity $newLibName
    if($list -eq $null){
       New-PnPList -Title $newLibName -Template DocumentLibrary -Url $newLibName -EnableVersioning 
    }
    $folderIdentity = $newLibName +"/"+ $company
    $folder = Resolve-PnPFolder -SiteRelativePath $folderIdentity
    Set-PnPFolderPermission -List $newLibName -Identity $folderIdentity -User $user -AddRole 'Contribute' -ClearExisting
    
    Connect-PnPOnline -Url  $searchListWeb -Interactive
    Set-PnPListItem -List $listId -Identity $listItem.Id -Values @{"avAnbudsForsporselUtfort"=$true}
}

#Disconnect-PnPOnline