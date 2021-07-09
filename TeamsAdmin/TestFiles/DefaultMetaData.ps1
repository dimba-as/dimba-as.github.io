
# Write to the Azure Functions log stream.
Write-Host "TeamsAdminDefaultMetadata starter.."

$listItemId = 7


if (-not $listItemId) {
    $body = "listItemId is required"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::NotAcceptable
        Body = $body
    })
    return
}
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
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Legger på default metadata";} -Connection $adminConn
$destinationTitle = $listItem["Title"]
$destinationSPUrl = $listItem["avSharePointUrl"]

#Get template
$templateList = "Prosjektmaler"
$templateSiteCol = $listItem["avMal"]
$templateListItem = Get-PnPListItem -List $templateList -Id $templateSiteCol.LookupId -Connection $adminConn
$templateSiteUrl = $templateListItem["avUrl"]
$columns = $templateListItem["avColumns"]

#################################

#################################
#Get metadatafile from template site: client_LocationBasedDefaults.html
$templateConn = Connect-PnPOnline -url $templateSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$templateWeb = Get-PnpWeb -Includes ServerRelativeUrl -Connection $templateConn
$templateWebServerRelativeUrl = $templateWeb.ServerRelativeUrl
$templateLibrary = Get-PnPList -Identity $libraryName -Includes RootFolder.ServerRelativeUrl -Connection $templateConn
if(-not $templateLibrary){
    $templateLibrary = Get-PnPList -Identity "Documents" -Includes RootFolder.ServerRelativeUrl -Connection $templateConn
}

$fileUrl = $templateLibrary.RootFolder.ServerRelativeUrl + "/Forms/client_LocationBasedDefaults.html"

$LBDFileString = Get-PnPFile -Url $fileUrl -AsString -Connection $templateConn 


#################################
#Connect to destination, replace values in file and upload to destination: client_LocationBasedDefaults.html
$destConn = Connect-PnPOnline -url $destinationSPUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$destWeb = Get-PnpWeb -Includes ServerRelativeUrl -Connection $destConn

$destLibrary = Get-PnPList -Identity $libraryName -Includes RootFolder.ServerRelativeUrl -Connection $destConn 
if(-not $destLibrary){
    $destLibrary = Get-PnPList -Identity "Documents" -Includes RootFolder.ServerRelativeUrl -Connection $destConn
}

#To trigger the default metadata event receiver
$field = Add-PnPField -List $destLibrary -DisplayName "DUMMYFIELD" -InternalName "DUMMYFIELD" -Type Text -Connection $destConn
Set-PnPDefaultColumnValues -List $destLibrary -Field "DUMMYFIELD" -Value "DUMMYVALUE" -Connection $destConn
Remove-PnPField -List $destLibrary -Identity $field.Id -Force -Connection $destConn


$destServerRelativeUrl = $destWeb.ServerRelativeUrl
$LBDNewFileString = $LBDFileString.Replace($templateWebServerRelativeUrl,$destServerRelativeUrl)
$LBDNewFileString = $LBDNewFileString.Replace("{SiteTitle}", $destinationTitle)


#<a href="/sites/1006-TEST1006/Delte%20dokumenter"><DefaultValue FieldName="avProsjektnavn">TEST1006</DefaultValue></a>
#Other columns
$columnsArray = $columns.Split(";")
$newColsString = ""
foreach($col in $columnsArray)
{
    $newColsString+= '<a href="'+$destServerRelativeUrl+'/Delte%20dokumenter"><DefaultValue FieldName="'+$col+'">'+$listItem[$col]+'</DefaultValue></a>'

    #Set-PnPDefaultColumnValues -List $libraryName -Field $col -Value  $listItem[$col] -Connection $destConn
}
$newColsString += "</MetadataDefaults>"

$LBDNewFileString = $LBDNewFileString.Replace("</MetadataDefaults>",$newColsString)
Set-Content -Path "D:\home\data\client_LocationBasedDefaults.html" -Value $LBDNewFileString -Encoding utf8 -Force 
Write-Host $LBDNewFileString


$destRootFolderServerRelativeUrl = $destLibrary.RootFolder.ServerRelativeUrl
$destFolderUrl =  $destRootFolderServerRelativeUrl + "/Forms"

$formsFolder= Get-PnpFolder -Url $destFolderUrl -Connection $destConn
$result = Add-PnPFile -Path "D:\home\data\client_LocationBasedDefaults.html" -Folder $formsFolder -Connection $destConn 


#################################
#Update list item
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn
#################################
#Remove-Item -Path "D:\home\data\client_LocationBasedDefaults.html"


################################
#Set metadata to copied items

$destConn = Connect-PnPOnline -url $destinationSPUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
$destLibrary = Get-PnPList -Identity $libraryName -Includes RootFolder.ServerRelativeUrl -Connection $destConn 
if(-not $destLibrary){
    $destLibrary = Get-PnPList -Identity "Documents" -Includes RootFolder.ServerRelativeUrl -Connection $destConn
}

$defVals= Get-PnPDefaultColumnValues -List $destLibrary  -Connection $destConn 
$defVals = $defVals | Sort-Object Path

$fileItems = Get-PnPFolderItem -FolderSiteRelativeUrl $destLibrary.RootFolder.Name -ItemType All -Recursive -Connection $destConn 
$ctx = Get-PnPContext 
foreach($fileItem in $fileItems){
    $fileServerRelUrl = $fileItem.ServerRelativeUrl
    if($fileServerRelUrl -notlike "*/Forms/*")
    {
        $ctx.Load($fileItem.ListItemAllFields)
        $ctx.ExecuteQuery()
        $item = Get-PnPListItem -List $destLibrary -Id $fileItem.ListItemAllFields.Id -Connection $destConn
        foreach($defVal in $defVals){
            $defValMatch = "*"+$defVal.Path +"*"
            if($item.FieldValues.FileDirRef -like $defValMatch)
            {
                 $res= Set-PnPListItem -List $destLibrary -Identity $fileItem.ListItemAllFields.Id -Values @{$defVal.Field=$defVal.Value} -Connection $destConn -UpdateType SystemUpdate
            }
        }
    }

}

$ctx.Dispose()


################################


Disconnect-PnPOnline -Connection $adminConn
Disconnect-PnPOnline -Connection $destConn
Disconnect-PnPOnline -Connection $templateConn
