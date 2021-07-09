using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TeamsAdminStarter starter...."
function Connect-To-Graph {
    param(
        [Parameter(Mandatory=$true)][string] $clientId,
        [Parameter(Mandatory=$true)][string] $clientSecret,
        [Parameter(Mandatory=$true)][string] $tenantDomainName
        )
        
        $Body = @{    
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        client_Id     = $clientId
        Client_Secret = $clientSecret
        } 

        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantDomainName/oauth2/v2.0/token" -Method POST -Body $Body

        return $ConnectGraph.access_token
        
}

function Get-Graph-Header {
    param([Parameter(Mandatory=$true)][string] $token)
        
       return @{
          "Authorization" = "Bearer $token"
          "Content-type" = "application/json"
        }

        
}
function Get-Team-Body {
  
        
       return   '{
                          "memberSettings": {
                            "allowCreateUpdateChannels": true,
                            "allowDeleteChannels": true,
                            "allowAddRemoveApps": true,
                            "allowCreateUpdateRemoveTabs": true,
                            "allowCreateUpdateRemoveConnectors": true    
                          },
                          "guestSettings": {
                            "allowCreateUpdateChannels": true,
                            "allowDeleteChannels": true 
                          },
                          "messagingSettings": {
                            "allowUserEditMessages": true,
                            "allowUserDeleteMessages": true,
                            "allowOwnerDeleteMessages": true,
                            "allowTeamMentions": true,
                            "allowChannelMentions": true    
                          },
                          "funSettings": {
                            "allowGiphy": true,
                            "giphyContentRating": "strict",
                            "allowStickersAndMemes": true,
                            "allowCustomMemes": true
                          }
                        }'

        
}
function Get-Group-Body {
     param(
        [Parameter(Mandatory=$true)][string] $Title,
        [Parameter(Mandatory=$true)][string] $mailNickName,
        [Parameter(Mandatory=$true)][string] $Description,
        [Parameter(Mandatory=$true)][object] $owners,
        [object] $members
        )

        $ownersString = ""
        if($owners.Length -gt 0){
            foreach($owner in $owners)
            {
                 $ownersString+= '"https://graph.microsoft.com/v1.0/users/' +$owner.Email+'",'
            }
             $ownersString=$ownersString.Substring(0,$ownersString.Length-1)
        }

        $membersString = ""
        if($members.Length -gt 0){
            foreach($member in $members)
            {
                 $membersString+= '"https://graph.microsoft.com/v1.0/users/' +$member.Email+'",'
            }
             $membersString=$membersString.Substring(0,$membersString.Length-1)
        }


       if($membersString){
        return '{
                "displayName":  "' + $Title +'",
                "mailNickname":  "' + $mailNickName +'",
                "description":  "' + $Description +'",    
                "owners@odata.bind":  [' + $ownersString + '],
                "members@odata.bind":  [' + $membersString + '],
                "groupTypes":  [ "Unified" ],
                "mailEnabled":  "true",
                "securityEnabled":  "false",
                "visibility": "Private"
            }'
       }else{
        return '{
                "displayName":  "' + $Title +'",
                "mailNickname":  "' + $mailNickName +'",
                "description":  "' + $Description +'",    
                "owners@odata.bind":  [' + $ownersString + '],
                "groupTypes":  [ "Unified" ],
                "mailEnabled":  "true",
                "securityEnabled":  "false",
                "visibility": "Private"
            }'
       
       }
}
$listItemId = $Request.Query.listItemId

if (-not $listItemId) {
    $body = "listItemId is required"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::NotAcceptable
        Body = $body
    })
    return
}

try{
    $clientId  = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminId' -AsPlainText
    $clientSecret = Get-AzKeyVaultSecret -VaultName 'DimbaTeamsAdminKV' -Name 'TeamsAdminSecret' -AsPlainText

    #$clientId = "3e39183d-c9a2-4176-98a5-45a6f9298dcd"
    #$clientSecret = "B-1Q5E~.176t.R6M1KFlNuLOKjBTXrYB_z"
    $tenantDomainName = "dimbaas.onmicrosoft.com"
    $adminSiteUrl = "https://dimbaas.sharepoint.com/sites/aviadoas"
    $listName = "Anbud"

    #################################
    #Connect to SPO
    $adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
    #################################

    #################################
    #Get item from list
    $listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
    Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Under opprettelse";} -Connection $adminConn

    $SitePrefix = $listItem["avSitePrefix"]
    $Title = $listItem["Title"]
    $Description = $listItem["avAnbudBeskrivelse"]
    $owners = $listItem["avEiere"]
    $members = $listItem["avMedlemmer"]
    $mailNickName = $SitePrefix+"-"+$listItemId #$Title.Replace(" ","")
    $templateSiteUrl = $listItem["avTemplateSiteUrl"]
    #################################

    #################################
    #Connect to graph
    $token = Connect-To-Graph -clientId $clientId -clientSecret $clientSecret -tenantDomainName $tenantDomainName
    $headers = Get-Graph-Header -token $token
    #################################

    #################################
    #Get group body JSON
    $groupBody= Get-Group-Body -Title $Title -mailNickName $mailNickName -Description $Description -owners $owners -members $members
    #################################
    #################################
    #Create unified group
    Write-Host "Oppretter" $Title -ForegroundColor Green
    Write-Host "mailNickName" $mailNickName -ForegroundColor Green

    $GetGroupsUrl = "https://graph.microsoft.com/v1.0/groups/"
    $groupsResponse=Invoke-RestMethod -Uri $GetGroupsUrl  -Method Get -Headers $headers 
    $groups = $groupsResponse.value

    $group = $groups|Where-Object{$_.mailNickName -match $mailNickName}
    if(-not $group.Id)
    {
        $CreateGroupUrl = "https://graph.microsoft.com/v1.0/groups/"
        $group=Invoke-RestMethod -Uri $CreateGroupUrl  -Method POST -Headers $headers -Body $groupBody 
    }
    Write-Host "Group: $group" 
    Start-Sleep -Seconds 10
    #################################

    #################################
    #Get team body JSON
    $teamBody = Get-Team-Body
    #################################

    $groupId = $group.id


    #################################
    #Check if team exists
    $team=$null
    try{
        $team = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams/$($groupId)" -Method Get -Headers $headers -ErrorAction SilentlyContinue
    }
    catch{
        Write-Host "Team failed to create" 
        Write-Host $_
    }
    #################################

    #################################
    #Create Team
    if(-not $team)
    {
        $team = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($groupId)/team" -Method PUT -Headers $headers -Body $teamBody
    }
    Start-Sleep -Seconds 10
    $teamsUrl = $team.webUrl
    Write-Host "teamsUrl: $teamsUrl" 
    #################################

    #################################
    #Get SPO Url
    $webUrlResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($groupId)/sites/root/webUrl" -Method Get -Headers $headers 
    $spoUrl = $webUrlResponse.value
    Write-Host "spoUrl: $spoUrl" 
    #################################

    #################################
    #Update list item
    Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";"avSharePointUrl"=$spoUrl;"avTeamsUrl"=$teamsUrl;"avGroupId"=$groupId;"avGroupEmail"=$mailNickName;} -Connection $adminConn

   
    #################################
    Write-Host "Legger på mal: $templateSiteUrl" 
    Invoke-RestMethod -Uri "https://dimbateamsadmin.azurewebsites.net/api/TeamsAdminTemplate?code=4vXVQqZbT0YEfawJJQkI8FVcpxW4DHLyCuYaBk3COSIuXb5dFVPigQ==&listItemId=$listItemId" -Method Get

    #################################
    Write-Host "Setter defaultmetadata fra: $templateSiteUrl" 
    Invoke-RestMethod -Uri "https://dimbateamsadmin.azurewebsites.net/api/TeamsAdminDefaultMetadata?code=mwe/aPJ6YMIpAk14taqkB0Otu1IR2eafubdtols2iePFzJmLfXfRcA==&listItemId=$listItemId" -Method Get

    #################################
    Write-Host "Kopierer innhold fra: $templateSiteUrl" 
    Invoke-RestMethod -Uri "https://dimbateamsadmin.azurewebsites.net/api/TeamsAdminContent?code=tDsiqdg26vZOWeM2pwKnazhf9rJOGGhBMrpdCpkE0Is/CTljCSOXrw==&listItemId=$listItemId" -Method Get


    #################################
    #Update list item
    Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus"="Opprettet";} -Connection $adminConn


    $body = $spoUrl

    # #################################
}catch{
    $body = "Something went wrong"
    Write-Host $_
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::NotAcceptable
        Body = $body
    })
    return

}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
