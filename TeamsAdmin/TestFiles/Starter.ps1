
# Write to the Azure Functions log stream.
Write-Host "TeamsAdminStarter starter...."
function Connect-To-Graph {
  param(
    [Parameter(Mandatory = $true)][string] $clientId,
    [Parameter(Mandatory = $true)][string] $clientSecret,
    [Parameter(Mandatory = $true)][string] $tenantDomainName
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
  param([Parameter(Mandatory = $true)][string] $token)
        
  return $headers = @{
    "Authorization" = "Bearer $token"
    "Content-type"  = "application/json"
  }

        
}
function Get-Team-Body {
  
        
  return $teamBody =
  '{
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
    [Parameter(Mandatory = $true)][string] $Title,
    [Parameter(Mandatory = $true)][string] $mailNickName,
    [Parameter(Mandatory = $true)][string] $Description,
    [Parameter(Mandatory = $true)][object] $owners,
    [object] $members
  )

  $ownersString = ""
  if ($owners.Length -gt 0) {
    foreach ($owner in $owners) {
      $ownersString += '"https://graph.microsoft.com/v1.0/users/' + $owner.Email + '",'
    }
    $ownersString = $ownersString.Substring(0, $ownersString.Length - 1)
  }

  $membersString = ""
  if ($members.Length -gt 0) {
    foreach ($member in $members) {
      $membersString += '"https://graph.microsoft.com/v1.0/users/' + $member.Email + '",'
    }
    $membersString = $membersString.Substring(0, $membersString.Length - 1)
  }


  if ($membersString) {
    return $groupBody =
    '{
                "displayName":  "' + $Title + '",
                "mailNickname":  "' + $mailNickName + '",
                "description":  "' + $Description + '",    
                "owners@odata.bind":  [' + $ownersString + '],
                "members@odata.bind":  [' + $membersString + '],
                "groupTypes":  [ "Unified" ],
                "mailEnabled":  "true",
                "securityEnabled":  "false",
                "visibility": "Private"
            }'
  }
  else {
    return $groupBody =
    '{
                "displayName":  "' + $Title + '",
                "mailNickname":  "' + $mailNickName + '",
                "description":  "' + $Description + '",    
                "owners@odata.bind":  [' + $ownersString + '],
                "groupTypes":  [ "Unified" ],
                "mailEnabled":  "true",
                "securityEnabled":  "false",
                "visibility": "Private"
            }'
  }
      

        
}

$listItemId = 12

if (-not $listItemId) {
  $body = "listItemId is required"
  Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
      StatusCode = [HttpStatusCode]::NotAcceptable
      Body       = $body
    })
  return
}

# try {
$clientId = "153767e6-5de3-4c6b-89a3-d7299fc3aeb4"
$clientSecret = "8Kv1qIE1vK.j.__WdZSqpuQ~2o.hApSMt0"
$tenantDomainName = "norgesvel.onmicrosoft.com"

$adminSiteUrl = "https://norgesvel.sharepoint.com/sites/Prosjekt"

#################################
#Connect to SPO
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Get item from list
$listName = "ProsjektoversiktProvisjonering"
$listItem = Get-PnPListItem -List $listName -Id $listItemId  -Connection $adminConn
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus" = "Under opprettelse"; } -Connection $adminConn

#$SitePrefix = $listItem["avSitePrefix"]
$Title = $listItem["Title"]
$Prosjektnummer = $listItem["avProsjektnummer"]
$Description = $listItem["avBeskrivelse"]
$owners = $listItem["avProsjekteier"]
$members = $listItem["avProsjektmedlemmer"]
$mailNickName = $Prosjektnummer + "-" + $Title
$createTeam = $listItem["avTeam"]
$createPrivateChannel = $listItem["avPrivateChannel"]
#Get template
$templateList = "Prosjektmaler"
$templateSiteCol = $listItem["avMal"]
$templateListItem = Get-PnPListItem -List $templateList -Id $templateSiteCol.LookupId -Connection $adminConn
$templateSiteUrl = $templateListItem["avUrl"]
#################################

#################################
#Connect to graph
$token = Connect-To-Graph -clientId $clientId -clientSecret $clientSecret -tenantDomainName $tenantDomainName
$headers = Get-Graph-Header -token $token
$graphConn = Connect-PnPOnline -Url $adminSiteUrl -AccessToken $token -ReturnConnection
#################################

#################################
#Get group body JSON
$groupBody = Get-Group-Body -Title $mailNickName -mailNickName $mailNickName -Description $Description -owners $owners -members $members
#################################
#################################
#Create unified group
Write-Host "Oppretter" $Title -ForegroundColor Green
$GetGroupsUrl = 'https://graph.microsoft.com/v1.0/groups?$filter=startsWith(mailNickname,' + "'$mailNickName'" + ')'
$groupsResponse = Invoke-RestMethod -Uri $GetGroupsUrl  -Method Get -Headers $headers 
$groups = $groupsResponse.value

$group = $groups | Where-Object { $_.mailNickName -match $mailNickName }
if ($null -eq $group.Id) {
  $CreateGroupUrl = "https://graph.microsoft.com/v1.0/groups/"
  $group = Invoke-RestMethod -Uri $CreateGroupUrl  -Method POST -Headers $headers -Body $groupBody 
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
$team = $null
try {
  $team = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/teams/$($groupId)" -Method Get -Headers $headers -ErrorAction SilentlyContinue
}
catch {

}
#################################

#################################
#Create Team
  #Create Team
    if ($null -eq $team -and $createTeam -eq $true) {
       $team = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($groupId)/team" -Method PUT -Headers $headers -Body $teamBody
        Start-Sleep -Seconds 10
    }
    $teamsUrl = $team.webUrl
    Write-Host "teamsUrl: $teamsUrl" 

    #Create Private Channel
    $channel = Get-PnPTeamsChannel -Team $groupId -Identity "Ekstern deling" -Connection $graphConn
    if ($createPrivateChannel -eq $true -and $null -eq $channel) {
         $privateChannel = Add-PnPTeamsChannel -Team $groupId -DisplayName "Ekstern deling" -Description "Denne kanalen skal benyttes til ekstern samhandling" -Private -OwnerUPN $owners.Email -Connection $graphConn -ErrorAction SilentlyContinue 
    }
#################################
#Get SPO Url
$webUrlResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($groupId)/sites/root/webUrl" -Method Get -Headers $headers 
$spoUrl = $webUrlResponse.value
Write-Host "spoUrl: $spoUrl" 
#################################

#################################
#Connect to SPO again
$adminConn = Connect-PnPOnline -url $adminSiteUrl -ClientSecret $clientSecret -ClientId $clientId -WarningAction Ignore -ReturnConnection
#################################

#################################
#Update list item
Set-PnPListItem -List $listName -Identity $listItemId  -Values @{"avStatus" = "Opprettet"; "avSharePointUrl" = $spoUrl; "avTeamsUrl" = $teamsUrl; "avGroupId" = $groupId; "avGroupEmail" = $mailNickName; } -Connection $adminConn

$body = $spoUrl

# #################################
# }
# catch {
#     $body = "Something went wrong"
#     Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
#             StatusCode = [HttpStatusCode]::NotAcceptable
#             Body       = $body
#         })
#     return

# }
Disconnect-PnPOnline -Connection $adminConn
Disconnect-PnPOnline -Connection $graphConn
