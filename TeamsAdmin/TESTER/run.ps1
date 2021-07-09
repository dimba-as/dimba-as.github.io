using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "TESTER STARTED"
$listItemId = $Request.Body.listItemId
$listId = $Request.Body.listId
$siteUrl = $Request.Body.siteUrl

Write-Host $listItemId
Write-Host $listId
Write-Host $siteUrl
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = "OK"
})
