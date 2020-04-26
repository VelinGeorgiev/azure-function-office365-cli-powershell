using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$o365 = "$(Get-Location)\HttpTrigger1\node_modules\.bin\o365.cmd"

$out1 = & $o365 login --authType identity --debug #-u a04566df-9a65-4e90-ae3d-574572a16423 

Write-Host $out1

$out2 = & $o365 spo web get -u https://veling.sharepoint.com

Write-Host $out2

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $out2
})
