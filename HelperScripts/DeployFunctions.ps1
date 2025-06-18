# Add the PowerShell functions to the function app

$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

$functionAppNamePowerShell = $environmentVariables.FunctionAppName

$workingDirectory = Get-Location

$functionLocationPowerShell = ".\AzureFunctionsV2\PowerShell"
Set-Location -Path $functionLocationPowerShell
func azure functionapp publish $functionAppNamePowerShell --powershell

Set-Location -Path $workingDirectory
