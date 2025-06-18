$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

# M365 CLI login
m365 login --appId $environmentVariables.M365CLIAppId --tenant $environmentVariables.TenantId --authType browser

m365 entra approleassignment add --appObjectId $environmentVariables.MSIId --resource "Microsoft Graph" --scopes "CallRecords.Read.All,Sites.Selected"

