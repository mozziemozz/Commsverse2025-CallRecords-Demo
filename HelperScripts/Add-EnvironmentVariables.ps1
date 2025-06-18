$environmentVariables = @{
    "TenantId" = $tenantId = Read-Host -Prompt "Enter your tenant ID"
    "M365CLIAppId" = $m365CLIAppId = Read-Host -Prompt "Enter your M365 CLI app ID"
    "MSIId" = $msiId = Read-Host -Prompt "Enter your MSI ID (Managed Identity/Service Principal object ID of the function app)"
    "CallRecordsReadAllAppId" = $callRecordsReadAllAppId = Read-Host -Prompt "Enter the app id for the Call Records Read All app (skip by pressing Enter if you don't have this yet)"
    "GroupId" = $groupId = Read-Host -Prompt "Enter the group ID of the Team's SharePoint site where you want to save the call records"
    "SiteName" = $siteName = Read-Host -Prompt "Enter the name of the SharePoint site where you want to save the call records. I.e. 'Commsverse Demo 2025'"
    "FunctionAppName" = $functionAppName = Read-Host -Prompt "Enter the name of the Azure Function App where the PowerShell functions will be deployed"
}

if (!(Test-Path -Path ".\.local")) {

    New-Item -Path ".\.local" -ItemType Directory -Force | Out-Null

}

Set-Content -Path ".\.local\environmentVariables.json" -Value ($environmentVariables | ConvertTo-Json -Depth 10)