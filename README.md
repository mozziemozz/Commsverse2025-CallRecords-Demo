# GraphSubscriptionCallRecordsDemoCommsverse2025

## Azure Functions

1. Create a Linux Azure Function App with Consumption or Flex Consumption plan (including Application Insights) and choose PowerShell 7.4 as runtime.
1. Create a managed identity for the function app
1. Install [Azure Function Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local?tabs=windows%2Cisolated-process%2Cnode-v4%2Cpython-v2%2Chttp-trigger%2Ccontainer-apps&pivots=programming-language-powershell#install-the-azure-functions-core-tools)
1. Install [Az.Accounts](https://www.powershellgallery.com/packages/Az.Accounts/5.1.0) PowerShell module (`Install-PSResource -Name "Az.Accounts"`)
1. Install [Node.JS](https://nodejs.org/en/download/package-manager/all#windows-1) (`winget install OpenJS.NodeJS`)
1. Install [M365CLI](https://pnp.github.io/cli-microsoft365/user-guide/installing-cli) (`npm install -g @pnp/cli-microsoft365`)
1. Run [`m365 setup`](https://pnp.github.io/cli-microsoft365/cmd/setup)
1. Run `.\HelperScripts\Add-EnvironmentVariables.ps1` and fill in all your information/ids.
1. Run `.\HelperScripts\AddGraphPermissionsToMSI.ps1` to add the required permissions to the MSI.
1. Run `.\HelperScripts\AddMSIToSitesSelectedOfSPOSite.ps1` to add the MSI to the allowed sites in Sites.Selected of the SharePoint site.
1. Run `.\HelperScripts\InstallModules.ps1` to download and save the PowerShell modules required by the functions to `.\AzureFunctionsV2\PowerShell\Modules`.
1. Sign into Azure PowerShell using `Connect-AzAccount`.
1. Run `.\HelperScripts\DeployFunctions.ps1` to deploy the functions from `.\AzureFunctionsV2\PowerShell` to the function app.
1. Copy the function url of **Receive-GraphNotifications** and create a new environment variable for the function app called `ReceiveGraphNotificationsFunctionUrl` with the url as value.
1. Create a new function app environment variable called `GraphSubscriptionClientState` and enter any value you like. (I.e. random 8 letters and numbers)
1. Run `.\HelperScripts\Find-DriveId.ps1` to find the drive id of the Team's SharePoint site's document library
1. Create a new function app environment variable called `SharePointDriveId` and add set the drive id as the value.
1. Open the **Renew-GraphSubscription** function and click **Test/Run**. This will create a a new graph subscription. Note: you will only need to run this manually once. It runs daily at 12:15 PM UTC every day.
1. That's it, now any new call record version will be saved to your SharePoint site.

## Call Queue Missed/Answered Script

1. Create a new app registration in Entra ID.
1. Register it as Single Tenant and `http://localhost` as redirect uri for platform **Mobile and desktop applications**.
1. Configure the following application permissions for Microsoft Graph:
    - `CallRecord-PstnCalls.Read.All`
    - `CallRecords.Read.All`
    - `User.Read.All`
1. Grant admin consent to the permissions from the app registration's permissions.
1. Create a new client secret and copy the value
1. Open the project at root folder level in VS Code or any other IDE that supports PowerShell. Your working directory should be the same as where `.\Get-CallRecordByUserId.ps1` is located.
1. Open `.\.local\environmentVariables.json` and paste your the **Application Id** of the newly created app registration as the value for **CallRecordsReadAllAppId**.
1. Replace the `$userId` with the user id of one of your resource accounts in `.\Get-CallRecordByUserId.ps1` on line **1**. This is the user id of the resource account that is assigned to the voice app (auto attendant/call queue) for which you want to check answered/missed calls.
1. Run `.\Get-CallRecordByUserId.ps1`. On first run, it will prompt you for your app registration's client secret. Paste the secret. The secret will be encrypted and stored in `.\.local\SecureCreds\CommsverseDemo2025.txt`. Note: Only the SID that encrypted it can decrypt it again.
1. Uncomment line **39** if you want to export the call history to a csv file. By default it will only output to grid view.