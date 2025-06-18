& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -Command {
    $modules = @(
        @{ Name = "Microsoft.Graph.Identity.SignIns"; Version = "2.28.0" },
        @{ Name = "Microsoft.Graph.Authentication"; Version = "2.28.0" },
        @{ Name = "Microsoft.Graph.ChangeNotifications"; Version = "2.28.0" },
        @{ Name = "Microsoft.Graph.CloudCommunications"; Version = "2.28.0" },
        @{ Name = "Az.Accounts"; Version = "5.1.0" },
        @{ Name = "Az.Storage"; Version = "9.0.0" }
    )

    $modulesFolder = ".\AzureFunctionsV2\PowerShell\Modules"

    foreach ($mod in $modules) {
        Save-Module -Name $mod.Name -RequiredVersion $mod.Version -Path $modulesFolder -Force
    }
}