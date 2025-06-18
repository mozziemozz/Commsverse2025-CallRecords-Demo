. .\Modules\SecureCredsMgmt.ps1

$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

. Get-MZZSecureCreds -FileName "CommsverseDemo2025"
$appSecret = $passwordEncrypted

# Create new powershell credential object from app id (user name) and app secret (password)
$clientSecretCredential = New-Object System.Management.Automation.PSCredential ($($environmentVariables.CallRecordsReadAllAppId), $appSecret)

# Connect to Graph
Connect-MgGraph -ClientSecretCredential $clientSecretCredential -TenantId $environmentVariables.TenantId -NoWelcome

$mgContext = Get-MgContext | Select-Object AppName, AuthType, TokenCredentialType, Scopes

Write-Host "Graph Context:" -ForegroundColor Green

$mgContext | Format-List

$userId = "483c7f8d-446c-4e32-b9ec-3129ada9c044" # This is the user id of the top level resource account (auto attendant or call queue)

# $userId = "f4dc33df-7bff-42a2-8ea9-e9c5bcfb7866"

# $fromTime = (Get-Date).ToUniversalTime().AddDays(-20).ToString("yyyy-MM-ddTHH:mm:ssZ")
# $toTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# $callRecords = (Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/communications/callRecords?`$filter=participants_v2/any(p:p/id eq '$userId') and startDateTime ge $fromTime and startDateTime lt $toTime" -ContentType "application/json").value | ConvertTo-Json -Depth 99 | ConvertFrom-Json -Depth 99

$callRecords = (Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/beta/communications/callRecords?`$filter=participants_v2/any(p:p/id eq '$($userId)')" -ContentType "application/json").value

# Create proper objects
$callRecords = $callRecords | ConvertTo-Json -Depth 99 | ConvertFrom-Json -Depth 99

$callRecords | Sort-Object startDateTime | Format-Table id, version, type, modalities, startDateTime, endDateTime, @{Name="CallerNumber";Expression={$_.organizer_v2.id.Substring(0,5)}} # Anonymized output for demo

$callSummaries = @()

foreach ($callRecord in $callRecords | Where-Object { $_.startDateTime -gt (Get-Date -Date "20.05.2025").ToUniversalTime() }) {

    . .\DemoScripts\CallRecords\Analyze-CallQueueCallRecord.ps1 -CallId $callRecord.id

}

$callSummaries = $callSummaries | Sort-Object -Property StartDateTime

# $callSummaries | Export-Csv -Path "C:\Temp\CallSummariesAll-$userId.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"

$callSummaries | Out-GridView -Title "Call Summaries"

Write-Host "Press any key to return to the demo selector menu..." -ForegroundColor Blue
Read-Host