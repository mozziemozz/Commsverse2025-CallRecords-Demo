$callRecordId = "d35d9037-8eb6-43e4-a51e-0236387ee97f"

$subscriptionId = "12345678-1234-1234-1234-123456789012"

$clientState = Get-Content -Path .\.local\Resources\clientState.txt

$functionUrl = Get-Content -Path .\.local\Resources\functionUrl.txt

$body = @"
{
    "value": [
    {
        "subscriptionId": "$subscriptionId",
        "clientState": "$clientState",
        "changeType": "updated",
        "resource": "communications/callRecords/$callRecordId",
        "subscriptionExpirationDateTime": "2024-10-18T14:01:52.406+02:00",
        "resourceData": "System.Management.Automation.OrderedHashtable",
        "tenantId": "4bffbf87-53a0-4fce-b58b-4179cb3a3b7d"
    }
    ]
}
"@

Write-Host "Press enter to send a simulated Graph subscription notification to the Azure Function..." -ForegroundColor Cyan

$body

Read-Host

$startDateTime = (Get-Date).ToUniversalTime()

$request = Invoke-WebRequest -Method Post -ContentType "application/json" -Uri $functionUrl -Body $body

$endDateTime = (Get-Date).ToUniversalTime()

Write-Host "Function called successfully in $((New-TimeSpan -Start $startDateTime -End $endDateTime).TotalSeconds) seconds." -ForegroundColor Green

$request | Format-List StatusCode, StatusDescription, RawContent

Write-Host "Press any key to return to the demo selector menu..." -ForegroundColor Blue
Read-Host