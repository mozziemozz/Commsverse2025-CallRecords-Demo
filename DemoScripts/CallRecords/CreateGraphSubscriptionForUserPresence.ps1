Disconnect-MgGraph

Connect-MgGraph

$existingSubscriptions = Get-MgSubscription

if ($existingSubscriptions) {

    foreach ($existingSubscription in $existingSubscriptions) {

        Remove-MgSubscription -SubscriptionId $existingSubscription.id

    }

}

$clientState = Get-Content -Path .\.local\Resources\clientState.txt

$functionUrl = Get-Content -Path .\.local\Resources\functionUrlPresence.txt

$userId = Get-Content -Path .\.local\Resources\userId.txt

$body = [ordered]@{
    "changeType"         = "updated"
    "NotificationUrl"    = "$functionUrl"
    "resource"           = "communications/presences/$userId"
    "expirationDateTime" = (Get-Date).AddMinutes(59).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    "clientState"        = "$clientState"
}

$jsonBody = $body | ConvertTo-Json

$newSubscription = Invoke-MgGraphRequest -Method Post -Uri "https://graph.microsoft.com/v1.0/subscriptions" -Body $jsonBody -ContentType "application/json"

Write-Host "Created new subscription for user presence with ID: $($newSubscription.id)" -ForegroundColor Green

# Removed function url key from output for security reasons
Get-MgSubscription -SubscriptionId $newSubscription.id | Format-List ApplicationId, ChangeType, ClientState, CreatorId, EncryptionCertificate, EncryptionCertificateId, ExpirationDateTime, Id, IncludeResourceData, LatestSupportedTlsVersion, LifecycleNotificationUrl, NotificationQueryOptions, @{Name="NotificationUrl";Expression={$_.NotificationUrl.Split("?")[0]}}, NotificationUrlAppId, Resource, AdditionalProperties

Write-Host "Let's change the presence of the user in Teams..." -ForegroundColor Cyan

Write-Host "Press any key to return to the demo selector menu..." -ForegroundColor Blue
Read-Host