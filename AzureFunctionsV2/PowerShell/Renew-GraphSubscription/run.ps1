# Input bindings are passed in via param block.
param($Timer)

$InformationPreference = "Continue"

# Cron expression: 0 15 12 * * *
# Runs at 12:15 PM UTC every day.

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Information "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Information "PowerShell timer trigger function ran! TIME: $currentUTCtime"

# Disable-AzContextAutosave -Scope Process | Out-Null
# Connect-AzAccount -Identity | Out-Null

Connect-MgGraph -Identity

$clientState = ($env:GraphSubscriptionClientState)
$functionUrl = ($env:ReceiveGraphNotificationsFunctionUrl)

# Write-Information "Client state: $clientState"

$existingSubscriptions = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/subscriptions" -ContentType "application/json"

if ($existingSubscriptions.value.resource -eq "/communications/callRecords") {

    $existingSubscription = $existingSubscriptions.value | Where-Object { $_.resource -eq "/communications/callRecords" }

    Write-Information "Groups subscription already exists. Renewing subscription."

    $jsonBodyRenew = @{
        "expirationDateTime" = (Get-Date).AddDays(2).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        "clientState"        = "$clientState"
    }

    $jsonBodyRenew = $jsonBodyRenew | ConvertTo-Json

    $updateSubscription = Invoke-MgGraphRequest -Method Patch -Uri "https://graph.microsoft.com/v1.0/subscriptions/$($existingSubscription.id)" -Body $jsonBodyRenew -ContentType "application/json"

    $updateSubscription

}

else {

    Write-Information "Creating Call Records subscription."

    $body = [ordered]@{
        "changeType"         = "created,updated"
        "notificationUrl"    = "$functionUrl"
        "resource"           = "/communications/callRecords"
        "expirationDateTime" = (Get-Date).AddDays(2).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        "clientState"        = "$clientState"
    }

    $jsonBody = $body | ConvertTo-Json

    $jsonBody

    $createSubscription = Invoke-MgGraphRequest -Method Post -Uri "https://graph.microsoft.com/v1.0/subscriptions" -Body $jsonBody -ContentType "application/json"

    $createSubscription

}