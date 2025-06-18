# Input bindings are passed in via param block.
param($QueueItem, $TriggerMetadata)

$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"

$functionInvocationTime = (Get-Date).ToUniversalTime()

$invocationId = $TriggerMetadata.InvocationId.Substring(0, 8)

[string]$userId = $QueueItem.userId.Trim()

[int]$currentAttempt = $QueueItem.attempt

[int]$retryAttempt = $currentAttempt

[int]$deliveryCount = $TriggerMetadata.DeliveryCount

Connect-MgGraph -Identity -NoWelcome

if ($deliveryCount -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] Retry attempt because of Azure Storage Queue issues. (Built-in Retry Logic)"

}

else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] First attempt to check user presence id. (Built-in Retry Logic)"

}

if ($currentAttempt -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] Retry attempt because of delayed data availability in Graph API. (Custom Retry Logic)"

    if ($currentAttempt -gt 3) {

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] Retry attempt limit reached. No action."

        return

    }

}
else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] First attempt to check user presence id. (Custom Retry Logic)"

}

try {

    Write-Host "Getting presence for user with ID: '$userId' using 'Get-MgUserPresence -UserId `$userId'" -ForegroundColor Cyan

    $userPresence = Get-MgUserPresence -UserId $userId

    $userPresence | Format-Table

}
catch {

    Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][UserId: $userId] user presence not found in v1.0 nor beta API. Error: $($_.Exception.Message)"

}

$functionEndTime = (Get-Date).ToUniversalTime()

$functionExecutionTime = $functionEndTime - $functionInvocationTime

Write-Information "Function execution time: $($functionExecutionTime.TotalSeconds) seconds."
