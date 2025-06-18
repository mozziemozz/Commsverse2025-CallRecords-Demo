# Input bindings are passed in via param block.
param($QueueItem, $TriggerMetadata)

$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"

$functionInvocationTime = (Get-Date).ToUniversalTime()

$invocationId = $TriggerMetadata.InvocationId.Substring(0, 8)

# $queueItemPayload = $QueueItem | ConvertFrom-Json

[string]$callRecordId = $QueueItem.callRecordId.Trim()

[int]$currentAttempt = $QueueItem.attempt

[int]$retryAttempt = $currentAttempt

[int]$deliveryCount = $TriggerMetadata.DeliveryCount

Connect-MgGraph -Identity -NoWelcome

$callRecordFetchSuccess = $false

if ($deliveryCount -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Retry attempt because of Azure Storage Queue issues. (Built-in Retry Logic)"

}

else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] First attempt to check call record id. (Built-in Retry Logic)"

}

if ($currentAttempt -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Retry attempt because of delayed data availability in Graph API. (Custom Retry Logic)"

    if ($currentAttempt -gt 3) {

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Retry attempt limit reached. No action."

        return

    }

}
else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] First attempt to check call record id. (Custom Retry Logic)"

}

try {

    $checkCallRecord = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/communications/callRecords/$callRecordId" -ErrorAction Stop

    $callRecordFetchSuccess = $true

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record found in v1.0 API."

}
catch {

    try {
        
        $checkCallRecord = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/communications/callRecords/$callRecordId" -ErrorAction Stop

        $callRecordFetchSuccess = $true

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record found in beta API."

    }
    catch {
        
        $callRecordFetchSuccess = $false

        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record not found in v1.0 nor beta API. Error: $($_.Exception.Message)"

    }
    
}

# Push Ids to queues
if ($callRecordFetchSuccess -eq $true) {

    try {
        $payload = @{
            callRecordId = $callRecordId
            attempt = "1"
        } | ConvertTo-Json

        Push-OutputBinding -Name outputQueueItemProcess -Value $payload

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Added call record id to 'call-record-ids-process' queue."

        $Error.Clear()

    }
    catch {

        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Failed to push message to 'call-record-ids-process' queue: $($_.Exception.Message)"
    
    }

}

else {

    try {
        
        Import-Module Az.Storage

        $connectionString = $env:AzureWebJobsStorage
        $queueName = "call-record-ids-precheck"

        $queueClient = [Azure.Storage.Queues.QueueClient]::new($connectionString, $queueName)
        $queueClient.CreateIfNotExists()

        # Build message with attempt count
        
        $retryAttempt ++
        
        [string]$retryAttemptString = $retryAttempt

        $payload = @{
            attempt = $retryAttemptString
            callRecordId = $callRecordId
        } | ConvertTo-Json

        # Encode message to Base64
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        $messageBase64 = [Convert]::ToBase64String($bytes)

        # Define visibility timeout
        $visibilityTimeout = [System.TimeSpan]::FromSeconds($currentAttempt * 60) # 60 seconds

        # Send to queue
        $null = $queueClient.SendMessageAsync($messageBase64, $visibilityTimeout)

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Added call record id to 'call-record-ids-precheck' queue with a delay of $($currentAttempt * 60)s."

        $Error.Clear()

    }
    catch {
        
        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Failed to push message to 'call-record-ids-precheck' queue: $($_.Exception.Message)"
    
    }


}

$functionEndTime = (Get-Date).ToUniversalTime()

$functionExecutionTime = $functionEndTime - $functionInvocationTime

Write-Information "Function execution time: $($functionExecutionTime.TotalSeconds) seconds."
