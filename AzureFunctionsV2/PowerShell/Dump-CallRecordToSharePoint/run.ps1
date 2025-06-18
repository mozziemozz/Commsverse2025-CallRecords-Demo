# Input bindings are passed in via param block.
param($QueueItem, $TriggerMetadata)

$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"

$functionInvocationTime = (Get-Date).ToUniversalTime()

$invocationId = $TriggerMetadata.InvocationId.Substring(0, 8)

[string]$callRecordId = $QueueItem.callRecordId.Trim()

[int]$currentAttempt = $QueueItem.attempt

[int]$retryAttempt = $currentAttempt

[int]$deliveryCount = $TriggerMetadata.DeliveryCount

# $sharePointDriveId = $env:SharePointDriveId

$sharePointDriveId = "b!cgUfpf-xqUWt9f2we_wQEN08ZX7BDYBCth5m-iHi0DEWRJwb6v3_QIxoxckzE3VF"

Connect-MgGraph -Identity -NoWelcome

if ($deliveryCount -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Retry attempt because of Azure Storage Queue issues. (Built-in Retry Logic)"

}

else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] First attempt to check call record id. (Built-in Retry Logic)"

}

if ($currentAttempt -gt 1) {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Retry attempt because of delayed data availability in Graph API. (Custom Retry Logic)"

    if ($currentAttempt -gt 3) {

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Retry attempt limit reached. No action."

        return

    }

}
else {

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] First attempt to check call record id. (Custom Retry Logic)"

}

# Write out the queue message and insertion time to the information log.
Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Queue item insertion time: $($TriggerMetadata.InsertionTime)"
Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Function invocation time: $($functionInvocationTime)"

$callRecordFetchSuccess = $false
$callRecordSessionsFetchSuccess = $false
$callRecordExportSuccess = $false

$apiVersion = "$apiVersion"

try {

    $checkCallRecord = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/$apiVersion/communications/callRecords/$callRecordId" -ErrorAction Stop

    $callRecordFetchSuccess = $true

    Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record found in v1.0 API."

}
catch {

    $apiVersion = "beta"

    try {
        
        $checkCallRecord = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/$apiVersion/communications/callRecords/$callRecordId" -ErrorAction Stop

        $callRecordFetchSuccess = $true

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record found in beta API."

    }
    catch {
        
        $callRecordFetchSuccess = $false

        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record not found in v1.0 nor beta API. Error: $($_.Exception.Message)"

    }
    
}

if ($callRecordFetchSuccess) {

    try {
        
        $checkCallRecordSessions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/$apiVersion/communications/callRecords/$callRecordId/sessions" -ErrorAction Stop

        if ($checkCallRecordSessions) {

            $callRecordSessionsFetchSuccess = $true

            Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record sessions found in $apiVersion API."

        }

        else {

            $callRecordSessionsFetchSuccess = $false

            Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record sessions not found in $apiVersion API."

        }

    }
    catch {
        
        $callRecordSessionsFetchSuccess = $false

        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record sessions not found in $apiVersion API. Error: $($_.Exception.Message)"

    }

}

else {

    $callRecordExportSuccess = $false

}

if ($callRecordFetchSuccess -and $callRecordSessionsFetchSuccess) {

    $callRecordExport = $checkCallRecord | ConvertTo-Json -Depth 99 | Out-String
    $callRecordSessionsExport = $checkCallRecordSessions | ConvertTo-Json -Depth 99 | Out-String

    $exportDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH-mm-ssZ")

    $destinationNameCallRecord = "$($checkCallRecord.id)_call_v$($checkCallRecord.version)_$($exportDateTime).json"
    $destinationNameCallRecordSessions = "$($checkCallRecord.id)_sessions_v$($checkCallRecord.version)_$($exportDateTime).json"

    $fileContentCallRecord = [System.Text.Encoding]::UTF8.GetBytes($callRecordExport)
    $fileContentCallRecordSessions = [System.Text.Encoding]::UTF8.GetBytes($callRecordSessionsExport)

    try {
        
        # Upload the files to SharePoint
        Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/drives/$sharePointDriveId/root:/CallRecords/$destinationNameCallRecord`:/content" -Body $fileContentCallRecord -ContentType "application/octet-stream"
        Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/drives/$sharePointDriveId/root:/CallRecords/$destinationNameCallRecordSessions`:/content" -Body $fileContentCallRecordSessions -ContentType "application/octet-stream"

        $callRecordExportSuccess = $true

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record and sessions exported to SharePoint."

    }
    catch {
        
        $callRecordExportSuccess = $false

        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId: $callRecordId] Call record and sessions export to SharePoint failed. Error: $($_.Exception.Message)"

    }

}

if ($callRecordFetchSuccess -eq $false -or $callRecordExportSuccess -eq $false) {

    try {
        
        Import-Module Az.Storage

        $connectionString = $env:AzureWebJobsStorage
        $queueName = "call-record-ids-process"

        $queueClient = [Azure.Storage.Queues.QueueClient]::new($connectionString, $queueName)
        $queueClient.CreateIfNotExists()

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

        Write-Information "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Added call record id to 'call-record-ids-process' with a delay of $($currentAttempt * 60)s"

        $Error.Clear()

    }
    catch {
        
        Write-Error "[FunctionId: $invocationId][Count: $deliveryCount][Attempt: $currentAttempt][CallRecordId $callRecordId] Failed to push message to 'call-record-ids-process' queue: $($_.Exception.Message)"
    
    }

}