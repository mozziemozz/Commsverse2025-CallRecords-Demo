using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"

$invocationId = $TriggerMetadata.InvocationId.Substring(0, 8)

$clientState = ($env:GraphSubscriptionClientState)

if ($TriggerMetadata.validationToken) {

    Write-Information "[FunctionId: $invocationId][CallRecordId NULL] Function was invoked by Graph to update subscription. No new call records to process."

}

# Only send call record id to storage queue if client state matches
elseif ($Request.body.value.clientState -eq $clientState -and $Request.body.value.changeType -in @("created", "updated")) {

    [string]$callRecordId = $($Request.body.value).resource.Split("/")[-1].Trim()

    Write-Information "[FunctionId: $invocationId][CallRecordId $callRecordId] Client state matches. Change type: '$($Request.body.value.changeType)'."

    $maxRetries = 3
    $delayBetweenRetries = 1000 # milliseconds
    $attemptCounter = 1

    do {
        try {
            # Send call record id to storage queue
            $payload = @{
                callRecordId = $callRecordId
                attempt = "1"
            } | ConvertTo-Json

            Push-OutputBinding -Name outputQueueItem -Value $payload
            $success = $true # Indicate success

            $Error.Clear()
        }
        catch {
            $attemptCounter ++
                
            if ($attemptCounter -eq $maxRetries) {
                $success = $false # Indicate failure
            }
            else {
                Start-Sleep -Milliseconds $delayBetweenRetries # Wait before retrying
            }
        }
    } until ($success -or $attemptCounter -eq $maxRetries)

    if ($success) {
        Write-Information "[FunctionId: $invocationId][CallRecordId $callRecordId] call record id added to 'call-record-ids-precheck' queue."
    }
    else {
        Write-Error "[FunctionId: $invocationId][CallRecordId $callRecordId] Max retries reached. Could not add call record id to 'call-record-ids-precheck' queue."
    }

}

else {

    Write-Information "Client state does not match. Ignoring."

}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        Headers    = @{'Content-Type' = 'text/plain' }
        StatusCode = [HttpStatusCode]::OK
        Body       = $TriggerMetadata.validationToken
    })