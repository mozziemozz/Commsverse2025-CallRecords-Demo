using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"

$invocationId = $TriggerMetadata.InvocationId.Substring(0, 8)

$clientState = ($env:GraphSubscriptionClientState)

if ($TriggerMetadata.validationToken) {

    Write-Information "[FunctionId: $invocationId][UserId NULL] Function was invoked by Graph to update subscription. No presence changes to process."

}

# Only send call record id to storage queue if client state matches
elseif ($Request.body.value.clientState -eq $clientState -and $Request.body.value.changeType -in @("created", "updated")) {

    [string]$userId = $($Request.body.value).resourceData.id.Trim()

    Write-Information "[FunctionId: $invocationId][UserId $userId] Client state matches. Change type: '$($Request.body.value.changeType)'."

    $Request.body.value | ConvertTo-Json -Depth 99

    $payload = @{
        userId = $userId
        attempt = "1"
    } | ConvertTo-Json

    Push-OutputBinding -Name outputQueueItem -Value $payload

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