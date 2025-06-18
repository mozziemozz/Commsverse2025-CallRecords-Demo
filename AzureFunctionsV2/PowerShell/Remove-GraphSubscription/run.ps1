using namespace System.Net

param($Request, $TriggerMetadata)

$InformationPreference = "Continue"

Connect-MgGraph -Identity

$clientState = $env:GraphSubscriptionClientState
$functionUrl = $env:ReceiveGraphNotificationsFunctionUrl

$existingSubscriptions = Get-MgSubscription

Write-Information "Existing subscriptions:"
$existingSubscriptions | ConvertTo-Json -Depth 10

if ($existingSubscriptions) {

    foreach ($subscription in $existingSubscriptions) {

        Remove-MgSubscription -SubscriptionId $subscription.Id -Confirm:$false

        if ($?) {
            Write-Information "Subscription with ID $($subscription.Id) removed successfully."
        }
        else {
            Write-Information "Failed to remove subscription with ID $($subscription.Id)."
        }

    }

}
else {

    Write-Information "No existing graph subscription found."

}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $body
})
