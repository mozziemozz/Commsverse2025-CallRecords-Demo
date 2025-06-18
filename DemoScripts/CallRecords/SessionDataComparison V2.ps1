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

$v1CallRecord = Get-Content -Path .\.local\SampleData\CallRecords\8ae62ca5-c0bf-438d-9303-2a18b1ddc43a_call_v1_2025-05-18T18-42-28Z.json | ConvertFrom-Json -Depth 99

$v1CallRecord | fl id, version, startDateTime, endDateTime, lastModifiedDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
@{Name="ParticipantsPhone";Expression={$_.participants.phone.id}}, `
@{Name="ParticipantsUsers";Expression={$_.participants.user.displayName -join "; "}}

Write-Host "Call Record v1 was last modified on $($v1CallRecord.lastModifiedDateTime)" -ForegroundColor Cyan
Write-Host "Call Record v1 was published after $(($v1CallRecord.lastModifiedDateTime - $v1CallRecord.endDateTime).ToString('hh\:mm\:ss')) (last modified date time - call end date time)" -ForegroundColor Cyan

Read-Host "Press any key to show sessions of V1..."

# Sessions data v1
$v1 = Get-Content -Path .\.local\SampleData\CallRecordSessions\8ae62ca5-c0bf-438d-9303-2a18b1ddc43a_sessions_v1_2025-05-18T18-42-28Z.json | ConvertFrom-Json -Depth 99

# Read-Host "Press any key to show user agent roles in V1..."

$v1.value | Sort-Object -Property startDateTime | ft startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
@{Name="callerSidePhone";Expression={$_.caller.identity.phone.id}}, `
@{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}, `
@{Name="CallerSideUserAgentHeader";Expression={$_.caller.useragent.headerValue.Substring(0,12)}}, `
@{Name="CallerSideUserAgentRole";Expression={$_.caller.useragent.role}}, `
@{Name="CalleeSideUsers";Expression={$_.callee.identity.user.displayName}}, `
@{Name="CalleeSideUserAgentRole";Expression={$_.callee.useragent.role}}

Write-Host "Call Record V1 contains $($v1.value.Count) sessions." -ForegroundColor Cyan

$agentSessionV1 = $v1.value | Where-Object { $_.caller.identity.user.displayName -eq "Evelyn Carter" }

$agentSessionV1Duration = ($agentSessionV1.endDateTime - $agentSessionV1.startDateTime).ToString('hh\:mm\:ss')

Write-Host "Longest Agent Session Duration in V1 for user '$($agentSessionV1.caller.identity.user.displayName)': $agentSessionV1Duration" -ForegroundColor Cyan

Read-Host "Press any key to show V2..."

$v2CallRecord = Get-Content -Path .\.local\SampleData\CallRecords\8ae62ca5-c0bf-438d-9303-2a18b1ddc43a_call_v2_2025-05-18T19-47-03Z.json | ConvertFrom-Json -Depth 99

$v2CallRecord | fl id, version, startDateTime, endDateTime, lastModifiedDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
@{Name="ParticipantsPhone";Expression={$_.participants.phone.id}}, `
@{Name="ParticipantsUsers";Expression={$_.participants.user.displayName -join "; "}}

Write-Host "Call Record v2 was last modified on $($v2CallRecord.lastModifiedDateTime)" -ForegroundColor Cyan
Write-Host "Call Record v2 was published after $(($v2CallRecord.lastModifiedDateTime - $v2CallRecord.endDateTime).ToString('hh\:mm\:ss')) (last modified date time (v2) - call end date time)" -ForegroundColor Cyan

Write-Host "Call Record v2 was published $(($v2CallRecord.lastModifiedDateTime - $v1CallRecord.lastModifiedDateTime).ToString('hh\:mm\:ss')) after v1 (last modified date time (v2) - last modified date time (v1). This only works if you keep both versions...)" -ForegroundColor Cyan

Read-Host "Press any key to show sessions of V2..."

# Sessions data v2
$v2 = Get-Content -Path .\.local\SampleData\CallRecordSessions\8ae62ca5-c0bf-438d-9303-2a18b1ddc43a_sessions_v2_2025-05-18T19-47-03Z.json | ConvertFrom-Json -Depth 99

# Read-Host "Press any key to show user agent roles of V2..."

$v2.value | Sort-Object -Property startDateTime | ft startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
@{Name="callerSidePhone";Expression={$_.caller.identity.phone.id}}, `
@{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}, `
@{Name="CallerSideUserAgentHeader";Expression={$_.caller.useragent.headerValue.Substring(0,12)}}, `
@{Name="CallerSideUserAgentRole";Expression={$_.caller.useragent.role}}, `
@{Name="CalleeSideUsers";Expression={$_.callee.identity.user.displayName}}, `
@{Name="CalleeSideUserAgentRole";Expression={$_.callee.useragent.role}}

Write-Host "Call Record V2 contains $($v2.value.Count) sessions. Additional sessions in V2: $($v2.value.Count - $v1.value.Count)" -ForegroundColor Cyan

# Read-Host "Press any key to sort by duration..."

# $v2.value | % {$_ | Add-Member -MemberType NoteProperty -Name duration -Value ($_.endDateTime - $_.startDateTime).TotalSeconds }

# $sessionDurations = $v2.value | Sort-Object -Property duration -Descending

# $sessionDurations | ft startDateTime, endDateTime, duration, `
# @{Name="callerSidePhone";Expression={$_.caller.identity.phone.id}}, `
# @{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}, `
# @{Name="CallerSideUserAgentRole";Expression={$_.caller.useragent.role}}, `
# @{Name="CallerSideUserId";Expression={$_.caller.identity.user.id}}, `
# @{Name="CalleeSideUsers";Expression={$_.callee.identity.user.displayName}}, `
# @{Name="CalleeSideUserAgentRole";Expression={$_.callee.useragent.role}}

# Find the session answered by the agent
$agentSessionV2 = $v2.value | Where-Object { $_.caller.identity.user.displayName -eq "Evelyn Carter" -and ($_.endDateTime - $_.startDateTime).TotalSeconds -gt 0 }

$firstSession = $v2.value | Sort-Object -Property startDateTime | Select-Object -First 1

$agentSessionV2Duration = ($agentSessionV2.endDateTime - $agentSessionV2.startDateTime).ToString('hh\:mm\:ss')

$totalWaitTime = ($agentSessionV2.startDateTime - $firstSession.startDateTime).ToString('hh\:mm\:ss')

$totalDurationCallRecord = ($v2CallRecord.endDateTime - $v2CallRecord.startDateTime).ToString('hh\:mm\:ss')

Write-Host "Longest Agent Session Duration in V1 for user '$($agentSessionV2.caller.identity.user.displayName)': $agentSessionV2Duration" -ForegroundColor Cyan
Write-Host "Total Caller Wait/Queue Time: $totalWaitTime" -ForegroundColor Cyan
Write-Host "Total Call Duration: $totalDurationCallRecord" -ForegroundColor Cyan

$callerSideParticipants = $v2.value.caller.identity.user

Read-Host "Press any key to filter for voice app sessions (exclude virtual conferencing assistant and user id from answered session)..."

Write-Host '$voiceAppSessions = $callerSideParticipants | Where-Object { $_.displayName -eq $null -and $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($agentSessionV2.caller.identity.user.id)) }' -ForegroundColor Yellow

$voiceAppSessions = $callerSideParticipants | Where-Object { $_.displayName -eq $null -and $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($agentSessionV2.caller.identity.user.id)) }

Write-Host "Display Names of voice apps are not available in the call record, so it will only show the IDs of the voice apps' resource accounts." -ForegroundColor Yellow

$voiceAppSessions | ft id, displayName

Read-Host "Press any key to fetch the user display names from graph..."

foreach ($voiceAppSession in $voiceAppSessions.id) {

    if ($voiceAppSession -eq "483c7f8d-446c-4e32-b9ec-3129ada9c044") {

        # Only for demo purposes
        Write-Host "This is the top level auto attendant. The session started at the same time as the entire call." -ForegroundColor Cyan

    }

    else {

        Write-Host "This is the 'final' voice app in which the call was answered. The session started last of all sessions that contain voice apps." -ForegroundColor Cyan

    }

    $v2.value | Where-Object { $_.caller.identity.user.id -eq $($voiceAppSession) } | ft startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
        @{Name="callerSidePhone";Expression={$_.caller.identity.phone.id}}, `
        @{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}, `
        @{Name="CallerSideUserAgentRole";Expression={$_.caller.useragent.role}}, `
        @{Name="CallerSideUserId";Expression={$_.caller.identity.user.id}}, `
        @{Name="CalleeSideUserDisplayName";Expression={Get-MgUser -UserId $_.caller.identity.user.id | Select-Object -ExpandProperty DisplayName}}

    if ($voiceAppSession -eq "db5b2edb-29f9-444c-9e5c-321acfb3d039") {

        Write-Host "This is the session that was answered by the agent. The session doesn't actually include the id of the call queue and started after the call queue session, because the session above just represents the wait in that queue." -ForegroundColor Cyan
        Write-Host "However, by looking at the start times of the sessions that did involve voice apps, we can determine, that the session above was the last call queue session before the agent answered the call." -ForegroundColor Cyan

        $agentSessionV2 | ft startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
        @{Name="callerSidePhone";Expression={$_.caller.identity.phone.id}}, `
        @{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}, `
        @{Name="CallerSideUserAgentRole";Expression={$_.caller.useragent.role}}, `
        @{Name="CallerSideUserId";Expression={$_.caller.identity.user.id}}

    }

}

Write-Host "Press any key to return to the demo selector..." -ForegroundColor Blue
Read-Host