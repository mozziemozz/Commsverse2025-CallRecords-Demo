$sessionDataFiles = Get-ChildItem -Path .\.local\SampleData\CallRecordSessions\UserAgentExamples

foreach ($sessionDataFile in $sessionDataFiles) {

    $sessions = Get-Content -Path $sessionDataFile.FullName | ConvertFrom-Json -Depth 99

    # Write-Host "Processing file: $($sessionDataFile.Name)" -ForegroundColor Cyan

    $currentSession = $sessions.value | Sort-Object -Property startDateTime | Select-Object startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
        @{Name="callerSideParticipantsPhone";Expression={$_.caller.identity.phone.id}}, `
        @{Name="CalleeSideParticipantsUsers";Expression={$_.callee.identity.user.displayName}}, `
        @{Name="UserAgentPlatform";Expression={$_.callee.useragent.platform}}, `
        @{Name="CalleeHostName";Expression={$_.callee.name}}, `
        @{Name="CalleeCPUName";Expression={$_.callee.cpuName}}
        # @{Name="UserAgentHeader";Expression={$_.callee.useragent.headerValue}}

    $userAgentHeader = $sessions.value.callee.userAgent.headerValue
        
    Write-Host $userAgentHeader -ForegroundColor Cyan

    Read-Host "Press any key to reveal user agent platform information..."

    $currentSession | ft

    Read-Host "Press any key to continue to next session..."

}

Write-Host "Press any key to return to the demo selector menu..." -ForegroundColor Blue
Read-Host

# $sessions | ft