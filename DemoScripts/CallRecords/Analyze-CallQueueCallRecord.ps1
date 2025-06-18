<#

    .SYNOPSIS

    .DESCRIPTION

    .NOTES

        Version:        1.3.1
        Author:         Martin Heusser
        Link:           https://heusser.pro
        Changes:        2024-05-12: Initial version
                        2024-07-04: Fix unkown calls reprocessing

    .PARAMETER

#>

[CmdletBinding()]
param
(
    [Parameter (Mandatory = $true)]
    [string] $CallId,
    [Parameter (Mandatory = $false)]
    [string] $tenantId,
    [Parameter (Mandatory = $false)]
    [string] $ChangeType

)

if ($host.Name -in @("Visual Studio Code Host", "ConsoleHost")) {

    $localDebugMode = $true

}

else {

    $localDebugMode = $false

}

$callIds = @($callId)

foreach ($callId in $callIds) {

    Write-Output "CallId: $callId"

    $call = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/communications/callRecords/$callId" -ContentType "application/json" -OutputType PSObject

    if (!$?) {

        $call = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/communications/callRecords/$callId" -ContentType "application/json" -OutputType PSObject

        if ($?) {

            $callOrganizer = $call.organizer.phone.id

        }

    }

    else {

        $callOrganizerId = $call.organizer_v2.id

        $callOrganizer = $call.organizer_v2.id

        if ($callOrganizer -notmatch "^\+") {

            $callOrganizer = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($callOrganizer)" -ContentType "application/json").displayName

            $callerIsInternalCaller = $true

        }

        else {

            $callerIsInternalCaller = $false

        }

    }

    if ($call.type -eq "peerToPeer") {

        Write-Output "Call is a peer-to-peer call. Disregarding call record..."
        # exit

    }

    $sessions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/communications/callRecords/$($call.id)/sessions" -OutputType PSObject

    if (!$sessions -or !$call) {

        Write-Output "Error in retrieving call and/or sessions from Graph..."

    }

    else {

    }

    if ($call -and $sessions) {

        Write-Output "Call record and sessions retrieved..."

        # dump call and session data as Json to SharePoint
        Write-Output "Call record:"
        $call

        $callStartDateTime = (Get-Date -Date $call.startDateTime).ToUniversalTime()
        $callEndDateTime = (Get-Date -Date $call.endDateTime).ToUniversalTime()

        $fromTime = $callStartDateTime.AddMinutes(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $toTime = $callEndDateTime.AddMinutes(1).ToString("yyyy-MM-ddTHH:mm:ssZ")

        Write-Output "From time: $fromTime"
        Write-Output "To time: $toTime"

        $pstnCalls = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/beta/communications/callRecords/getPstnCalls(fromDateTime=$($fromTime),toDateTime=$($toTime))" -ContentType "application/json" -OutputType PSObject

        if (!$pstnCalls.value) {

            $pstnCalls = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/beta/communications/callRecords/getDirectRoutingCalls(fromDateTime=$($fromTime),toDateTime=$($toTime))" -ContentType "application/json" -OutputType PSObject

        }

        if (!$pstnCalls.value) {

            Write-Host "No matching PSTN call found. Disregarding call record."

        }

    }

    $pstnCalls = $pstnCalls.value | Where-Object { $_.callType -eq "ucap_in" -or $_.callType -eq "oc_ucap_in" -or $_.callType -eq "ByotInUcap" }

    if ($pstnCalls.Count -gt 1) {

        # Check for matching call id for Direct Routing and Operator Connect calls
        $matchingPstnCall = $pstnCalls | Where-Object { $_.callId -eq $call.id }

        # If Calling Plan call
        if (-not $matchingPstnCall) {

            if ($pstnCalls.callerNumber -match "\*") {

                $matchingPstnCall = $pstnCalls | Where-Object { $($callOrganizer).Replace('+', '') -like "*$($_.callerNumber.Replace('*', '').Replace('+',''))*" -and $_.userId -in $sessions.value.caller.identity.user.id }

            }

            else {

                $matchingPstnCall = $pstnCalls | Where-Object { $_.callerNumber -eq $callOrganizer }

            }

        }

        if ($matchingPstnCall.Count -gt 1) {

            $matchingPstnCall = $matchingPstnCall | Sort-Object -Property callerNumber -Unique

        }

    }

    else {

        $matchingPstnCall = $pstnCalls

    }

    if (!$matchingPstnCall) {

        Write-Output "Internal call, no matching PSTN call"
        
        $topLevelResourceAccountSession = ($sessions.value | Sort-Object startDateTime | Select-Object -First 1).caller.identity.user

        if (!$topLevelResourceAccountSession.displayName) {

            $topLevelResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($topLevelResourceAccountSession.id)" -ContentType "application/json" -OutputType PSObject)

            $topLevelResourceAccountId = $topLevelResourceAccount.id
            $topLevelResourceAccount = $topLevelResourceAccount.displayName

        }

        else {

            $topLevelResourceAccount = $topLevelResourceAccountSession.displayName
            $topLevelResourceAccountId = $topLevelResourceAccountSession.id

        }


        $matchingPstnCall = @{ userId = $topLevelResourceAccountId }

        $resourceAccountNumber = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($topLevelResourceAccountId)" -ContentType "application/json").businessPhones

        $calleeNumber = $resourceAccountNumber[0]

    }

    else {

        Write-Output "Matching PSTN call:"
        $matchingPstnCall

        $topLevelResourceAccount = $matchingPstnCall.userDisplayName
        $topLevelResourceAccountId = $matchingPstnCall.userId

        $calleeNumber = $matchingPstnCall.calleeNumber

    }


    $calleeSideParticipants = $sessions.value.callee.identity.user
    $callerSideParticipants = $sessions.value.caller.identity.user

    Write-Output "Call record participants:"
    $call.participants.user | Format-Table

    # Write-Output "Callee side participants:"
    # $calleeSideParticipants | Format-Table

    $calleeSideParticipants = $sessions.value.callee.identity.user | Sort-Object -Property id -Unique

    Write-Output "Callee side participants (unique):"
    $calleeSideParticipants | Format-Table

    Write-Output "Caller side participants:"
    $callerSideParticipants | Format-Table

    $bothSidesParticipants = $callerSideParticipants | Where-Object { $calleeSideParticipants.id -contains $_.id }

    Write-Output "Both sides participants:"
    $bothSidesParticipants | Format-Table

    $conferencingVirtualAssistantCount = ($callerSideParticipants | Where-Object { $_.id -eq "9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e" }).Count

    $calleeSideParticipantsInCallerSideParticipants = @()

    foreach ($calleeSideParticipant in $calleeSideParticipants) {

        if ($calleeSideParticipant.id -in $callerSideParticipants.id) {

            $calleeSideParticipantsInCallerSideParticipants += $calleeSideParticipant

        }

    }

    $calleeSideParticipantsVoiceAppTypes = @()

    $preCheckAnsweredSessions = @()

    $preCheckAgentSessions = $sessions.value | Where-Object { $_.caller.identity.user.displayName -ne $null -and $_.caller.identity.user.id -ne $callOrganizerId }

    foreach ($session in $preCheckAgentSessions) {

        # Alternative: $session.value.callee.userAgent.role

        switch -regex ($session.callee.userAgent.headerValue) {
            "CallQueue" {
                Write-Output "$($session.callee.identity.user.displayName) is type: Call Queue"
                $calleeSideParticipantsVoiceAppTypes += "CallQueue"
            }
            "AutoAttendant" {
                Write-Output "$($session.callee.identity.user.displayName) is type: Auto Attendant"
                $calleeSideParticipantsVoiceAppTypes += "AutoAttendant"
            }
            Default {}
        }

        if ($null -ne $session.caller.identity.user.displayName -and $session.callee.userAgent.role -notin @("skypeForBusinessAutoAttendant", "skypeForBusinessCallQueues")) {

            # Azure Automation retrieves the session start and end date times as strings
            $sessionEndDateTime = (Get-Date -Date $session.endDateTime).ToUniversalTime()
            $sessionStartDateTime = (Get-Date -Date $session.startDateTime).ToUniversalTime()

            $sessionDuration = $($sessionEndDateTime) - $($sessionStartDateTime)

            if ($sessionStartDateTime -eq $sessionEndDateTime) {

                Write-Output "Session duration: $sessionDuration (no answered calls detected)"

            }

            else {

                Write-Output "Session duration: $sessionDuration"

                $preCheckAnsweredSessions += $session

            }

        }

    }

    Write-Output "Callee side participants in caller side participants:"
    $calleeSideParticipantsInCallerSideParticipants | Format-Table

    $callRecordLastModifiedDateTime = (Get-Date -Date $call.lastModifiedDateTime).ToUniversalTime()

    $currentDateTime = (Get-Date).ToUniversalTime()

    $lastModifiedToReportTimeDifference = ($currentDateTime - $callRecordLastModifiedDateTime).TotalHours

    if ($preCheckAnsweredSessions) {

        $callRecordDataSufficientCase = "1"

        $callRecordDataSufficient = $true

    }

    elseif ($calleeSideParticipants.id.Count -eq 1 -and $calleeSideParticipantsInCallerSideParticipants.id.Count -eq 1) {

        $callRecordDataSufficientCase = "2"

        $callRecordDataSufficient = $true
        
    }

    elseif ($calleeSideParticipants.id.Count -eq 2 -and $calleeSideParticipantsVoiceAppTypes -contains "AutoAttendant" -and $calleeSideParticipantsVoiceAppTypes -contains "CallQueue" `
            -and $conferencingVirtualAssistantCount -ge $calleeSideParticipants.id.Count -and $call.version -ge 3) {

        $callRecordDataSufficientCase = "3"

        $callRecordDataSufficient = $true

    }

    elseif ($calleeSideParticipants.id.Count -eq 2 -and $calleeSideParticipantsVoiceAppTypes -contains "AutoAttendant" -and $calleeSideParticipantsVoiceAppTypes -contains "CallQueue" `
            -and $conferencingVirtualAssistantCount -ge $calleeSideParticipants.id.Count -and $call.version -ge 2) {

        Write-Output "Fetching historical call records for resource account..."

        $historicalCallsForResourceAccount = (Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/beta/communications/callRecords?`$filter=participants_v2/any(p:p/id eq '$($matchingPstnCall.userId)')" -ContentType "application/json" -OutputType PSObject).value

        $historicalSessionsForResourceAccount = @()

        $calleeSideParticipantsQueue = $calleeSideParticipants | Where-Object { $_.id -ne $($matchingPstnCall.userid) }

        $historicalCallCounter = 1

        foreach ($historicalCall in $historicalCallsForResourceAccount) {

            # Write-Output "Fetching sessions for historical call $($historicalCallCounter)/$($historicalCallsForResourceAccount.Count)..."

            $historicalSessions = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/communications/callRecords/$($historicalCall.id)/sessions" -ContentType "application/json" -OutputType PSObject).value
        
            if ($historicalSessions.Callee.identity.user.Count -ge 3) {

                $historicalSessionsForResourceAccount += $historicalSessions | Where-Object { $_.callee.identity.user.id -eq ($calleeSideParticipantsQueue.id) }

            }

            $historicalCallCounter ++
            
        }

        $historicalSessionsQueueTimes = @()

        foreach ($historicalSession in $historicalSessionsForResourceAccount) {

            $historicalSessionStartDateTime = (Get-Date -Date $historicalSession.startDateTime).ToUniversalTime()
            $historicalSessionEndDateTime = (Get-Date -Date $historicalSession.endDateTime).ToUniversalTime()

            $historicalSessionsQueueTimes += ($historicalSessionEndDateTime - $historicalSessionStartDateTime).TotalSeconds

        }

        $currentSessionForFirstQueue = ($sessions.value | Where-Object { $_.callee.identity.user.id -eq $calleeSideParticipantsQueue.id })

        $currentSessionForFirstQueueStartDateTime = (Get-Date -Date $currentSessionForFirstQueue.startDateTime).ToUniversalTime()
        $currentSessionForFirstQueueEndDateTime = (Get-Date -Date $currentSessionForFirstQueue.endDateTime).ToUniversalTime()

        $currentSessionForFirstQueueTime = ($currentSessionForFirstQueueEndDateTime - $currentSessionForFirstQueueStartDateTime).TotalSeconds

        # Find the minimum and maximum values in the array
        $minValue = $historicalSessionsQueueTimes | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
        $maxValue = $historicalSessionsQueueTimes | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

        if ($currentSessionForFirstQueueTime -ge $minValue -and $currentSessionForFirstQueueTime -le $maxValue) {

            Write-Output "Session for first queue '$($calleeSideParticipantsQueue.displayName)' reached queue time out but there is no overflow queue present in callee side participants. Expecting an updated call record. Call record is incomplete..."

            $callRecordDataSufficientCase = "4.1"

            $callRecordDataSufficient = $false

        }

        else {

            Write-Output "Session for first queue '$($calleeSideParticipantsQueue.displayName)' did not reach queue time out. Not expecting an updated call record with an overflow queue in callee side participants. Call record is complete..."

            $callRecordDataSufficientCase = "4.2"

            $callRecordDataSufficient = $true

        }

        Write-Output "Session duration for first queue: $currentSessionForFirstQueueTime. Historical session durations for first queue with more than 3 callee side participants: min: $minValue, max: $maxValue."

    }

    # In some cases, the CQ is not present in caller side participants and the call record doesn't get higher than v1. This happens for calls where an AA forwarded to CQ but not to an overflow CQ as well. If the call record did not get above v1 for more than 1 hour it's considered complete
    elseif ($calleeSideParticipants.id.Count -eq 2 -and $calleeSideParticipantsVoiceAppTypes -contains "AutoAttendant" -and $calleeSideParticipantsVoiceAppTypes -contains "CallQueue" `
            -and $conferencingVirtualAssistantCount -ge $calleeSideParticipants.id.Count -and $call.version -eq 1 -and $lastModifiedToReportTimeDifference -gt 1) {

        $callRecordDataSufficientCase = "5"

        Write-Output "Call record is missing CQ in caller side participants but there hasn't been an update to v1 for more than 1 hour. Processing call record..."

        $callRecordDataSufficient = $true

    }

    elseif ($calleeSideParticipants.id.Count -eq 2 -and $calleeSideParticipantsVoiceAppTypes -notcontains "AutoAttendant" -and $calleeSideParticipantsVoiceAppTypes -contains "CallQueue" `
            -and $calleeSideParticipantsInCallerSideParticipants.id.Count -eq 2 `
            -and $conferencingVirtualAssistantCount -ge $calleeSideParticipants.id.Count -and $call.version -ge 2) {

        $callRecordDataSufficientCase = "6"

        $callRecordDataSufficient = $true

    }

    elseif ($calleeSideParticipants.id.Count -gt 2 -and $calleeSideParticipantsInCallerSideParticipants.id.Count -ge 2) {

        $callRecordDataSufficientCase = "7"

        $callRecordDataSufficient = $true

    }

    else {

        # Add record to list as unknown
        if ($calleeSideParticipants.id.Count -eq 2 -and $calleeSideParticipantsVoiceAppTypes -contains "AutoAttendant" -and $calleeSideParticipantsVoiceAppTypes -contains "CallQueue" `
                -and $conferencingVirtualAssistantCount -ge $calleeSideParticipants.id.Count -and $call.version -eq 1) {

            # Write-Output "Call record is incomplete..." 
            # Write-Output "Both AA and CQ are present on callee side but only 1 voice app is present on caller side."

            $callRecordDataSufficientCase = "8.1"

            $callRecordDataSufficient = $false

        }

        elseif ($call.lastModifiedDateTime -and $lastModifiedToReportTimeDifference -ge 4) {
            
            # Call record is not updated for more than 4 hours, so it is considered complete

            $callRecordDataSufficientCase = "8.2"

            $callRecordDataSufficient = $true

        }

        else {

            # Write-Output "Call record is incomplete..."

            $callRecordDataSufficientCase = "8.3"

            $callRecordDataSufficient = $false

        }

    }

    if ($callRecordDataSufficient -eq $true -or $callRecordDataSufficient -eq $false) {

        Write-Output "Call record has enough data to determine missed/answered: $callRecordDataSufficient"
        Write-Output "Call record data sufficient case: $callRecordDataSufficientCase"

        $voiceAppSessions = $callerSideParticipants | Where-Object { $_.displayName -eq $null -and $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($matchingPstnCall.userId)) }

        if (!$voiceAppSessions.id) {

            $voiceAppSessionsWithoutTopLevelResourceAccount = $calleeSideParticipants | Where-Object { $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($matchingPstnCall.userId)) }

            if ($voiceAppSessionsWithoutTopLevelResourceAccount.Id.Count -gt 1) {

                $voiceAppSessionsStartTimes = ($sessions.value | Where-Object { $_.callee.identity.user.id -in $voiceAppSessionsWithoutTopLevelResourceAccount.id } | Sort-Object -Property startDateTime)[-1]

                $finalResourceAccount = $voiceAppSessionsStartTimes.callee.identity.user.displayName
                $finalResourceAccountId = $voiceAppSessionsStartTimes.callee.identity.user.id

                if ($finalResourceAccountId -and !$finalResourceAccount) {

                    $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

                }

            }

            else {

                $finalResourceAccount = $voiceAppSessionsWithoutTopLevelResourceAccount.displayName
                $finalResourceAccountId = $voiceAppSessionsWithoutTopLevelResourceAccount.id

                if ($finalResourceAccountId -and !$finalResourceAccount) {

                    $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

                }

            }

        }

        else {

            $finalResourceAccount = ($calleeSideParticipants | Where-Object { $_.id -eq $voiceAppSessions.id }).displayName
            $finalResourceAccountId = ($calleeSideParticipants | Where-Object { $_.id -eq $voiceAppSessions.id }).id

            if ($finalResourceAccountId -and !$finalResourceAccount) {

                $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

            }

        }

        Write-Output "Top level voice app (resource account) name: $topLevelResourceAccount"
        Write-Output "Final voice app (resource account) name: $finalResourceAccount"

        $callerNumber = $callOrganizer

        # Check agent sessions to determine if call was answered
        $agentSessions = $sessions.value | Where-Object { $_.caller.identity.user.displayName -ne $null -and $_.caller.identity.user.id -ne $callOrganizerId }

        # Call never went into a queue
        if (!$agentSessions) {

            Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not forwarded to any call queue."

            Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not answered by an agent"

            $answeredSessions = $null

        }

        # Call went into a queue
        else {

            $answeredSessions = @()
            $missedSessions = @()

            foreach ($session in $agentSessions) {

                # Azure Automation retrieves the session start and end date times as strings
                $sessionEndDateTime = (Get-Date -Date $session.endDateTime).ToUniversalTime()
                $sessionStartDateTime = (Get-Date -Date $session.startDateTime).ToUniversalTime()

                $sessionDuration = $($sessionEndDateTime) - $($sessionStartDateTime)

                if ($sessionStartDateTime -eq $sessionEndDateTime) {

                    $missedSessions += $session

                    Write-Output "Session duration: $sessionDuration"

                }

                else {

                    $answeredByAgent = $session.caller.identity.user.displayName | Where-Object { $_ -ne $null }

                    $answeredByAgentId = $session.caller.identity.user.id | Where-Object { $_ -ne $null }

                    Write-Output "Session duration: $sessionDuration"

                    $answeredSessions += $session

                }

            }

        }

        if ($answeredSessions) {

            if ($answeredSessions.Count -gt 1) {

                $answeredSession = ($answeredSessions | Sort-Object -Property startDateTime -Descending)[-1]

            }

            else {

                $answeredSession = $answeredSessions

            }

            $answeredSessionStartDateTime = (Get-Date -Date $answeredSession.startDateTime).ToUniversalTime()
            $answeredSessionEndDateTime = (Get-Date -Date $answeredSession.endDateTime).ToUniversalTime()

            $timeInQueue = ($callStartDateTime - $answeredSessionStartDateTime).ToString("hh\:mm\:ss")

            $netCallDuration = ($answeredSessionEndDateTime - $answeredSessionStartDateTime).ToString("hh\:mm\:ss")

            Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was answered by an agent: $answeredByAgent"

            $result = "Answered"

            $callDuration = ($callEndDateTime - $callStartDateTime).ToString("hh\:mm\:ss")

            $answeredPlatform = $answeredSession.caller.userAgent.platform

        }

        else {

            Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not answered by an agent"

            $result = "Missed"

            $answeredByAgent = $null

            $callDuration = ($callEndDateTime - $callStartDateTime).ToString("hh\:mm\:ss")

            $timeInQueue = $callDuration

            $netCallDuration = "00:00:00"

            $answeredPlatform = $null

        }

        $callSummary = [PSCustomObject]@{
            IsInternalCall = $callerIsInternalCaller
            CallType = $call.type
            CallerNumber = $callerNumber
            CalledVoiceApp = $topLevelResourceAccount
            AllInvolvedVoiceApps = (($sessions.value | Where-Object { $_.callee.userAgent.Role -in @("skypeForBusinessCallQueues", "skypeForBusinessAutoAttendant") -and $_.callee.identity.user.displayName -ne "" }).callee.identity.user.displayName | Sort-Object startDateTime | Sort-Object -Unique) -join "; "
            FinalVoiceApp = $finalResourceAccount
            CalledNumber = $calleeNumber
            StartDateTime = $callStartDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
            EndDateTime = $callEndDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
            Result = $result
            AnsweredBy = $answeredByAgent
            Platform = $answeredPlatform
            CallOfferedTo   = ($agentSessions.caller.identity.user.displayname | Sort-Object -Unique) -join "; " # To do: sort by start date time
            CallDuration = $callDuration
            NetCallDuration = $netCallDuration
            QueueDuration = $timeInQueue
            CallId = $call.id
            CallRecordVersion = $call.version
            CallRecordLastModifiedDateTime = $callRecordLastModifiedDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
            CallRecordDataSufficient = $callRecordDataSufficient
            CallRecordDataSufficientCase = $callRecordDataSufficientCase
        }

        Write-Host "Call summary:" -ForegroundColor Cyan

        $callSummary

        $callSummaries += $callSummary

    }

}