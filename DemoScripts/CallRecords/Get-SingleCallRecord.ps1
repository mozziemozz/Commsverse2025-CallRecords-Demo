. .\Modules\SecureCredsMgmt.ps1

$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

. Get-MZZSecureCreds -FileName "CommsverseDemo2025"
$appSecret = $passwordEncrypted

# Create new powershell credential object from app id (user name) and app secret (password)
$clientSecretCredential = New-Object System.Management.Automation.PSCredential ($($environmentVariables.CallRecordsReadAllAppId), $passwordEncrypted)

# Connect to Graph
Connect-MgGraph -ClientSecretCredential $clientSecretCredential -TenantId $environmentVariables.TenantId -NoWelcome

$mgContext = Get-MgContext | Select-Object ClientId, AppName, AuthType, TokenCredentialType, Scopes

Write-Host "Graph Context:" -ForegroundColor Green

$mgContext | Format-List

$callRecordId = "8ae62ca5-c0bf-438d-9303-2a18b1ddc43a"

$callRecord = Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/communications/callRecords/$($callRecordId)" -ContentType "application/json" -OutputType PSObject

# Write-Host "This is the raw call record object (duration is not a part of it):" -ForegroundColor Green

# $callRecord | fl * # Show all properties of the call record

Write-Host "This is a processed call record object with some additional properties (duration, participants phone numbers and users):" -ForegroundColor Green

$callRecord | fl id, version, startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
@{Name="ParticipantsPhone";Expression={$_.participants.phone.id}}, `
@{Name="ParticipantsUsers";Expression={$_.participants.user.displayName -join "; "}}

# Get all sessions of the call record
$callRecordSessions = Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/communications/callRecords/$($callRecordId)/sessions" -ContentType "application/json" -OutputType PSObject | Select-Object -ExpandProperty value

# Filter sessions where the caller is a user and the session has a duration greater than 0 seconds (answered session by an agent)
$answeredAgentSession = $callRecordSessions | Where-Object { $null -ne $_.caller.identity.user.displayName -and ($_.endDateTime - $_.startDateTime).TotalSeconds -gt 0 }

# Get all segments of the sessions
$callRecordSessionsSegments = Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/communications/callRecords/$($callRecordId)/sessions?expand=segments" -ContentType "application/json" -OutputType PSObject | Select-Object -ExpandProperty value
