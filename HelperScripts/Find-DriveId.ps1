$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

Connect-MgGraph -Scopes "Sites.FullControl.All","Group.Read.All"

# This is the group id of the group to which the site belongs
$groupId = $environmentVariables.GroupId

# This is the site name to find the site id
$siteName = $environmentVariables.SiteName

# Get site id
$siteId = (Get-MgGroupSite -GroupId $groupId -SiteId "root" | Where-Object { $_.DisplayName -eq $siteName }).Id

$driveId = Get-MgSiteDrive -SiteId $siteId | Out-GridView -Title "Select the document library to use for storing the call records" -PassThru | Select-Object -ExpandProperty Id

Write-Host "Drive ID copied to clipboard: $driveId" -ForegroundColor Green
Set-Clipboard -Value $driveId