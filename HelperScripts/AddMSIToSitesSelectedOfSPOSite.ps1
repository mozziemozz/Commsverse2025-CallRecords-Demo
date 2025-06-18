$environmentVariables = Get-Content -Path ".\.local\environmentVariables.json" | ConvertFrom-Json

Connect-MgGraph -Scopes "Sites.FullControl.All","Application.Read.All","Group.Read.All"

# This is the group id of the group to which the site belongs
$groupId = $environmentVariables.GroupId

# This is the site name to find the site id
$siteName = $environmentVariables.SiteName

# Get site id
$siteId = (Get-MgGroupSite -GroupId $groupId -SiteId "root" | Where-Object { $_.DisplayName -eq $siteName }).Id

$targetServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $environmentVariables.MSIId

$params = @{
    roles               = @(
        "manage"
    )
    grantedToIdentities = @(
        @{
            application = @{
                id          = "$($targetServicePrincipal.AppId)"
                displayName = "$($targetServicePrincipal.DisplayName)"
            }
        }
    )
}

$newSitePermission = New-MgSitePermission -SiteId $siteId -BodyParameter $params

# Get Site Permissions to verify the permission was created
$sitePermissions = Get-MgSitePermission -SiteId $siteId -PermissionId $newSitePermission.Id

Write-Host "Permission Id: $($newSitePermission.Id)"

Write-Host "Permission Roles: $($sitePermissions.roles)"

Write-Host "Permission Granted To Identities:"

$sitePermissions.grantedToIdentities.application
