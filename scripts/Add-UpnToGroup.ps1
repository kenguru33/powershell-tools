<#
.SYNOPSIS
  Add a user (by UPN) to an Entra ID security group by name.

.USAGE
  .\Add-UpnToGroup.ps1 user.upn@tenant.com "My-Security-Group"
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Upn,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $GroupName
)

# Connect to Graph (requires Group.ReadWrite.All, User.Read.All)
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All","Directory.Read.All" | Out-Null

# Find group
$escaped = $GroupName.Replace("'","''")
$g = @( Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All )
if ($g.Count -eq 0) { Write-Error "❌ Group not found: ${GroupName}"; exit 1 }
if ($g.Count -gt 1) { Write-Error "❌ Multiple groups named '${GroupName}' found."; exit 1 }
$group = $g[0]

Write-Host "➡️ Target group: $($group.DisplayName) (Id: $($group.Id))"

# Find user by UPN
$upnEsc = $Upn.Replace("'","''")
$u = @( Get-MgUser -Filter "userPrincipalName eq '${upnEsc}'" -All -ErrorAction SilentlyContinue ) | Select-Object -First 1

if (-not $u -or -not $u.Id) {
    Write-Error "❌ No user found with UPN '${Upn}'"
    exit 2
}

# Check if already member
$current = @{}
try {
    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
    foreach ($m in $members) { if ($m -and $m.Id) { $current[$m.Id] = $true } }
} catch { }

if ($current.ContainsKey($u.Id)) {
    Write-Host "ℹ️ User '${Upn}' is already a member of '${GroupName}'"
    exit 0
}

# Add to group
try {
    if (Get-Command New-MgGroupMember -ErrorAction SilentlyContinue) {
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $u.Id -ErrorAction Stop | Out-Null
    } else {
        Add-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($u.Id)" } -ErrorAction Stop | Out-Null
    }
    Write-Host "✅ Added '${Upn}' to group '${GroupName}'"
} catch {
    Write-Error "❌ Failed to add '${Upn}' to '${GroupName}': $($_.Exception.Message)"
    exit 3
}
