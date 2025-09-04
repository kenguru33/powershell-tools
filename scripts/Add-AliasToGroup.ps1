<#
.SYNOPSIS
  Add a user (resolved by alias) to an Entra ID security group by name.

.USAGE
  .\Add-AliasToGroup.ps1 alias@contoso.com "RS-Oslo-Crew"

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  (Permissions) Group.ReadWrite.All, User.Read.All, Directory.Read.All
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true, Position=0)]
  [string] $Alias,

  [Parameter(Mandatory=$true, Position=1)]
  [string] $GroupName
)

# Connect to Graph
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All","Directory.Read.All" | Out-Null

# --- Resolve Group (security-enabled, exact displayName) ---
$escapedGroup = $GroupName.Replace("'","''")
$groups = @(
  Get-MgGroup -Filter "displayName eq '${escapedGroup}' and securityEnabled eq true" `
              -ConsistencyLevel eventual -All
)

if ($groups.Count -eq 0) {
  Write-Error "❌ No security-enabled group found with name '${GroupName}'."
  exit 1
}
if ($groups.Count -gt 1) {
  Write-Error "❌ Multiple groups named '${GroupName}' found. Please disambiguate."
  $groups | Select-Object Id, DisplayName, Description, MailNickname
  exit 1
}
$group = $groups[0]
Write-Host "➡️ Target group: $($group.DisplayName) (Id: $($group.Id))"

# --- Resolve User by Alias ---
$aliasEsc = $Alias.Replace("'","''")
# Include Id explicitly
$filter = @"
mail eq '${aliasEsc}' or userPrincipalName eq '${aliasEsc}' or
otherMails/any(x:x eq '${aliasEsc}') or
proxyAddresses/any(p:p eq 'SMTP:${aliasEsc}') or
proxyAddresses/any(p:p eq 'smtp:${aliasEsc}')
"@ -replace "`r?`n"," "

$users = @(
  Get-MgUser -Filter $filter `
             -Property id,userPrincipalName,mail,otherMails,proxyAddresses,displayName `
             -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
)

if ($users.Count -eq 0) {
  Write-Error "❌ No user found for alias '${Alias}'."
  exit 2
}

# Prefer first match; if it lacks Id, try to refetch by UPN/mail
$user = $users | Select-Object -First 1
if (-not $user -or [string]::IsNullOrWhiteSpace($user.Id)) {
  Write-Warning "⚠️ Match returned without Id. Attempting to refetch by UPN/mail…"
  $refKey = if ($user.UserPrincipalName) { $user.UserPrincipalName } elseif ($user.Mail) { $user.Mail } else { $Alias }
  try {
    $refetched = @( Get-MgUser -Filter "userPrincipalName eq '${refKey}' or mail eq '${refKey}'" `
                              -Property id,userPrincipalName,displayName `
                              -ConsistencyLevel eventual -All -ErrorAction Stop )
    if ($refetched.Count -gt 0 -and $refetched[0].Id) {
      $user = $refetched[0]
    }
  } catch { }
}

if (-not $user -or [string]::IsNullOrWhiteSpace($user.Id)) {
  # Last resort: show what we got to aid debugging
  $users | Select-Object DisplayName, UserPrincipalName, Mail, Id | Format-Table -AutoSize
  Write-Error "❌ Resolved user has no Id. Cannot add to group."
  exit 2
}

Write-Host "👤 User: $($user.DisplayName)  UPN: $($user.UserPrincipalName)  Id: $($user.Id)"

# --- Check existing membership to avoid duplicates ---
$current = @{}
try {
  $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
  foreach ($m in $members) { if ($m -and $m.Id) { $current[$m.Id] = $true } }
} catch { }

if ($current.ContainsKey($user.Id)) {
  Write-Host "ℹ️ User is already a member of '${GroupName}'."
  exit 0
}

# --- Add user to group ---
try {
  if (Get-Command New-MgGroupMember -ErrorAction SilentlyContinue) {
    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop | Out-Null
  } else {
    Add-MgGroupMemberByRef -GroupId $group.Id `
      -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" } `
      -ErrorAction Stop | Out-Null
  }
  Write-Host "✅ Added '$($user.UserPrincipalName)' to group '${GroupName}'."
  exit 0
} catch {
  Write-Error "❌ Failed to add user: $($_.Exception.Message)"
  exit 3
}
