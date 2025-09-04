<#
.SYNOPSIS
  Bulk add users (resolved by alias) from a CSV to an Entra ID security group.

.DESCRIPTION
  Reads a CSV with header 'Email'. For each value:
    - Resolves against mail, userPrincipalName, otherMails, proxyAddresses
    - Adds the user to the specified *security-enabled* group
  With -StrictUpn: only add if CSV value equals the resolved user's UPN (case-insensitive).

.USAGE
  # Loose (resolve alias to user and add):
  .\Add-AliasesFromCsv-ToGroup.ps1 .\members.csv "RS-Oslo-Crew"

  # Strict (only add when CSV value == UPN):
  .\Add-AliasesFromCsv-ToGroup.ps1 .\members.csv "RS-Oslo-Crew" -StrictUpn

.CSV FORMAT
  Email
  torgeir.arntsen@rs.no
  arneil@rs.no

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  (Permissions) Group.ReadWrite.All, User.Read.All, Directory.Read.All
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true, Position=0)]
  [ValidateScript({ Test-Path $_ })]
  [string] $CsvPath,

  [Parameter(Mandatory=$true, Position=1)]
  [string] $GroupName,

  [switch] $StrictUpn
)

# Connect to Graph
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All","Directory.Read.All" | Out-Null

# --- Resolve group (security-enabled, exact displayName) ---
$escapedGroup = $GroupName.Replace("'","''")
$groups = @(
  Get-MgGroup -Filter "displayName eq '${escapedGroup}' and securityEnabled eq true" `
              -ConsistencyLevel eventual -All
)
if ($groups.Count -eq 0) { Write-Error "❌ No security-enabled group named '${GroupName}'."; exit 1 }
if ($groups.Count -gt 1) { Write-Error "❌ Multiple groups named '${GroupName}'. Use a unique name."; $groups | Select-Object Id,DisplayName,MailNickname; exit 1 }
$group = $groups[0]
Write-Host "➡️ Target group: $($group.DisplayName) (Id: $($group.Id))"

# --- Load CSV ---
$rows = Import-Csv -Path $CsvPath
if (-not $rows -or -not ($rows | Get-Member -Name Email -MemberType NoteProperty)) {
  Write-Error "❌ CSV must have a header 'Email' and at least one row."
  exit 1
}

# --- Current membership cache (avoid duplicates) ---
$current = @{}
try {
  $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
  foreach ($m in $members) { if ($m -and $m.Id) { $current[$m.Id] = $true } }
} catch { }

# --- Counters ---
[int]$added = 0; [int]$skipped = 0; [int]$notFound = 0; [int]$invalid = 0; [int]$notStrict = 0

foreach ($row in $rows) {
  $alias = ($row.Email).Trim()
  if ([string]::IsNullOrWhiteSpace($alias)) { continue }

  $aliasEsc = $alias.Replace("'","''")
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
    Write-Warning "No user found for '${alias}'"
    $notFound++
    continue
  }

  $user = $users | Select-Object -First 1
  if (-not $user -or [string]::IsNullOrWhiteSpace($user.Id)) {
    # Try a quick refetch by UPN/mail
    $refKey = if ($user.UserPrincipalName) { $user.UserPrincipalName } elseif ($user.Mail) { $user.Mail } else { $alias }
    $refetched = @( Get-MgUser -Filter "userPrincipalName eq '${refKey}' or mail eq '${refKey}'" `
                              -Property id,userPrincipalName,displayName `
                              -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue )
    if ($refetched.Count -gt 0 -and $refetched[0].Id) {
      $user = $refetched[0]
    }
  }
  if (-not $user -or [string]::IsNullOrWhiteSpace($user.Id)) {
    Write-Warning "Resolved user for '${alias}' has no Id; skipping."
    $invalid++
    continue
  }

  if ($StrictUpn -and ($alias -ine $user.UserPrincipalName)) {
    Write-Host "⛔ CSV value '${alias}' != UPN '$($user.UserPrincipalName)'; not adding (StrictUpn)."
    $notStrict++
    continue
  }

  if ($current.ContainsKey($user.Id)) {
    Write-Host "Already a member: $($user.UserPrincipalName)"
    $skipped++
    continue
  }

  try {
    if (Get-Command New-MgGroupMember -ErrorAction SilentlyContinue) {
      New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop | Out-Null
    } else {
      Add-MgGroupMemberByRef -GroupId $group.Id `
        -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" } `
        -ErrorAction Stop | Out-Null
    }
    $current[$user.Id] = $true
    $added++
    Write-Host "✅ Added: $($user.UserPrincipalName) (from '${alias}')"
  } catch {
    Write-Warning "Failed to add $($user.UserPrincipalName): $($_.Exception.Message)"
  }
}

# --- Summary ---
Write-Host "`n=== Summary: $($group.DisplayName) ==="
Write-Host "  Added:        ${added}"
Write-Host "  Skipped:      ${skipped} (already members)"
if ($StrictUpn) { Write-Host "  NotStrictUPN: ${notStrict} (CSV value != UPN)" }
Write-Host "  NotFound:     ${notFound}"
Write-Host "  Invalid:      ${invalid} (no Id)"
