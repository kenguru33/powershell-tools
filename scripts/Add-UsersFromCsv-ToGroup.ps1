<#
.SYNOPSIS
  Add users from a CSV to ANY Entra ID group (AAD Security, M365 Group, classic Distribution List),
  matching STRICTLY by UPN (no display name / alias / proxy fallback).

.RULES
  - Default CSV behavior: use column 0 (first column).
  - If your CSV has multiple columns and you want a different one, pass -ColumnName.
  - Input values MUST be valid UPNs (email-like). Non-email rows are skipped.

.USAGE
  .\Add-UsersFromCsv-ToGroup.ps1 "My Group" .\members.csv
  .\Add-UsersFromCsv-ToGroup.ps1 "My Group" .\members.csv -ColumnName UserPrincipalName

.REQUIREMENTS
  - Microsoft Graph PowerShell SDK
      Install-Module Microsoft.Graph -Scope CurrentUser
  - For classic Distribution Lists:
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
  - Permissions: Group.ReadWrite.All, User.Read.All, Directory.Read.All
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
  [Parameter(Mandatory=$true, ParameterSetName='ByName', Position=0)]
  [string] $GroupName,

  [Parameter(Mandatory=$true, ParameterSetName='ById', Position=0)]
  [string] $GroupId,

  [Parameter(Mandatory=$true, Position=1)]
  [ValidateScript({ Test-Path $_ })]
  [string] $CsvPath,

  [string] $ColumnName
)

# ========== Helpers: strict UPN resolvers ==========
function Resolve-GraphUserByUpn {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Upn)

  $u = $Upn.Trim()
  if ($u -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $null }

  $esc = $u.Replace("'","''")
  $users = @( Get-MgUser -Filter "userPrincipalName eq '${esc}'" -All -ErrorAction SilentlyContinue )
  if ($users.Count -eq 1) { return $users[0] }

  if ($users.Count -gt 1) {
    $exact = $users | Where-Object { $_.UserPrincipalName -ieq $u }
    if ($exact.Count -ge 1) { return $exact[0] }
  }
  return $null
}

function Resolve-EXORecipientByUpn {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Upn,
    [switch] $ExcludeGuests
  )

  $u = $Upn.Trim()
  if ($u -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $null }

  # Exact match on Exchange UserPrincipalName
  $users = @( Get-User -ResultSize Unlimited -Filter ("UserPrincipalName -eq '{0}'" -f $u) -ErrorAction SilentlyContinue )
  if ($users.Count -eq 0) { return $null }
  $user = if ($users.Count -gt 1) { ($users | Where-Object { $_.UserPrincipalName -ieq $u } | Select-Object -First 1) } else { $users[0] }
  if (-not $user) { return $null }

  $recip = Get-Recipient -Identity $user.Identity -ErrorAction SilentlyContinue
  if (-not $recip) { return $null }

  if ($ExcludeGuests -and $recip.RecipientTypeDetails -eq 'GuestMailUser') { return $null }
  return $recip
}

# ========== Connect Graph ==========
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All","Directory.Read.All" | Out-Null

# ========== Resolve group ==========
switch ($PSCmdlet.ParameterSetName) {
  'ByName' {
    $escaped = $GroupName.Replace("'","''")
    $g = @( Get-MgGroup -Filter "displayName eq '${escaped}'" -ConsistencyLevel eventual -All )
    if ($g.Count -eq 0) { throw "Group not found by name: ${GroupName}" }
    if ($g.Count -gt 1) { throw "Multiple groups named '${GroupName}' found. Use -GroupId." }
    $group = $g[0]
  }
  'ById' {
    try { $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop }
    catch { throw "Group not found by Id: ${GroupId}" }
  }
}
if (-not $group -or -not $group.Id) { throw "Resolved group object has no valid Id. Aborting." }

# Determine group type
$groupType =
  if ($group.GroupTypes -and ($group.GroupTypes -contains 'Unified')) { 'Unified' }            # M365 Group
  elseif ($group.MailEnabled -and -not $group.SecurityEnabled)        { 'DistributionList' }   # classic DL
  else                                                                { 'AadSecurity' }        # Security group

Write-Host ("Target group: {0} (Id: {1})" -f $group.DisplayName, $group.Id)
Write-Host ("Type: {0}" -f $groupType)

# ========== Load CSV (column 0 default or -ColumnName) ==========
$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows -or $rows.Count -eq 0) { throw "CSV file is empty." }

if ($ColumnName) {
  if (-not ($rows | Get-Member -Name $ColumnName -MemberType NoteProperty)) {
    throw "CSV does not contain a column named '$ColumnName'."
  }
  $getValue = { param($row) $row.$ColumnName }
  Write-Host ("Using CSV column: {0}" -f $ColumnName)
} else {
  $firstHeader = ($rows[0].PSObject.Properties | Select-Object -First 1).Name
  $getValue = { param($row) $row.$firstHeader }
  Write-Host ("Using first CSV column: {0}" -f $firstHeader)
}

# ========== Membership cache for Graph path ==========
$current = @{}
try {
  $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
  foreach ($m in $members) { if ($m -and $m.Id) { $current[$m.Id] = $true } }
} catch { }

# ========== If DL: connect EXO & cache member SMTPs ==========
if ($groupType -eq 'DistributionList') {
  if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "Classic Distribution List detected. Install ExchangeOnlineManagement: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
  }
  try {
    if (-not (Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }
  } catch { try { Connect-ExchangeOnline -ShowBanner:$false | Out-Null } catch { throw } }

  $dlIdentity = if ($group.Mail) { $group.Mail } else { $group.DisplayName }
  $dlMembers = @{}
  try {
    Get-DistributionGroupMember -Identity $dlIdentity -ResultSize Unlimited -ErrorAction SilentlyContinue |
      ForEach-Object { if ($_.PrimarySmtpAddress) { $dlMembers[$_.PrimarySmtpAddress.ToLowerInvariant()] = $true } }
  } catch { }
}

# ========== Tracking ==========
$Added=@(); $Skipped=@(); $NotFound=@(); $Invalid=@(); $Failed=@()

# ========== Process with live output ==========
$total = $rows.Count; $index = 0
foreach ($row in $rows) {
  $index++
  $val = (& $getValue $row)
  if ($null -eq $val) { continue }
  $upn = "$val".Trim('" ').Trim()
  if ([string]::IsNullOrWhiteSpace($upn)) { continue }

  Write-Progress -Activity "Adding members by UPN" -Status $upn -PercentComplete ([int](100 * $index / $total))

  # Strong guard: must be UPN (email-like)
  if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host ("❓ Skipping non-UPN value: {0}" -f $upn)
    $Invalid += $upn
    continue
  }

  if ($groupType -ne 'DistributionList') {
    # ===== Graph (AAD/M365) path — UPN only =====
    $u = Resolve-GraphUserByUpn -Upn $upn
    if (-not $u) {
      Write-Host ("❓ Not found (UPN): {0}" -f $upn)
      $NotFound += $upn
      continue
    }

    if ($current.ContainsKey($u.Id)) {
      Write-Host ("⏭️ Skipped (already member): {0}" -f $upn)
      $Skipped += $upn
      continue
    }

    try {
      New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $u.Id -ErrorAction Stop | Out-Null
      $current[$u.Id] = $true
      Write-Host ("✅ Added: {0}" -f $upn)
      $Added += $upn
    } catch {
      Write-Warning ("❌ Failed to add {0}: {1}" -f $upn, $_.Exception.Message)
      $Failed += $upn
    }
  }
  else {
    # ===== Distribution List (Exchange) path — UPN only =====
    $recip = Resolve-EXORecipientByUpn -Upn $upn
    if (-not $recip) {
      Write-Host ("❓ Not found (DL by UPN): {0}" -f $upn)
      $NotFound += $upn
      continue
    }

    $smtp = $null
    if ($recip.PrimarySmtpAddress) { $smtp = $recip.PrimarySmtpAddress.ToString().ToLowerInvariant() }
    if ($smtp -and $dlMembers.ContainsKey($smtp)) {
      Write-Host ("⏭️ Skipped (already in DL): {0}" -f $upn)
      $Skipped += $upn
      continue
    }

    # Build an unambiguous, SCALAR identifier for -Member (UPN first)
    $memberIdCandidates = @(
      ($upn),                                                   # 1) UPN from CSV (string)
      ($recip.PrimarySmtpAddress           | ForEach-Object { "$_" }), # 2) Primary SMTP (stringified)
      ($recip.ExternalDirectoryObjectId    | ForEach-Object { "$_" }), # 3) Entra/EXO GUID (stringified)
      ($recip.DistinguishedName            | ForEach-Object { "$_" }), # 4) DN (stringified)
      ($recip.Identity                     | ForEach-Object { "$_" })  # 5) Identity (stringified)
    ) | Where-Object { $_ -and ($_ -is [string]) -and $_.Trim() }

    $memberId = $memberIdCandidates | Select-Object -First 1
    if (-not $memberId) {
      Write-Warning ("❌ Failed to build a unique identifier for {0} (recipient: {1})" -f $upn, $recip.DisplayName)
      $Failed += $upn
      continue
    }

    try {
      Add-DistributionGroupMember -Identity $dlIdentity -Member $memberId -ErrorAction Stop
      if ($smtp) { $dlMembers[$smtp] = $true }
      Write-Host ("✅ Added to DL: {0}" -f $upn)
      $Added += $upn
    } catch {
      Write-Warning ("❌ Failed to add {0} to DL: {1}" -f $upn, $_.Exception.Message)
      $Failed += $upn
    }
  }
}

Write-Progress -Activity "Adding members by UPN" -Completed

# ========== Summary (counts for Added/Skipped; list only errors) ==========
Write-Host "`n=== Summary: $($group.DisplayName) ==="
Write-Host ("  Added:    {0}" -f $Added.Count)
Write-Host ("  Skipped:  {0}" -f $Skipped.Count)
Write-Host ("  NotFound: {0}" -f $NotFound.Count)
Write-Host ("  Invalid:  {0}" -f $Invalid.Count)
Write-Host ("  Failed:   {0}" -f $Failed.Count)

if ($NotFound.Count) {
  Write-Host "`n❓ Not found (UPN):"
  $NotFound | ForEach-Object { Write-Host ("   {0}" -f $_) }
}
if ($Invalid.Count) {
  Write-Host "`n⚠️ Invalid (non-UPN values):"
  $Invalid | ForEach-Object { Write-Host ("   {0}" -f $_) }
}
if ($Failed.Count) {
  Write-Host "`n❌ Failed:"
  $Failed | ForEach-Object { Write-Host ("   {0}" -f $_) }
}
