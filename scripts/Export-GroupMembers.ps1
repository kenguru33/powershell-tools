<#
.SYNOPSIS
  Export members of an Entra ID (Azure AD) group to a CSV.

.DEFAULT BEHAVIOR
  - Direct members only (no nested)
  - Users only (no groups/devices/SPs)
  - One column: UserPrincipalName

.USAGE
  Export-GroupMembers.ps1 "Personalhåndbok Sjøansatte" .\members.csv
  Export-GroupMembers.ps1 "Group Name" .\members.csv -Transitive
  Export-GroupMembers.ps1 "Group Name" .\members.csv -OnlyUsers -IncludeGroups
  Export-GroupMembers.ps1 "Group Name" .\out.csv -Domains rs.no contoso.com
  Export-GroupMembers.ps1 "Group Name" .\out.csv -Select Type,DisplayName,UserPrincipalName,Mail,Id
  Export-GroupMembers.ps1 -GroupId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" .\members.csv

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  Scopes: Group.Read.All, Directory.Read.All
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName='ByName')]
    [string] $GroupName,

    [Parameter(Mandatory = $true, ParameterSetName='ById')]
    [string] $GroupId,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $CsvPath,

    [switch] $Transitive,
    [switch] $OnlyUsers,
    [switch] $IncludeGroups,
    [switch] $AllTypes,
    [string[]] $Domains,
    [string[]] $Select,
    [switch] $Append,
    [switch] $Force
)

# ----- Defaults -----
if (-not $Select -or $Select.Count -eq 0) { $Select = @('UserPrincipalName') }

$validCols = @('Type','DisplayName','UserPrincipalName','Mail','Id')
$Select = $Select | ForEach-Object { $_.Trim() } | Where-Object { $_ }
foreach ($c in $Select) { if ($validCols -notcontains $c) { throw "Invalid column '$c'. Valid: $($validCols -join ', ')" } }

# Default type behavior: users-only unless caller overrides
$effectiveOnlyUsers    = $true
$effectiveIncludeGroup = $false
$effectiveAllTypes     = $false
if ($AllTypes) {
    $effectiveAllTypes = $true
    $effectiveOnlyUsers = $false
    $effectiveIncludeGroup = $false
} else {
    if ($PSBoundParameters.ContainsKey('OnlyUsers'))    { $effectiveOnlyUsers    = [bool]$OnlyUsers }
    if ($PSBoundParameters.ContainsKey('IncludeGroups')){ $effectiveIncludeGroup = [bool]$IncludeGroups }
}

# Connect
Connect-MgGraph -Scopes "Group.Read.All","Directory.Read.All" | Out-Null

# Resolve group
switch ($PSCmdlet.ParameterSetName) {
  'ByName' {
    $escaped = $GroupName.Replace("'","''")
    $groups = @( Get-MgGroup -Filter "displayName eq '${escaped}'" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue )
    if ($groups.Count -eq 0) { throw "No group found with name '${GroupName}'." }
    if ($groups.Count -gt 1) { Write-Warning "Multiple groups named '${GroupName}' found. Using the first match shown:"; $groups | Select-Object Id,DisplayName | Format-Table -AutoSize }
    $group = $groups[0]
  }
  'ById' {
    try { $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop } catch { throw "No group found with Id '${GroupId}'." }
  }
}

Write-Host "➡️ Exporting members from: $($group.DisplayName) (Id: $($group.Id))"

# Fetch members
$members = if ($Transitive) {
  @( Get-MgGroupTransitiveMember -GroupId $group.Id -All -ErrorAction SilentlyContinue )
} else {
  @( Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue )
}
Write-Host "ℹ️ Retrieved members (raw): $($members.Count)"

function Get-GraphType {
  param($obj)
  $odata = $obj.AdditionalProperties['@odata.type']
  if ($odata) { return ($odata -replace '#microsoft.graph.','') }
  foreach ($t in $obj.PSObject.TypeNames) {
    if ($t -match 'MicrosoftGraph(.*)$') { return ($Matches[1] -replace '^MicrosoftGraph','' -replace '^Microsoft','') }
  }
  return 'directoryObject'
}

# Domain filter normalization
$domainMatchers = @()
if ($Domains -and $Domains.Count -gt 0) {
    $domainMatchers = $Domains | ForEach-Object {
        $d = $_.Trim(); if (-not $d) { return $null }; if ($d.StartsWith('@')) { $d.ToLowerInvariant() } else { "@$($d.ToLowerInvariant())" }
    } | Where-Object { $_ }
}

# Identify types & collect user IDs for enrichment
$userShell = @()
$otherShell = @()

foreach ($m in $members) {
  $type = Get-GraphType $m
  if ($effectiveAllTypes) {
    if ($type -eq 'user') { $userShell += $m } else { $otherShell += $m }
  } else {
    if ($effectiveOnlyUsers -and -not $effectiveIncludeGroup) { if ($type -eq 'user') { $userShell += $m } }
    elseif ($effectiveIncludeGroup -and -not $effectiveOnlyUsers) { if ($type -eq 'group') { $otherShell += $m } }
    elseif ($effectiveIncludeGroup -and $effectiveOnlyUsers) { if ($type -in @('user','group')) { if ($type -eq 'user') { $userShell += $m } else { $otherShell += $m } } }
  }
}

# Enrich users: fetch full user objects to get reliable UPN/Mail
$userMap = @{}
foreach ($u in $userShell) {
  if (-not $u.Id) { continue }
  try {
    $fu = Get-MgUser -UserId $u.Id -Property id,displayName,userPrincipalName,mail -ErrorAction Stop
    if ($fu) { $userMap[$u.Id] = $fu }
  } catch { }
}

# Build rows (users from enriched map; others from directoryObject)
$rows = @()

# Users
foreach ($u in $userShell) {
  $full = if ($u.Id -and $userMap.ContainsKey($u.Id)) { $userMap[$u.Id] } else { $u }
  # Domain filter (if any)
  if ($domainMatchers.Count -gt 0) {
    $upn  = if ($full.UserPrincipalName) { $full.UserPrincipalName.ToLowerInvariant() } else { "" }
    $mail = if ($full.Mail)              { $full.Mail.ToLowerInvariant() } else { "" }
    $match = ($upn -and ($domainMatchers | Where-Object { $upn.EndsWith($_) })) -or
             ($mail -and ($domainMatchers | Where-Object { $mail.EndsWith($_) }))
    if (-not $match) { continue }
  }
  $rows += [pscustomobject]@{
    Type              = 'user'
    DisplayName       = $full.DisplayName
    UserPrincipalName = $full.UserPrincipalName
    Mail              = $full.Mail
    Id                = $full.Id
  }
}

# Non-users (groups, etc.) if included
foreach ($o in $otherShell) {
  # Domain filter only applies when we have UPN/Mail; most non-users won't match anyway
  if ($domainMatchers.Count -gt 0) {
    $upn  = if ($o.UserPrincipalName) { $o.UserPrincipalName.ToLowerInvariant() } else { "" }
    $mail = if ($o.Mail)              { $o.Mail.ToLowerInvariant() } else { "" }
    $match = ($upn -and ($domainMatchers | Where-Object { $upn.EndsWith($_) })) -or
             ($mail -and ($domainMatchers | Where-Object { $mail.EndsWith($_) }))
    if (-not $match) { continue }
  }
  $rows += [pscustomobject]@{
    Type              = (Get-GraphType $o)
    DisplayName       = $o.DisplayName
    UserPrincipalName = $o.UserPrincipalName
    Mail              = $o.Mail
    Id                = $o.Id
  }
}

Write-Host "ℹ️ Rows after filters & enrichment: $($rows.Count)"

# Project only selected columns (default = UPN only)
$projected = $rows | Select-Object -Property $Select
Write-Host "ℹ️ Rows after projection: $($projected.Count)"

# Export
$csvFullPath = if ([IO.Path]::IsPathRooted($CsvPath)) { $CsvPath } else { Join-Path -Path (Get-Location) -ChildPath $CsvPath }
$dir = Split-Path -Parent -Path $csvFullPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

if ($Append) {
  if ($projected -and $projected.Count -gt 0) {
    $projected | Export-Csv -Path $csvFullPath -NoTypeInformation -Encoding UTF8 -Append
  } else {
    Write-Warning "Nothing to append. No members matched the filters/columns."
    if (-not (Test-Path -LiteralPath $csvFullPath)) { ($Select -join ',') | Out-File -FilePath $csvFullPath -Encoding UTF8 }
  }
} else {
  if ($projected -and $projected.Count -gt 0) {
    $projected | Export-Csv -Path $csvFullPath -NoTypeInformation -Encoding UTF8 -Force:$Force
  } else {
    Write-Warning "No rows to export with selected columns: $($Select -join ', ')"
    ($Select -join ',') | Out-File -FilePath $csvFullPath -Encoding UTF8 -Force:$Force
  }
}

Write-Host "✅ Exported $($projected.Count) row(s) to '$csvFullPath'"
