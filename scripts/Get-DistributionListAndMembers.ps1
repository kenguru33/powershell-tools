<#
.SYNOPSIS
  Get an Exchange Online Distribution List (DL) and list its members.
  Also shows whether the group (and owners) are OnPrem or Cloud.

.USAGE
  .\Get-DistributionListAndMembers.ps1 "Alle Sjøansatte"
  .\Get-DistributionListAndMembers.ps1 "alle.sjoansatte@rs.no"
  .\Get-DistributionListAndMembers.ps1 "Alle Sjøansatte" -CsvPath .\members.csv
  .\Get-DistributionListAndMembers.ps1 "Alle Sjøansatte" -ShowOwners
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory, Position=0)]
  [string] $Identity,

  [string] $CsvPath,

  [switch] $ShowOwners
)

# ---------- Ensure Exchange Online module / connection ----------
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
  throw "ExchangeOnlineManagement module is required. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

try {
  if (-not (Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }
} catch { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }

# ---------- Resolve the DL ----------
try {
  $dl = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
} catch {
  throw ("Distribution list not found or ambiguous for '{0}' -- {1}" -f $Identity, $_.Exception.Message)
}

# Helper to convert IsDirSynced -> 'OnPrem'/'Cloud'
function Get-SourceLabel([bool]$isDirSynced) {
  if ($isDirSynced) { 'OnPrem (DirSync)' } else { 'Cloud' }
}

# ---------- Print DL details ----------
Write-Host ""
Write-Host "================ DISTRIBUTION LIST ================"

$dlView = [pscustomobject]@{
  DisplayName                        = $dl.DisplayName
  PrimarySmtpAddress                 = "$($dl.PrimarySmtpAddress)"
  Alias                              = $dl.Alias
  RecipientTypeDetails               = $dl.RecipientTypeDetails
  GroupType                          = $dl.GroupType
  ManagedBy                          = ($dl.ManagedBy -join '; ')
  HiddenFromAddressListsEnabled      = $dl.HiddenFromAddressListsEnabled
  RequireSenderAuthenticationEnabled = $dl.RequireSenderAuthenticationEnabled
  MemberJoinRestriction              = $dl.MemberJoinRestriction
  MemberDepartRestriction            = $dl.MemberDepartRestriction
  WhenCreated                        = $dl.WhenCreated
  WhenChanged                        = $dl.WhenChanged
  IsDirSynced                        = $dl.IsDirSynced
  Source                             = Get-SourceLabel $dl.IsDirSynced
}

$dlView | Format-List

# ---------- Owners (optional, with source) ----------
if ($ShowOwners) {
  Write-Host ""
  Write-Host "Owners:"
  $ownerRefs = @()
  try { $ownerRefs = Get-DistributionGroup -Identity $dl.Identity | Select-Object -ExpandProperty ManagedBy } catch {}

  if (-not $ownerRefs -or $ownerRefs.Count -eq 0) {
    Write-Host "  <none>"
  } else {
    $ownerRows = foreach ($o in $ownerRefs) {
      # Resolve owner to a recipient to read IsDirSynced
      $rec = $null
      try { $rec = Get-Recipient -Identity $o -ErrorAction Stop } catch {}
      [pscustomobject]@{
        DisplayName          = if ($rec) { $rec.DisplayName } else { "$o" }
        PrimarySmtpAddress   = if ($rec) { "$($rec.PrimarySmtpAddress)" } else { $null }
        RecipientTypeDetails = if ($rec) { $rec.RecipientTypeDetails } else { $null }
        IsDirSynced          = if ($rec) { $rec.IsDirSynced } else { $null }
        Source               = if ($rec) { Get-SourceLabel $rec.IsDirSynced } else { 'Unknown' }
      }
    }
    $ownerRows | Format-Table -AutoSize
  }
}

# ---------- Members ----------
Write-Host "`nFetching members..."
$members = @()
try {
  $members = @( Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction Stop )
} catch {
  Write-Warning ("Failed to get members for '{0}' -- {1}" -f $dl.DisplayName, $_.Exception.Message)
}

if (-not $members -or $members.Count -eq 0) {
  Write-Host "Members: <none>"
} else {
  Write-Host "Members:"
  $view = $members | Select-Object `
    DisplayName,
    PrimarySmtpAddress,
    RecipientTypeDetails,
    Identity
  $view | Format-Table -AutoSize

  if ($CsvPath) {
    try {
      $exportRows = $members | ForEach-Object {
        [pscustomobject]@{
          DisplayName          = $_.DisplayName
          PrimarySmtpAddress   = ($_.PrimarySmtpAddress | ForEach-Object { "$_" })
          RecipientTypeDetails = $_.RecipientTypeDetails
          Identity             = ($_.Identity | ForEach-Object { "$_" })
        }
      }
      $dir = [System.IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $CsvPath -ErrorAction SilentlyContinue))
      if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }
      if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
      $exportRows | Export-Csv -LiteralPath $CsvPath -Encoding UTF8 -NoTypeInformation
      Write-Host ("`n✅ Exported {0} member(s) to: {1}" -f $exportRows.Count, $CsvPath)
    } catch {
      Write-Warning ("Failed to export CSV to '{0}' -- {1}" -f $CsvPath, $_.Exception.Message)
    }
  }
}

# ---------- Quick summary ----------
Write-Host ("`nSummary: '{0}' is {1} and has {2} member(s)." -f $dl.DisplayName, (Get-SourceLabel $dl.IsDirSynced), ($members | Measure-Object).Count)
