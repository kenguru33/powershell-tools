#requires -Modules ExchangeOnlineManagement
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [Parameter(Position=0, Mandatory=$true)]
  [string]$Identity,

  [switch]$Force
)

function Ensure-EXO {
  # Make sure the module is installed & importable in THIS pwsh
  if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

function Connect-EXO-WithFallback {
  # First try REST (no remote PS session). Then verify cmdlets exist.
  try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$false
  } catch {
    Write-Warning "REST connect failed or not available, falling back to RPS..."
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$true
  }

  # If Get-DistributionGroup still isn't present, force RPS mode and re-connect
  if (-not (Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue)) {
    Write-Warning "Get-DistributionGroup not found after REST connect; reconnecting with RPS..."
    try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$true
  }

  if (-not (Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue)) {
    throw "Exchange cmdlets still unavailable in this shell. You may be using a different pwsh than the one with the module. See notes printed below."
  }
}

try {
  Ensure-EXO
  Connect-EXO-WithFallback

  # Resolve the group
  $group = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
  if (-not $group) {
    throw "Distribution group '$Identity' not found."
  }

  Write-Host "Found: $($group.DisplayName) <$($group.PrimarySmtpAddress)>" -ForegroundColor Cyan

  # Count members
  $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue
  $count   = $members.Count

  if ($count -gt 0 -and -not $Force) {
    Write-Warning "Group has $count member(s). Aborting delete. Use -Force to override."
    return
  }

  if ($PSCmdlet.ShouldProcess($group.DisplayName, "Remove-DistributionGroup")) {
    Remove-DistributionGroup -Identity $group.Identity -Confirm:$false -ErrorAction Stop
    Write-Host "✅ Deleted: $($group.DisplayName)" -ForegroundColor Green
  }
}
catch {
  Write-Error "❌ Failed: $($_.Exception.Message)"
  Write-Host "`nDiagnostics:" -ForegroundColor Yellow
  Write-Host (" pwsh path         : {0}" -f (Get-Command pwsh).Source)
  Write-Host (" PSVersion         : {0}" -f $PSVersionTable.PSVersion)
  Write-Host (" Module available? : {0}" -f [bool](Get-Module ExchangeOnlineManagement -ListAvailable))
  Write-Host (" Module loaded?    : {0}" -f [bool](Get-Module ExchangeOnlineManagement))
  Write-Host (" Cmdlet visible?   : {0}" -f [bool](Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue))
  Write-Host " PSModulePath:"; $env:PSModulePath -split [IO.Path]::PathSeparator | ForEach-Object { "  - $_" }
}
finally {
  try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
}
