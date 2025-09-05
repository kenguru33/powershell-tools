#requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param(
  [Parameter(Position=0, Mandatory=$true)]
  [string]$DisplayName,

  [Parameter(Position=1, Mandatory=$true)]
  [string]$MailNickname,

  [Parameter(Position=2, Mandatory=$true)]
  [string]$Domain,

  [ValidateSet('Distribution','Security')]
  [string]$Type = 'Distribution'
)

function Ensure-EXO {
  if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
  }
  Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

function Connect-EXO-WithFallback {
  # Try REST first
  try {
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$false
  } catch {
    Write-Warning "REST connect failed; trying RPS..."
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$true
  }

  # If classic creation cmdlet is missing, force RPS
  if (-not (Get-Command New-DistributionGroup -ErrorAction SilentlyContinue)) {
    Write-Warning "New-DistributionGroup not available after REST connect; reconnecting with RPS..."
    try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
    Connect-ExchangeOnline -ShowBanner:$false -UseRPSSession:$true
  }

  if (-not (Get-Command New-DistributionGroup -ErrorAction SilentlyContinue)) {
    throw "Exchange creation cmdlets are still unavailable in this shell."
  }
}

function Validate-Alias {
  param([string]$AliasRaw)

  $alias = ($AliasRaw -replace '\s','')
  if ([string]::IsNullOrWhiteSpace($alias)) { throw "MailNickname cannot be empty." }
  if ($alias -notmatch '^[A-Za-z0-9._-]+$') {
    throw "MailNickname '$AliasRaw' contains invalid characters. Use letters, digits, dot, underscore, or hyphen."
  }
  $alias = $alias.ToLowerInvariant()
  return $alias
}

function Validate-Domain {
  param([string]$Domain)
  $accepted = Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.DomainName -ieq $Domain }
  if (-not $accepted) {
    $all = (Get-AcceptedDomain | Select-Object -ExpandProperty DomainName) -join ', '
    throw "Domain '$Domain' is not an accepted domain. Accepted: $all"
  }
}

function Check-Alias-Unique {
  param([string]$Alias)
  # Prefer REST cmdlet for portability
  try {
    $r = Get-EXORecipient -RecipientTypeDetails UserMailbox,SharedMailbox,MailUser,MailContact,RoomMailbox,EquipmentMailbox,GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup,DynamicDistributionGroup -Filter "Alias -eq '$Alias'" -ErrorAction Stop
    if ($r) { throw "Alias '$Alias' already exists on '$($r.DisplayName)'." }
  } catch {
    # If REST cmdlet missing (very rare after Connect), try classic
    if (Get-Command Get-Recipient -ErrorAction SilentlyContinue) {
      $r2 = Get-Recipient -Filter "Alias -eq '$Alias'" -ErrorAction SilentlyContinue
      if ($r2) { throw "Alias '$Alias' already exists on '$($r2.DisplayName)'." }
    } else {
      Write-Warning "Could not run alias uniqueness check (Get-EXORecipient/Get-Recipient unavailable). Proceeding..."
    }
  }
}

try {
  Ensure-EXO
  Connect-EXO-WithFallback

  $alias = Validate-Alias -AliasRaw $MailNickname
  Validate-Domain -Domain $Domain
  Check-Alias-Unique -Alias $alias

  $smtp = "$alias@$Domain"
  Write-Host "Creating $Type group '$DisplayName' with alias '$alias' <$smtp>..." -ForegroundColor Cyan

  $newParams = @{
    Name               = $DisplayName
    DisplayName        = $DisplayName
    Alias              = $alias
    PrimarySmtpAddress = $smtp
    Type               = $Type   # Distribution or Security
    ErrorAction        = 'Stop'
  }

  $dg = New-DistributionGroup @newParams

  # Show result
  $dg = Get-DistributionGroup -Identity $dg.Identity
  Write-Host "✅ Created" -ForegroundColor Green
  Write-Host (" Name : {0}" -f $dg.DisplayName)
  Write-Host (" Alias: {0}" -f $dg.Alias)
  Write-Host (" SMTP : {0}" -f ($dg.PrimarySmtpAddress -as [string]))
  Write-Host (" Type : {0}" -f $dg.GroupType)
}
catch {
  Write-Error "❌ Failed to create distribution group: $($_.Exception.Message)"
  Write-Host "`nDiagnostics:" -ForegroundColor Yellow
  Write-Host (" Cmdlet New-DistributionGroup present?  {0}" -f [bool](Get-Command New-DistributionGroup -ErrorAction SilentlyContinue))
  Write-Host (" Cmdlet Get-EXORecipient present?      {0}" -f [bool](Get-Command Get-EXORecipient -ErrorAction SilentlyContinue))
  Write-Host (" Cmdlet Get-Recipient present?         {0}" -f [bool](Get-Command Get-Recipient -ErrorAction SilentlyContinue))
  Write-Host " PSModulePath:"; $env:PSModulePath -split [IO.Path]::PathSeparator | ForEach-Object { "  - $_" }
}
finally {
  try { Disconnect-ExchangeOnline -Confirm:$false } catch {}
}
