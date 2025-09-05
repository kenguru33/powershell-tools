<#
.SYNOPSIS
  Hide (or unhide) a recipient from the Global Address List (GAL).

.DESCRIPTION
  Works for:
    - User mailboxes (including Shared)         -> Set-Mailbox
    - Distribution lists (DL)                    -> Set-DistributionGroup
    - Mail-enabled security groups               -> Set-DistributionGroup
    - Microsoft 365 groups (GroupMailbox)        -> Set-UnifiedGroup
    - Dynamic distribution groups                -> Set-DynamicDistributionGroup
    - Mail users (mail-enabled users)            -> Set-MailUser
    - Mail contacts                              -> Set-MailContact

.USAGE
  # Hide a DL or user (default action)
  .\Hide-FromAddressBook.ps1 "RS-Oslo-Crew"

  # Unhide
  .\Hide-FromAddressBook.ps1 "RS-Oslo-Crew" -Unhide

  # Use SMTP address, alias or UPN
  .\Hide-FromAddressBook.ps1 "svein.moe@rs.no"

.REQUIREMENTS
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  Connect-ExchangeOnline with an account that can manage recipients
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string] $Identity,

  [switch] $Unhide,

  # Pass -WhatIf to see what would happen without changing anything
  [switch] $WhatIf
)

# Ensure module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
  Write-Error "ExchangeOnlineManagement module not found. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
  exit 1
}

# Connect if not already
try {
  if (-not (Get-ConnectionInformation)) {
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
  }
} catch {
  # Older EXO module versions might throw if not connected; just try to connect
  try { Connect-ExchangeOnline -ShowBanner:$false | Out-Null } catch { throw }
}

# Resolve the recipient
$recipient = $null
try {
  # Get-Recipient resolves most identifiers (DisplayName/UPN/Alias/SMTP)
  $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
} catch {
  Write-Error "❌ Could not resolve recipient for '$Identity'."
  exit 2
}

$rtype = $recipient.RecipientType
$rdetail = $recipient.RecipientTypeDetails
$targetValue = if ($Unhide) { $false } else { $true }
$actionVerb = if ($Unhide) { "Unhide" } else { "Hide" }

Write-Host ("Target: {0}  <{1}>  Type: {2} / {3}" -f $recipient.DisplayName, $recipient.PrimarySmtpAddress, $rtype, $rdetail)

# Dispatch by recipient type
switch -Regex ($rdetail) {
  '^(UserMailbox|SharedMailbox)$' {
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-Mailbox @params
    break
  }

  '^(MailUniversalSecurityGroup|MailUniversalDistributionGroup)$' {
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-DistributionGroup @params
    break
  }

  '^GroupMailbox$' {  # Microsoft 365 group
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-UnifiedGroup @params
    break
  }

  '^DynamicDistributionGroup$' {
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-DynamicDistributionGroup @params
    break
  }

  '^MailUser$' {
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-MailUser @params
    break
  }

  '^MailContact$' {
    $params = @{ Identity = $recipient.Identity; HiddenFromAddressListsEnabled = $targetValue }
    if ($WhatIf) { $params.WhatIf = $true }
    Set-MailContact @params
    break
  }

  default {
    Write-Error "❌ Recipient type '$rdetail' is not supported by this script."
    exit 3
  }
}

Write-Host ("✅ {0} '{1}' from GAL." -f $actionVerb, $recipient.DisplayName)
