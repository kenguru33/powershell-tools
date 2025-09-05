<#
.SYNOPSIS
  For each input Id/GUID/DN, print that recipient’s details,
  then immediately print ONLY that recipient’s SMTP aliases.
  If no aliases exist, print "Aliases for object above: <none>".

.USAGE
  .\Get-RecipientsById.ps1 "id-or-dn-1","id-or-dn-2"
  "id1","id2" | .\Get-RecipientsById.ps1
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true, Position=0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
  [string[]] $Id
)

# Ensure Exchange connection
if (-not (Get-Module ExchangeOnlineManagement)) {
  throw "ExchangeOnlineManagement module is required. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
try {
  if (-not (Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false | Out-Null }
} catch {
  Connect-ExchangeOnline -ShowBanner:$false | Out-Null
}

foreach ($identity in $Id) {
  try {
    # IMPORTANT: Get ALL matches for this identity (array), not just one
    $recips = @( Get-Recipient -Identity $identity -ResultSize Unlimited -ErrorAction Stop )
  } catch {
    Write-Warning ("Failed to find recipient(s) for Id: {0} -- {1}" -f $identity, $_.Exception.Message)
    continue
  }

  if (-not $recips -or $recips.Count -eq 0) {
    Write-Warning ("No recipient found for Id: {0}" -f $identity)
    continue
  }

  foreach ($recip in $recips) {
    # --- Print THIS object's details ---
    ($recip | Select-Object `
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientType,
        RecipientTypeDetails,
        Id,
        DistinguishedName |
      Format-List | Out-String -Width 4096) | Write-Host

    # --- Immediately print THIS object's aliases (only these) ---
    $smtpFound = $false
    if ($recip.EmailAddresses) {
      foreach ($pa in $recip.EmailAddresses) {
        $s = $pa.ToString()
        if ($s.StartsWith('SMTP:') -or $s.StartsWith('smtp:')) {
          if (-not $smtpFound) {
            Write-Host "Aliases for the object above:"
            $smtpFound = $true
          }
          $isPrimary = $s.StartsWith('SMTP:')
          $addr      = $s.Substring(5)
          $tag       = if ($isPrimary) { " [PRIMARY]" } else { "" }
          Write-Host ("  - {0}{1}" -f $addr, $tag)
        }
      }
    }

    if (-not $smtpFound) {
      Write-Host "Aliases for the object above: <none>"
    }

    # blank line between objects
    Write-Host ""
  }
}
