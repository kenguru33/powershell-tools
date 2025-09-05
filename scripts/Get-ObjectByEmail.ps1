<#
.SYNOPSIS
  Get any Entra ID object (user, group, contact, guest, etc.) by email address.

.USAGE
  # Single email
  .\Get-ObjectByEmail.ps1 "svein.hansen@rs.no"

  # Multiple emails
  .\Get-ObjectByEmail.ps1 "svein.hansen@rs.no","oslo-crew@rs.no"

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  Connect-MgGraph -Scopes "Directory.Read.All"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]] $Email
)

# Ensure connection
Connect-MgGraph -Scopes "Directory.Read.All" | Out-Null

foreach ($addr in $Email) {
    $search = $addr.Trim()
    if ([string]::IsNullOrWhiteSpace($search)) { continue }

    Write-Host "`n🔎 Searching for: $search"

    # Try user first (UPN or mail)
    $user = Get-MgUser -Filter "userPrincipalName eq '$search' or mail eq '$search'" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
    if ($user) {
        Write-Host "✅ Found User:"
        $user | Select-Object Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType
        continue
    }

    # Try groups
    $group = Get-MgGroup -Filter "mail eq '$search' or proxyAddresses/any(p:p eq 'SMTP:$search')" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
    if ($group) {
        Write-Host "✅ Found Group(s):"
        $group | Select-Object Id, DisplayName, Mail, MailEnabled, SecurityEnabled, GroupTypes
        continue
    }

    # Try contacts
    $contact = Get-MgContact -Filter "mail eq '$search'" -All -ErrorAction SilentlyContinue
    if ($contact) {
        Write-Host "✅ Found Contact:"
        $contact | Select-Object Id, DisplayName, Mail
        continue
    }

    # Try mail users (guests, external)
    $mailUser = Get-MgUser -Filter "otherMails/any(c:c eq '$search')" -All -ErrorAction SilentlyContinue
    if ($mailUser) {
        Write-Host "✅ Found Mail User (Guest/External):"
        $mailUser | Select-Object Id, DisplayName, UserPrincipalName, Mail, UserType
        continue
    }

    Write-Warning "❌ No object found for $search"
}
