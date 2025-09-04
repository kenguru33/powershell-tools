<# 
.SYNOPSIS
  Look up the UserPrincipalName (UPN) for an email alias in Entra ID (Azure AD).

.USAGE
  # Positional (parameter 0):
  .\Get-UpnByAlias.ps1 sveintm@rs.no

  # Or named:
  .\Get-UpnByAlias.ps1 -Alias sveintm@rs.no

.NOTES
  Requires Microsoft Graph PowerShell SDK:
    Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding()]
param(
  [string] $Alias,
  [switch] $Raw  # if set, only output the UPN (clean for scripting)
)

# Accept positional parameter 0 if -Alias was not provided
if (-not $Alias -and $args.Count -gt 0) {
  $Alias = $args[0]
}

if ([string]::IsNullOrWhiteSpace($Alias)) {
  Write-Error "Provide an alias (e.g. '.\Get-UpnByAlias.ps1 someone@contoso.com' or -Alias)."
  exit 1
}

# Connect (needs read scope)
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All" | Out-Null

# Build a single OData filter that covers:
# - mail
# - userPrincipalName
# - otherMails (aliases)
# - proxyAddresses (SMTP: / smtp:)
$aliasEsc = $Alias.Replace("'","''")
$filter = @"
mail eq '${aliasEsc}' or userPrincipalName eq '${aliasEsc}' or
otherMails/any(x:x eq '${aliasEsc}') or
proxyAddresses/any(p:p eq 'SMTP:${aliasEsc}') or
proxyAddresses/any(p:p eq 'smtp:${aliasEsc}')
"@ -replace "`r?`n"," "  # compress to one line

# Execute and coerce to array
$users = @(
  Get-MgUser -Filter $filter -Property userPrincipalName,mail,otherMails,proxyAddresses,displayName `
             -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
)

if ($users.Count -eq 0) {
  Write-Error "No user found for alias '${Alias}'."
  exit 2
}

# Prefer first match (you can widen if needed)
$u = $users | Select-Object -First 1

if ($Raw) {
  # Output only the UPN for scripting
  $u.UserPrincipalName
  exit 0
}

# Pretty output
[pscustomobject]@{
  DisplayName       = $u.DisplayName
  UserPrincipalName = $u.UserPrincipalName
  Mail              = $u.Mail
  OtherMails        = ($u.OtherMails -join "; ")
  ProxyAddresses    = ($u.ProxyAddresses -join "; ")
}
