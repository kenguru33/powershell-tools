<#
.SYNOPSIS
  Get information about an Entra ID user by UPN.

.USAGE
  .\Get-UserByUpn.ps1 user@domain.com

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $UserPrincipalName
)

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All" | Out-Null

try {
    $user = Get-MgUser -UserId $UserPrincipalName -Property id,displayName,userPrincipalName,mail,accountEnabled,givenName,surname,jobTitle,department,createdDateTime -ErrorAction Stop
} catch {
    Write-Error "❌ User not found for UPN: $UserPrincipalName"
    exit 1
}

Write-Host "✅ Found user:`n"
$user | Select-Object Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, GivenName, Surname, JobTitle, Department, CreatedDateTime
