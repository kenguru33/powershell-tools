<#
.SYNOPSIS
  Create (or reuse) an Entra ID security group.

.USAGE
  # Minimum (just name)
  .\Create-SecurityGroup.ps1 "RS-Oslo-Crew"

  # With description
  .\Create-SecurityGroup.ps1 "RS-Oslo-Crew" "Rescue team in Oslo"

  # With description and explicit mailNickname
  .\Create-SecurityGroup.ps1 "RS-Oslo-Crew" "Rescue team in Oslo" "rsoslocrew"

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $GroupName,

    [Parameter(Position = 1)]
    [string] $Description = "Security group created via PowerShell",

    [Parameter(Position = 2)]
    [string] $MailNickname
)

# Connect (needs Group.ReadWrite.All)
Connect-MgGraph -Scopes "Group.ReadWrite.All" | Out-Null

# Try to find existing group
$escaped = $GroupName.Replace("'","''")
$existing = Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All

if ($existing) {
    $group = $existing[0]
    Write-Host "ℹ️ Using existing security group: '$($group.DisplayName)'  Id: $($group.Id)"
    return
}

# Ensure a unique MailNickname if not provided
if (-not $MailNickname -or [string]::IsNullOrWhiteSpace($MailNickname)) {
    $base = ($GroupName -replace '[^a-zA-Z0-9]','').ToLower()
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'sg' }
    $MailNickname = "$base$([System.Guid]::NewGuid().ToString('N').Substring(0,6))"
}

$group = New-MgGroup `
    -DisplayName $GroupName `
    -MailEnabled:$false `
    -MailNickname $MailNickname `
    -SecurityEnabled:$true `
    -Description $Description

Write-Host "✅ Created security group '${GroupName}'"
Write-Host "   Id: $($group.Id)"
