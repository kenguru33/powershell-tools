<# 
.SYNOPSIS
  Create (or reuse) an Entra ID security group.

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
#>

param(
    [Parameter(Mandatory = $true)] [string] $GroupName,
    [string] $Description = "Security group created via PowerShell",
    [string] $MailNickname # optional; will be generated if omitted
)

# Connect (needs Group.ReadWrite.All)
Connect-MgGraph -Scopes "Group.ReadWrite.All" | Out-Null

# Try find existing group (security-enabled)
$escaped = $GroupName.Replace("'","''")
$existing = Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All

if ($existing) {
    $group = $existing[0]
    Write-Host "ℹ️ Using existing security group: '$($group.DisplayName)'  Id: $($group.Id)"
    return
}

# Ensure a unique mailNickname (required even if MailEnabled:$false)
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
