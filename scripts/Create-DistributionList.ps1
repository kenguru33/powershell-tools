<#
.SYNOPSIS
  Create a Distribution List (mail-enabled security group) in Entra ID.

.USAGE
  # Minimal (positional 0 = display name, positional 1 = mailNickname)
  .\Create-DistributionList.ps1 "RSFA-DL-Crew" "rsfacrew"

  # With description
  .\Create-DistributionList.ps1 "RSFA-DL-Crew" "rsfacrew" "Crew Distribution List"
  
.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
  Permissions: Group.ReadWrite.All
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $DisplayName,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $MailNickname,

    [Parameter(Position = 2)]
    [string] $Description = "Distribution list created via PowerShell"
)

# Connect
Connect-MgGraph -Scopes "Group.ReadWrite.All" | Out-Null

# Check if already exists
$escaped = $DisplayName.Replace("'","''")
$existing = @( Get-MgGroup -Filter "displayName eq '${escaped}' and mailEnabled eq true" -All )
if ($existing.Count -gt 0) {
    Write-Host "ℹ️ Distribution list already exists: $($existing[0].DisplayName) (Id: $($existing[0].Id))"
    exit 0
}

# Create Distribution List (mail-enabled security group)
$group = New-MgGroup `
    -DisplayName $DisplayName `
    -MailEnabled:$true `
    -MailNickname $MailNickname `
    -SecurityEnabled:$false `
    -GroupTypes @() `
    -Description $Description

Write-Host "✅ Created distribution list '$DisplayName'"
Write-Host "   Id: $($group.Id)"
Write-Host "   Email: $($group.Mail)"
