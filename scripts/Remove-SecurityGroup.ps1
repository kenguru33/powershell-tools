<#
.SYNOPSIS
  Delete an Entra ID security group (only if empty).

.USAGE
  .\Remove-SecurityGroup.ps1 "My-Security-Group"
  .\Remove-SecurityGroup.ps1 -GroupId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName='ByName')]
    [string] $GroupName,

    [Parameter(Mandatory = $true, ParameterSetName='ById')]
    [string] $GroupId,

    [switch] $Force # if set, will delete even if group has members
)

# Connect (needs Group.ReadWrite.All + Directory.Read.All)
Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.Read.All" | Out-Null

# Resolve group
switch ($PSCmdlet.ParameterSetName) {
    'ByName' {
        $escaped = $GroupName.Replace("'","''")
        $g = @( Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All )
        if ($g.Count -eq 0) { Write-Error "❌ Group not found: ${GroupName}"; exit 1 }
        if ($g.Count -gt 1) { Write-Error "❌ Multiple groups named '${GroupName}' found. Use -GroupId."; exit 1 }
        $group = $g[0]
    }
    'ById' {
        try { $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop }
        catch { Write-Error "❌ Group not found by Id: ${GroupId}"; exit 1 }
    }
}

if (-not $group -or -not $group.Id) {
    Write-Error "❌ Resolved group has no valid Id."
    exit 1
}

Write-Host "➡️ Found group: $($group.DisplayName) (Id: $($group.Id))"

# Check members
$members = @()
try {
    $members = @( Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop )
} catch {
    Write-Warning "⚠️ Could not list members for group ${GroupName}: $($_.Exception.Message)"
}

if ($members.Count -gt 0 -and -not $Force) {
    Write-Warning "⚠️ Group '${GroupName}' has $($members.Count) members. Use -Force to delete anyway."
    exit 2
}

# Delete
try {
    Remove-MgGroup -GroupId $group.Id -ErrorAction Stop
    Write-Host "✅ Deleted group: $($group.DisplayName) (Id: $($group.Id))"
} catch {
    Write-Error "❌ Failed to delete group: $($_.Exception.Message)"
    exit 3
}
