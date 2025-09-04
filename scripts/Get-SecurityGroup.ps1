<#
.SYNOPSIS
  Get information about an Entra ID security group.

.USAGE
  # By name (positional parameter 0)
  .\Get-SecurityGroup.ps1 "RS-Oslo-Crew"

  # By Id
  .\Get-SecurityGroup.ps1 -GroupId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"

.REQUIREMENTS
  Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Position = 0, ParameterSetName = 'ByName', Mandatory = $true)]
    [string] $GroupName,

    [Parameter(ParameterSetName = 'ById', Mandatory = $true)]
    [string] $GroupId
)

# Connect (needs Directory.Read.All or Group.Read.All)
Connect-MgGraph -Scopes "Directory.Read.All","Group.Read.All" | Out-Null

switch ($PSCmdlet.ParameterSetName) {
    'ByName' {
        $escaped = $GroupName.Replace("'","''")
        $groups = @( Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All )
        if ($groups.Count -eq 0) {
            Write-Error "❌ No security group found with name '${GroupName}'"
            exit 1
        }
        if ($groups.Count -gt 1) {
            Write-Warning "⚠️ Multiple groups named '${GroupName}' found:"
            $groups | Select-Object Id, DisplayName, MailNickname, Description, MailEnabled, SecurityEnabled
            exit 0
        }
        $group = $groups[0]
    }
    'ById' {
        try {
            $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        } catch {
            Write-Error "❌ No security group found with Id '${GroupId}'"
            exit 1
        }
    }
}

Write-Host "✅ Found group:"
$group | Select-Object Id, DisplayName, MailNickname, Description, MailEnabled, SecurityEnabled, Visibility
