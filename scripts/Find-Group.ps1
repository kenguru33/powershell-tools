<#
.SYNOPSIS
  Get information about an Entra ID group (any type) with fuzzy name search.

.USAGE
  # Fuzzy by name (positional parameter 0)
  .\Get-SecurityGroup.ps1 "RS-Oslo"

  # Exact by Id
  .\Get-SecurityGroup.ps1 -GroupId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"

.NOTES
  - Fuzzy search order: exact -> startswith -> full-text ($search on displayName)
  - Works for any group type (Security, Mail-enabled, M365 Group)
  - Requires Microsoft Graph PowerShell SDK
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Position = 0, ParameterSetName = 'ByName', Mandatory = $true)]
    [string] $GroupName,

    [Parameter(ParameterSetName = 'ById', Mandatory = $true)]
    [string] $GroupId
)

# Connect (read scopes)
Connect-MgGraph -Scopes "Directory.Read.All","Group.Read.All" | Out-Null

function Select-GroupOutput {
    param([Parameter(Mandatory=$true)] $Group)
    $Group | Select-Object `
        Id,
        DisplayName,
        Mail,
        MailNickname,
        SecurityEnabled,
        MailEnabled,
        Visibility,
        GroupTypes,
        Description,
        CreatedDateTime
}

switch ($PSCmdlet.ParameterSetName) {
    'ById' {
        try {
            $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
            Write-Host "✅ Found group by Id:`n"
            Select-GroupOutput -Group $group
            break
        } catch {
            Write-Error "❌ No group found with Id '${GroupId}'"
            exit 1
        }
    }

    'ByName' {
        $escaped = $GroupName.Replace("'","''")

        # 1) Exact displayName match (any type)
        $exact = @(
            Get-MgGroup -Filter "displayName eq '${escaped}'" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
        )

        if ($exact.Count -eq 1) {
            Write-Host "✅ Exact match:`n"
            Select-GroupOutput -Group $exact[0]
            break
        }

        # 2) StartsWith on displayName (narrow but fuzzy)
        $starts = @()
        try {
            $starts = @(
                Get-MgGroup -Filter "startsWith(displayName,'${escaped}')" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
            )
        } catch { }

        # 3) Full-text search on displayName (broad fuzzy)
        $searchRes = @()
        try {
            # $search requires ConsistencyLevel: eventual
            $searchRes = @(
                Get-MgGroup -Search "displayName:${GroupName}" -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
            )
        } catch { }

        # Merge & de-dup by Id (prefer exact > starts > search)
        $all = @()
        $seen = @{}
        foreach ($g in @($exact + $starts + $searchRes)) {
            if ($g -and $g.Id -and -not $seen.ContainsKey($g.Id)) {
                $seen[$g.Id] = $true
                $all += $g
            }
        }

        if ($all.Count -eq 0) {
            Write-Error "❌ No groups found matching '${GroupName}'."
            exit 1
        }

        if ($all.Count -eq 1) {
            Write-Host "✅ Match:`n"
            Select-GroupOutput -Group $all[0]
            break
        }

        Write-Warning "⚠️ Multiple groups matched '${GroupName}'. Refine your query or use -GroupId:"
        $all |
            Sort-Object DisplayName |
            Select-Object Id, DisplayName, Mail, MailNickname, SecurityEnabled, MailEnabled, GroupTypes, Visibility |
            Format-Table -AutoSize

        # Exit without error since we returned useful info
        exit 0
    }
}
