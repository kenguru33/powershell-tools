[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
    [string] $GroupName,

    [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
    [string] $GroupId,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateScript({ Test-Path $_ })]
    [string] $CsvPath
)

# Connect (needs Group.ReadWrite.All + User.Read.All)
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All","Directory.Read.All" | Out-Null

# Resolve group
switch ($PSCmdlet.ParameterSetName) {
    'ByName' {
        $escaped = $GroupName.Replace("'","''")
        $g = @( Get-MgGroup -Filter "displayName eq '${escaped}' and securityEnabled eq true" -ConsistencyLevel eventual -All )
        if ($g.Count -eq 0) { throw "Group not found by name: ${GroupName}" }
        if ($g.Count -gt 1) { throw "Multiple groups named '${GroupName}' found. Use -GroupId." }
        $group = $g[0]
    }
    'ById' {
        try { $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop }
        catch { throw "Group not found by Id: ${GroupId}" }
    }
}

if (-not $group -or -not $group.Id -or [string]::IsNullOrWhiteSpace($group.Id)) {
    throw "Resolved group object has no valid Id. Aborting."
}

Write-Host "Target group: $($group.DisplayName) (Id: $($group.Id))"

# Load CSV (must have header 'Email')
$rows = Import-Csv -Path $CsvPath
if (-not $rows -or -not ($rows | Get-Member -Name Email -MemberType NoteProperty)) {
    throw "CSV must have a header 'Email' and at least one row."
}

# Cache current member Ids to skip duplicates
$current = @{}
try {
    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
    foreach ($m in $members) {
        if ($m -and $m.Id) { $current[$m.Id] = $true }
    }
} catch { }

$added = 0; $skipped = 0; $notFound = @(); $invalid = @()

foreach ($row in $rows) {
    $email = ($row.Email).Trim()
    if ([string]::IsNullOrWhiteSpace($email)) { continue }

    # Look up user by mail or UPN; force array typing and guard count
    $emailEsc = $email.Replace("'","''")
    $users = @(
        Get-MgUser -Filter "mail eq '${emailEsc}' or userPrincipalName eq '${emailEsc}'" `
                   -ConsistencyLevel eventual -All -ErrorAction SilentlyContinue
    )

    if ($users.Count -eq 0) {
        Write-Warning "No user found for '${email}'"
        $notFound += $email
        continue
    }

    $u = $users | Select-Object -First 1
    if (-not $u -or -not $u.Id -or [string]::IsNullOrWhiteSpace($u.Id)) {
        Write-Warning "User object for '${email}' has no valid Id; skipping."
        $invalid += $email
        continue
    }

    if ($current.ContainsKey($u.Id)) {
        $skipped++
        Write-Host "Already a member: ${email}"
        continue
    }

    try {
        if (Get-Command New-MgGroupMember -ErrorAction SilentlyContinue) {
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $u.Id -ErrorAction Stop | Out-Null
        } else {
            Add-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($u.Id)" } -ErrorAction Stop | Out-Null
        }
        $current[$u.Id] = $true
        $added++
        Write-Host "✅ Added: ${email}"
    } catch {
        Write-Warning "Failed to add ${email}: $($_.Exception.Message)"
    }
}

# Summary
Write-Host "`n=== Summary: $($group.DisplayName) ==="
Write-Host "  Added:   ${added}"
Write-Host "  Skipped: ${skipped} (already members)"
if ($notFound.Count) {
    Write-Host "  Not found ($($notFound.Count)):"; $notFound | ForEach-Object { Write-Host "    ${_}" }
}
if ($invalid.Count) {
    Write-Host "  Invalid user objects ($($invalid.Count)):"; $invalid | ForEach-Object { Write-Host "    ${_}" }
}
