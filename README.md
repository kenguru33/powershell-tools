# powershell-tools

## Installation

Enter powershell:
```
pwsh
```

One-time for current session
```
$env:PATH = "/Users/bernt/Projects/PowerShell-Tools/scripts" + [IO.Path]::PathSeparator + $env:PATH
```
Create Profile or use existing:
```
New-Item -ItemType File -Path $PROFILE -Force
```

Add the path to your profile (Use nvim or whatever editor you prefer):
```
nvim $PROFILE
```
Add this to your profile:
```
$scriptDir = "/Users/bernt/Projects/PowerShell-Tools/scripts"
if (-not ($env:PATH -split [IO.Path]::PathSeparator | Where-Object { $_ -eq $scriptDir })) {
    $env:PATH = $scriptDir + [IO.Path]::PathSeparator + $env:PATH
}

function Get-SecurityGroup { /Users/bernt/Projects/PowerShell-Tools/scripts/Get-SecurityGroup.ps1 @args }
function Create-SecurityGroup { /Users/bernt/Projects/PowerShell-Tools/scripts/Create-SecurityGroup.ps1 @args }
function Add-UsersFromCsv-ToGroup { /Users/bernt/Projects/PowerShell-Tools/scripts/Add-UsersFromCsv-ToGroup.ps1 @args }
function Add-UpnToGroup { /Users/bernt/Projects/PowerShell-Tools/scripts/Add-UpnToGroup.ps1 @args }
function Get-UpnByAlias { /Users/bernt/Projects/PowerShell-Tools/scripts/Get-UpnByAlias.ps1 @args }
function Remove-SecurityGroup { /Users/bernt/Projects/PowerShell-Tools/scripts/Remove-SecurityGroup.ps1 @args }
```

**Create Security Group** 
```
.\Create-SecurityGroup.ps1 -GroupName "RSFA-Security-Editors"
```

With custom nickname/description:
```
.\Create-SecurityGroup.ps1 -GroupName "RSFA-Security-Editors" -MailNickname "rsfaeditors" -Description "Editors group"
```
---
**Add users from a csv file**
Add to group by name
```
.\Add-UsersFromCsv-ToGroup.ps1 -GroupName "RSFA-Security-Editors" -CsvPath .\members.csv
```
By group Id
```
.\Add-UsersFromCsv-ToGroup.ps1 -GroupId "a1b2c3d4-e5f6-7890-abcd-1234567890ef" -CsvPath .\members.csv
```

*CSV format*
```
Email
bernt@domain.no
kenneth@domain.no
matheo@domain.no
```

**Add UPN to Group**
```
.\Add-UpnToGroup.ps1 svein.moe@rs.no "RSFA-Security-Editors"
```

**Get UPN By Alias**
```
.\Get-UpnByAlias.ps1 sveintm@rs.no
```
