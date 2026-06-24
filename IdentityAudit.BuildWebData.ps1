<#
.SYNOPSIS
Builds normalized JSON for the React identity audit dashboard.

.DESCRIPTION
Reads the latest IdentityAudit.Graph_V10 output folder, or a specified OutputFolder, and writes IdentityAudit-AppData.json.
The React app reads this JSON from web-dashboard/public/data/IdentityAudit-AppData.json.
#>

[CmdletBinding()]
param(
    [string] ${OutputRoot} = '.\IdentityAudit-Evidence',
    [string] ${OutputFolder},
    [string] ${WebDashboardPath} = '.\web-dashboard',
    [switch] ${AlsoWriteToRunFolder}
)

$ErrorActionPreference = 'Stop'

function Write-Stage([string] ${Message}) { Write-Host "[IdentityAudit][WebData] ${Message}" -ForegroundColor Cyan }
function Test-Value($Value) { return -not [string]::IsNullOrWhiteSpace([string] ${Value}) }
function Import-CsvSafe([string] ${Path}) { if (-not (Test-Path ${Path})) { return @() }; return @(Import-Csv ${Path} | Where-Object { $null -ne $_ }) }
function Read-JsonSafe([string] ${Path}) { if (-not (Test-Path ${Path})) { return $null }; ${raw}=Get-Content ${Path} -Raw; if (-not (Test-Value ${raw})) { return $null }; return ${raw} | ConvertFrom-Json }
function To-Int($Value) { try { return [int] ${Value} } catch { return 0 } }
function Set-MapValue($Map, $Key, $Value) { if (Test-Value ${Key}) { ${Map}[[string] ${Key}] = ${Value} } }
function Get-MapValue($Map, $Key) { if (-not (Test-Value ${Key})) { return $null }; if (${Map}.ContainsKey([string] ${Key})) { return ${Map}[[string] ${Key}] }; return $null }

if (-not (Test-Value ${OutputFolder})) {
    ${latest} = Get-ChildItem -Path ${OutputRoot} -Directory -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not ${latest}) { throw "No run folders found under ${OutputRoot}" }
    ${OutputFolder} = ${latest}.FullName
}
if (-not (Test-Path ${OutputFolder})) { throw "Output folder not found: ${OutputFolder}" }

Write-Stage "Reading ${OutputFolder}"

${users} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Users.csv')
${groups} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Groups.csv')
${members} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembers.csv')
${owners} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupOwners.csv')
${deptMatrix} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-DepartmentGroupMatrix.csv')
${exceptions} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Exceptions.csv')
${riskScores} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-RiskScores.csv')
${nodes} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Nodes.csv')
${edges} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Edges.csv')
${paths} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-PrivilegedPaths.csv')
${cycles} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-CircularNesting.csv')
${nestingStats} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-NestingStats.csv')
${groupUserDetail} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupUserDetail.csv')
${userGroupAssociations} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-UserGroupAssociations.csv')
${groupMembershipSummary} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembershipSummary.csv')
${ownerGroupAssociations} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-OwnerGroupAssociations.csv')
${manifestPath} = Join-Path ${OutputFolder} 'IdentityAudit-Manifest.md'
${manifest} = if (Test-Path ${manifestPath}) { Get-Content ${manifestPath} -Raw } else { '' }

if (@(${groupUserDetail}).Count -eq 0 -and @(${members}).Count -gt 0) {
    Write-Stage 'Detail CSVs missing; generating lightweight in-memory group/user detail rows'
    ${riskMap} = @{}
    foreach (${risk} in ${riskScores}) { Set-MapValue ${riskMap} ${risk}.GroupId ${risk} }
    ${groupMap} = @{}
    foreach (${group} in ${groups}) { Set-MapValue ${groupMap} ${group}.GroupId ${group} }
    ${ownerByGroup} = @{}
    foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) { Set-MapValue ${ownerByGroup} ${ownerGroup}.Name @(${ownerGroup}.Group) }
    ${groupUserDetail} = foreach (${member} in ${members}) {
        ${group}=Get-MapValue ${groupMap} ${member}.GroupId
        ${risk}=Get-MapValue ${riskMap} ${member}.GroupId
        ${ownersForGroup}=@(Get-MapValue ${ownerByGroup} ${member}.GroupId)
        [pscustomobject]@{
            GroupName=${member}.GroupName; GroupCategory=${member}.GroupCategory; RiskScore=$(if(${risk}){${risk}.RiskScore}else{''}); RiskDrivers=$(if(${risk}){${risk}.RiskDrivers}else{''}); Owners=(@(${ownersForGroup})|ForEach-Object{if($_.OwnerUPN){$_.OwnerUPN}else{$_.OwnerDisplayName}}|Select-Object -Unique) -join '; '; MemberDisplayName=${member}.MemberDisplayName; MemberUPN=${member}.MemberUPN; MemberType=${member}.MemberType; MemberDepartment=${member}.MemberDepartment; MemberJobTitle=${member}.MemberJobTitle; MemberCompanyName=${member}.MemberCompanyName; MemberAccountEnabled=${member}.MemberAccountEnabled; MembershipMode=${member}.MembershipMode; GroupId=${member}.GroupId; MemberId=${member}.MemberId
        }
    }
}

${groupCountsByCategory} = @(${groups} | Group-Object GroupCategory | ForEach-Object { [pscustomobject]@{ category=$_.Name; count=$_.Count } })
${memberCountsByType} = @(${members} | Group-Object MemberType | ForEach-Object { [pscustomobject]@{ type=$_.Name; count=$_.Count } })
${riskBuckets} = @(
    [pscustomobject]@{ bucket='0'; count=@(${riskScores} | Where-Object { (To-Int $_.RiskScore) -eq 0 }).Count },
    [pscustomobject]@{ bucket='1-24'; count=@(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 1 -and (To-Int $_.RiskScore) -le 24 }).Count },
    [pscustomobject]@{ bucket='25-49'; count=@(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 25 -and (To-Int $_.RiskScore) -le 49 }).Count },
    [pscustomobject]@{ bucket='50-74'; count=@(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 50 -and (To-Int $_.RiskScore) -le 74 }).Count },
    [pscustomobject]@{ bucket='75+'; count=@(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 75 }).Count }
)

${summary} = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('o')
    sourceFolder = ${OutputFolder}
    userCount = @(${users}).Count
    groupCount = @(${groups}).Count
    membershipRowCount = @(${members}).Count
    ownerRowCount = @(${owners}).Count
    nodeCount = @(${nodes}).Count
    edgeCount = @(${edges}).Count
    privilegedPathCount = @(${paths}).Count
    cycleCount = @(${cycles}).Count
    ownerlessGroupCount = @(${groups} | Where-Object { (To-Int $_.OwnerCount) -eq 0 }).Count
    roleAssignableGroupCount = @(${groups} | Where-Object { $_.IsAssignableToRole -eq 'True' -or $_.IsAssignableToRole -eq $true }).Count
    highRiskGroupCount = @(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 50 }).Count
    groupCountsByCategory = ${groupCountsByCategory}
    memberCountsByType = ${memberCountsByType}
    riskBuckets = ${riskBuckets}
}

${payload} = [pscustomobject]@{
    schemaVersion = '1.0'
    summary = ${summary}
    manifest = ${manifest}
    users = ${users}
    groups = ${groups}
    members = ${members}
    owners = ${owners}
    departmentMatrix = ${deptMatrix}
    exceptions = ${exceptions}
    riskScores = ${riskScores}
    nodes = ${nodes}
    edges = ${edges}
    paths = ${paths}
    cycles = ${cycles}
    nestingStats = ${nestingStats}
    groupUserDetail = ${groupUserDetail}
    userGroupAssociations = ${userGroupAssociations}
    groupMembershipSummary = ${groupMembershipSummary}
    ownerGroupAssociations = ${ownerGroupAssociations}
}

${dataDir} = Join-Path ${WebDashboardPath} 'public\data'
New-Item -ItemType Directory -Path ${dataDir} -Force | Out-Null
${webDataPath} = Join-Path ${dataDir} 'IdentityAudit-AppData.json'
${payload} | ConvertTo-Json -Depth 10 | Out-File ${webDataPath} -Encoding utf8 -Force
Write-Stage "Wrote ${webDataPath}"

if (${AlsoWriteToRunFolder}) {
    ${runDataPath} = Join-Path ${OutputFolder} 'IdentityAudit-AppData.json'
    ${payload} | ConvertTo-Json -Depth 10 | Out-File ${runDataPath} -Encoding utf8 -Force
    Write-Stage "Wrote ${runDataPath}"
}

Write-Host "Web data ready: ${webDataPath}"
