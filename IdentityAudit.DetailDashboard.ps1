<#
.SYNOPSIS
Builds a searchable detail dashboard from IdentityAudit.Graph_V10 output.

.DESCRIPTION
Reads the latest IdentityAudit-Evidence run folder, or a specified OutputFolder, and creates searchable group and user association detail.
This script does not connect to Microsoft Graph. It post-processes existing CSV evidence.
#>

[CmdletBinding()]
param(
    [string] ${OutputRoot} = '.\IdentityAudit-Evidence',
    [string] ${OutputFolder},
    [int] ${HighRiskScoreThreshold} = 50,
    [switch] ${OpenDashboard}
)

$ErrorActionPreference = 'Stop'

function Write-Stage([string] ${Message}) { Write-Host "[IdentityAudit][Detail] ${Message}" -ForegroundColor Cyan }
function Test-Value($Value) { return -not [string]::IsNullOrWhiteSpace([string] ${Value}) }
function HtmlSafe($Value) { if ($null -eq ${Value}) { return '' }; return [System.Net.WebUtility]::HtmlEncode([string] ${Value}) }
function Import-CsvSafe([string] ${Path}) { if (-not (Test-Path ${Path})) { return @() }; return @(Import-Csv ${Path} | Where-Object { $null -ne $_ }) }
function Export-CsvSafe($Rows, [string] ${Path}) { ${items} = @(${Rows} | Where-Object { $null -ne $_ }); if (${items}.Count -eq 0) { New-Item -ItemType File -Path ${Path} -Force | Out-Null } else { ${items} | Export-Csv ${Path} -NoTypeInformation } }
function Set-MapValue($Map, $Key, $Value) { if (Test-Value ${Key}) { ${Map}[[string] ${Key}] = ${Value} } }
function Get-MapValue($Map, $Key) { if (-not (Test-Value ${Key})) { return $null }; if (${Map}.ContainsKey([string] ${Key})) { return ${Map}[[string] ${Key}] }; return $null }
function Join-Unique($Values) { return (@(${Values} | Where-Object { Test-Value $_ } | Select-Object -Unique) -join '; ') }
function To-Int($Value) { try { return [int]${Value} } catch { return 0 } }

function New-SearchableTable([string] ${Title}, [string] ${Description}, $Rows, [string[]] ${Columns}, [string] ${TableId}) {
    ${items} = @(${Rows} | Where-Object { $null -ne $_ })
    ${html} = "<section class='panel'><div class='section-head'><div><h2>$(HtmlSafe ${Title})</h2><p class='muted'>$(HtmlSafe ${Description})</p></div><div class='count'>$(${items}.Count) rows</div></div>"
    if (${items}.Count -eq 0) { return ${html} + '<p class="muted">No records found.</p></section>' }
    ${html} += "<input class='search' placeholder='Search this table...' onkeyup='filterTable(this, `"${TableId}`")' />"
    ${html} += "<div class='table-wrap'><table id='${TableId}'><thead><tr>"
    foreach (${column} in ${Columns}) { ${html} += "<th>$(HtmlSafe ${column})</th>" }
    ${html} += '</tr></thead><tbody>'
    foreach (${row} in ${items}) {
        ${html} += '<tr>'
        foreach (${column} in ${Columns}) {
            ${value} = ''
            if (${row}.PSObject.Properties.Name -contains ${column}) { ${value} = ${row}.${column} }
            ${html} += "<td>$(HtmlSafe ${value})</td>"
        }
        ${html} += '</tr>'
    }
    ${html} += '</tbody></table></div></section>'
    return ${html}
}

if (-not (Test-Value ${OutputFolder})) {
    ${latest} = Get-ChildItem -Path ${OutputRoot} -Directory -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not ${latest}) { throw "No run folders found under ${OutputRoot}" }
    ${OutputFolder} = ${latest}.FullName
}
if (-not (Test-Path ${OutputFolder})) { throw "Output folder not found: ${OutputFolder}" }
Write-Stage "Reading output folder: ${OutputFolder}"

${members} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembers.csv')
${groups} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Groups.csv')
${owners} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupOwners.csv')
${riskScores} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-RiskScores.csv')
${paths} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-PrivilegedPaths.csv')
${cycles} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-CircularNesting.csv')
${nestingStats} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-NestingStats.csv')

${groupMap} = @{}
foreach (${group} in ${groups}) { Set-MapValue ${groupMap} ${group}.GroupId ${group} }
${riskMap} = @{}
foreach (${risk} in ${riskScores}) { Set-MapValue ${riskMap} ${risk}.GroupId ${risk} }
${ownerByGroup} = @{}
foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) { Set-MapValue ${ownerByGroup} ${ownerGroup}.Name @(${ownerGroup}.Group) }

Write-Stage 'Building group-user detail rows'
${groupUserDetail} = @()
foreach (${member} in ${members}) {
    ${group} = Get-MapValue ${groupMap} ${member}.GroupId
    ${risk} = Get-MapValue ${riskMap} ${member}.GroupId
    ${ownersForGroup} = @(Get-MapValue ${ownerByGroup} ${member}.GroupId)
    ${groupUserDetail} += [pscustomobject]@{
        GroupName = ${member}.GroupName
        GroupCategory = ${member}.GroupCategory
        RiskScore = $(if (${risk}) { ${risk}.RiskScore } else { '' })
        RiskDrivers = $(if (${risk}) { ${risk}.RiskDrivers } else { '' })
        IsAssignableToRole = $(if (${group}) { ${group}.IsAssignableToRole } else { ${member}.GroupIsAssignableToRole })
        IsDynamicGroup = $(if (${group}) { ${group}.IsDynamicGroup } else { '' })
        OwnerCount = $(if (${group}) { ${group}.OwnerCount } else { @(${ownersForGroup}).Count })
        Owners = Join-Unique (@(${ownersForGroup}) | ForEach-Object { if ($_.OwnerUPN) { $_.OwnerUPN } else { $_.OwnerDisplayName } })
        MemberDisplayName = ${member}.MemberDisplayName
        MemberUPN = ${member}.MemberUPN
        MemberType = ${member}.MemberType
        MemberDepartment = ${member}.MemberDepartment
        MemberJobTitle = ${member}.MemberJobTitle
        MemberCompanyName = ${member}.MemberCompanyName
        MemberAccountEnabled = ${member}.MemberAccountEnabled
        MembershipMode = ${member}.MembershipMode
        GroupId = ${member}.GroupId
        MemberId = ${member}.MemberId
    }
}
${groupUserDetail} = @(${groupUserDetail} | Sort-Object GroupName, MemberType, MemberDisplayName)

Write-Stage 'Building user-group association rows'
${ownedByPrincipal} = @{}
foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.OwnerId } | Group-Object OwnerId)) { Set-MapValue ${ownedByPrincipal} ${ownerGroup}.Name @(${ownerGroup}.Group) }
${userGroupAssociations} = @()
foreach (${memberGroup} in @(${members} | Where-Object { Test-Value $_.MemberId } | Group-Object MemberId)) {
    ${rows} = @(${memberGroup}.Group)
    if (${rows}.Count -eq 0) { continue }
    ${sample} = ${rows}[0]
    ${ownedGroups} = @(Get-MapValue ${ownedByPrincipal} ${memberGroup}.Name)
    ${riskRows} = @()
    foreach (${row} in ${rows}) {
        ${r} = Get-MapValue ${riskMap} ${row}.GroupId
        if (${r}) { ${riskRows} += ${r} }
    }
    ${highestRisk} = 0
    if (@(${riskRows}).Count -gt 0) { ${highestRisk} = (@(${riskRows} | Sort-Object { To-Int $_.RiskScore } -Descending | Select-Object -First 1).RiskScore) }
    ${userGroupAssociations} += [pscustomobject]@{
        MemberDisplayName = ${sample}.MemberDisplayName
        MemberUPN = ${sample}.MemberUPN
        MemberType = ${sample}.MemberType
        MemberDepartment = ${sample}.MemberDepartment
        MemberJobTitle = ${sample}.MemberJobTitle
        MemberCompanyName = ${sample}.MemberCompanyName
        MemberAccountEnabled = ${sample}.MemberAccountEnabled
        GroupCount = @(${rows}).Count
        Groups = Join-Unique (@(${rows}) | ForEach-Object { $_.GroupName })
        SecurityGroups = Join-Unique (@(${rows} | Where-Object { $_.GroupCategory -eq 'Security' }) | ForEach-Object { $_.GroupName })
        M365Groups = Join-Unique (@(${rows} | Where-Object { $_.GroupCategory -eq 'Microsoft365' }) | ForEach-Object { $_.GroupName })
        RoleAssignableGroupCount = @(${rows} | Where-Object { $_.GroupIsAssignableToRole -eq 'True' -or $_.GroupIsAssignableToRole -eq $true }).Count
        HighRiskGroupCount = @(${riskRows} | Where-Object { (To-Int $_.RiskScore) -ge ${HighRiskScoreThreshold} }).Count
        HighestRiskScore = ${highestRisk}
        OwnedGroupCount = @(${ownedGroups}).Count
        OwnedGroups = Join-Unique (@(${ownedGroups}) | ForEach-Object { $_.GroupName })
        MemberId = ${memberGroup}.Name
    }
}
${userGroupAssociations} = @(${userGroupAssociations} | Sort-Object @{Expression='GroupCount';Descending=$true}, MemberUPN, MemberDisplayName)

Write-Stage 'Building group membership summaries'
${groupMembershipSummary} = @()
foreach (${grouped} in @(${members} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) {
    ${rows} = @(${grouped}.Group)
    if (${rows}.Count -eq 0) { continue }
    ${sample} = ${rows}[0]
    ${group} = Get-MapValue ${groupMap} ${grouped}.Name
    ${risk} = Get-MapValue ${riskMap} ${grouped}.Name
    ${ownersForGroup} = @(Get-MapValue ${ownerByGroup} ${grouped}.Name)
    ${groupMembershipSummary} += [pscustomobject]@{
        GroupName = ${sample}.GroupName
        GroupCategory = ${sample}.GroupCategory
        RiskScore = $(if (${risk}) { ${risk}.RiskScore } else { '' })
        RiskDrivers = $(if (${risk}) { ${risk}.RiskDrivers } else { '' })
        MemberRows = @(${rows}).Count
        UserMembers = @(${rows} | Where-Object { $_.MemberType -eq 'User' }).Count
        GroupMembers = @(${rows} | Where-Object { $_.MemberType -eq 'Group' }).Count
        ServicePrincipalMembers = @(${rows} | Where-Object { $_.MemberType -eq 'ServicePrincipal' }).Count
        DeviceMembers = @(${rows} | Where-Object { $_.MemberType -eq 'Device' }).Count
        Departments = Join-Unique (@(${rows}) | ForEach-Object { $_.MemberDepartment })
        Owners = Join-Unique (@(${ownersForGroup}) | ForEach-Object { if ($_.OwnerUPN) { $_.OwnerUPN } else { $_.OwnerDisplayName } })
        IsAssignableToRole = $(if (${group}) { ${group}.IsAssignableToRole } else { ${sample}.GroupIsAssignableToRole })
        IsDynamicGroup = $(if (${group}) { ${group}.IsDynamicGroup } else { '' })
        GroupId = ${grouped}.Name
    }
}
${groupMembershipSummary} = @(${groupMembershipSummary} | Sort-Object @{Expression={ To-Int $_.RiskScore };Descending=$true}, @{Expression='MemberRows';Descending=$true}, GroupName)

Write-Stage 'Building owner association rows'
${ownerAssociations} = @()
foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.OwnerId } | Group-Object OwnerId)) {
    ${rows} = @(${ownerGroup}.Group)
    if (${rows}.Count -eq 0) { continue }
    ${sample} = ${rows}[0]
    ${ownerAssociations} += [pscustomobject]@{
        OwnerDisplayName = ${sample}.OwnerDisplayName
        OwnerUPN = ${sample}.OwnerUPN
        OwnerType = ${sample}.OwnerType
        OwnedGroupCount = @(${rows}).Count
        OwnedGroups = Join-Unique (@(${rows}) | ForEach-Object { $_.GroupName })
        OwnerId = ${ownerGroup}.Name
    }
}
${ownerAssociations} = @(${ownerAssociations} | Sort-Object @{Expression='OwnedGroupCount';Descending=$true}, OwnerUPN, OwnerDisplayName)

Export-CsvSafe ${groupUserDetail} (Join-Path ${OutputFolder} 'IdentityAudit-GroupUserDetail.csv')
Export-CsvSafe ${userGroupAssociations} (Join-Path ${OutputFolder} 'IdentityAudit-UserGroupAssociations.csv')
Export-CsvSafe ${groupMembershipSummary} (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembershipSummary.csv')
Export-CsvSafe ${ownerAssociations} (Join-Path ${OutputFolder} 'IdentityAudit-OwnerGroupAssociations.csv')

Write-Stage 'Writing detail dashboard'
${css} = @'
<style>
body{font-family:Segoe UI,Arial;margin:32px;background:#f5f7fb;color:#172033}.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}.card,.panel{background:#fff;border:1px solid #d9e0ec;border-radius:12px;padding:14px;margin:14px 0}.value{font-size:28px;font-weight:700}.muted{color:#667}.section-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.count{font-size:13px;color:#52627a;background:#eef3fb;border:1px solid #d8e1ee;padding:6px 10px;border-radius:999px}.search{width:100%;box-sizing:border-box;border:1px solid #ccd6e5;border-radius:10px;padding:10px 12px;margin:8px 0 12px 0;font-size:14px}.table-wrap{max-height:560px;overflow:auto;border:1px solid #edf1f7;border-radius:10px}table{border-collapse:collapse;width:100%;font-size:13px;background:white}td,th{border-bottom:1px solid #edf1f7;padding:7px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8fafc;z-index:1}nav a{display:inline-block;margin-right:10px;margin-bottom:8px;color:#174ea6;text-decoration:none}.small{font-size:12px}
</style>
<script>
function filterTable(input, tableId){
  var q=(input.value||'').toLowerCase();
  var table=document.getElementById(tableId);
  if(!table){return;}
  var rows=table.querySelectorAll('tbody tr');
  rows.forEach(function(row){ row.style.display=row.innerText.toLowerCase().indexOf(q)>=0 ? '' : 'none'; });
}
</script>
'@

${navHtml} = @'
<nav>
<a href="#groupSummary">Group summaries</a>
<a href="#groupUsers">Group to users</a>
<a href="#userGroups">User to groups</a>
<a href="#owners">Owners</a>
<a href="#graphFindings">Graph findings</a>
</nav>
'@

${dashboardPath} = Join-Path ${OutputFolder} 'IdentityAudit-DetailDashboard.html'
${uniqueUsers} = @(${groupUserDetail} | Where-Object { $_.MemberType -eq 'User' } | Select-Object -ExpandProperty MemberId -Unique).Count
${html} = "<!doctype html><html><head><meta charset='utf-8'><title>Identity Detail Dashboard</title>${css}</head><body>"
${html} += '<h1>Identity Detail Dashboard</h1>'
${html} += "<p class='muted'>Searchable group membership and user group-association detail from $(HtmlSafe ${OutputFolder}).</p>"
${html} += ${navHtml}
${html} += "<div class='cards'><div class='card'>Groups with members<div class='value'>$(@(${groupMembershipSummary}).Count)</div></div><div class='card'>Membership rows<div class='value'>$(@(${groupUserDetail}).Count)</div></div><div class='card'>Unique user members<div class='value'>${uniqueUsers}</div></div><div class='card'>User association rows<div class='value'>$(@(${userGroupAssociations}).Count)</div></div><div class='card'>Owner rows<div class='value'>$(@(${ownerAssociations}).Count)</div></div><div class='card'>Privileged paths<div class='value'>$(@(${paths}).Count)</div></div></div>"
${html} += '<a id="groupSummary"></a>' + (New-SearchableTable 'Group membership summaries' 'One row per group, with member-type counts, owners, departments, and risk score.' ${groupMembershipSummary} @('GroupName','GroupCategory','RiskScore','RiskDrivers','MemberRows','UserMembers','GroupMembers','ServicePrincipalMembers','DeviceMembers','Departments','Owners','IsAssignableToRole','IsDynamicGroup') 'tblGroupSummary')
${html} += '<a id="groupUsers"></a>' + (New-SearchableTable 'Group to users and members' 'Every membership row. Search by group, user, UPN, department, owner, member type, or risk driver.' ${groupUserDetail} @('GroupName','GroupCategory','RiskScore','RiskDrivers','Owners','MemberDisplayName','MemberUPN','MemberType','MemberDepartment','MemberJobTitle','MemberCompanyName','MemberAccountEnabled','MembershipMode','IsAssignableToRole','IsDynamicGroup') 'tblGroupUsers')
${html} += '<a id="userGroups"></a>' + (New-SearchableTable 'User/member to group associations' 'One row per member, showing all group associations and owned groups.' ${userGroupAssociations} @('MemberDisplayName','MemberUPN','MemberType','MemberDepartment','MemberJobTitle','MemberCompanyName','MemberAccountEnabled','GroupCount','HighestRiskScore','HighRiskGroupCount','RoleAssignableGroupCount','OwnedGroupCount','Groups','OwnedGroups') 'tblUserGroups')
${html} += '<a id="owners"></a>' + (New-SearchableTable 'Owner to group associations' 'One row per owner, showing all groups they own.' ${ownerAssociations} @('OwnerDisplayName','OwnerUPN','OwnerType','OwnedGroupCount','OwnedGroups') 'tblOwners')
${html} += '<a id="graphFindings"></a>' + (New-SearchableTable 'Privileged path candidates' 'Path candidates from the graph analysis layer.' ${paths} @('StartLabel','StartType','EntryGroup','TargetGroup','HopCount','Path','Risk') 'tblPaths')
${html} += (New-SearchableTable 'Circular group nesting' 'Detected group nesting cycles.' ${cycles} @('Length','Cycle') 'tblCycles')
${html} += (New-SearchableTable 'Nested group chokepoints' 'Groups with nested group edges.' ${nestingStats} @('GroupName','NestedGroupMemberCount','NestedIntoGroupCount','MemberCount','DepartmentCount') 'tblNesting')
${html} += "<p class='muted small'>Generated $(Get-Date). Full detail CSVs are saved in the same output folder.</p></body></html>"
${html} | Out-File ${dashboardPath} -Encoding utf8 -Force
Write-Stage "Detail dashboard written: ${dashboardPath}"

if (${OpenDashboard}) { Invoke-Item ${dashboardPath} }
