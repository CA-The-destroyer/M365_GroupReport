<#
.SYNOPSIS
Builds a no-Node, no-React, self-contained static HTML identity audit dashboard.

.DESCRIPTION
Reads the latest IdentityAudit.Graph_V10 output folder, or a specified OutputFolder, and creates:
- IdentityAudit-StaticDashboard.html

The generated dashboard embeds data directly in the HTML file. It does not require Node, npm, React, Vite, CDN access, or a local web server.
#>

[CmdletBinding()]
param(
    [string] ${OutputRoot} = '.\IdentityAudit-Evidence',
    [string] ${OutputFolder},
    [string] ${OutputFile},
    [switch] ${OpenDashboard}
)

$ErrorActionPreference = 'Stop'

function Write-Stage([string] ${Message}) { Write-Host "[IdentityAudit][Static] ${Message}" -ForegroundColor Cyan }
function Test-Value($Value) { return -not [string]::IsNullOrWhiteSpace([string] ${Value}) }
function Import-CsvSafe([string] ${Path}) { if (-not (Test-Path ${Path})) { return @() }; return @(Import-Csv ${Path} | Where-Object { $null -ne $_ }) }
function To-Int($Value) { try { return [int] ${Value} } catch { return 0 } }
function Set-MapValue($Map, $Key, $Value) { if (Test-Value ${Key}) { ${Map}[[string] ${Key}] = ${Value} } }
function Get-MapValue($Map, $Key) { if (-not (Test-Value ${Key})) { return $null }; if (${Map}.ContainsKey([string] ${Key})) { return ${Map}[[string] ${Key}] }; return $null }
function Join-Unique($Values) { return (@(${Values} | Where-Object { Test-Value $_ } | Select-Object -Unique) -join '; ') }

if (-not (Test-Value ${OutputFolder})) {
    ${latest} = Get-ChildItem -Path ${OutputRoot} -Directory -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not ${latest}) { throw "No run folders found under ${OutputRoot}" }
    ${OutputFolder} = ${latest}.FullName
}

if (-not (Test-Path ${OutputFolder})) { throw "Output folder not found: ${OutputFolder}" }
if (-not (Test-Value ${OutputFile})) { ${OutputFile} = Join-Path ${OutputFolder} 'IdentityAudit-StaticDashboard.html' }

Write-Stage "Reading ${OutputFolder}"

${users} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Users.csv')
${groups} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Groups.csv')
${members} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembers.csv')
${owners} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupOwners.csv')
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

${riskMap} = @{}
foreach (${risk} in ${riskScores}) { Set-MapValue ${riskMap} ${risk}.GroupId ${risk} }
${groupMap} = @{}
foreach (${group} in ${groups}) { Set-MapValue ${groupMap} ${group}.GroupId ${group} }
${ownerByGroup} = @{}
foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) { Set-MapValue ${ownerByGroup} ${ownerGroup}.Name @(${ownerGroup}.Group) }

if (@(${groupUserDetail}).Count -eq 0 -and @(${members}).Count -gt 0) {
    Write-Stage 'Building fallback group-user detail rows'
    ${groupUserDetail} = foreach (${member} in ${members}) {
        ${group} = Get-MapValue ${groupMap} ${member}.GroupId
        ${risk} = Get-MapValue ${riskMap} ${member}.GroupId
        ${ownersForGroup} = @(Get-MapValue ${ownerByGroup} ${member}.GroupId)
        [pscustomobject]@{
            GroupName = ${member}.GroupName
            GroupCategory = ${member}.GroupCategory
            RiskScore = $(if (${risk}) { ${risk}.RiskScore } else { '' })
            RiskDrivers = $(if (${risk}) { ${risk}.RiskDrivers } else { '' })
            Owners = Join-Unique (@(${ownersForGroup}) | ForEach-Object { if ($_.OwnerUPN) { $_.OwnerUPN } else { $_.OwnerDisplayName } })
            MemberDisplayName = ${member}.MemberDisplayName
            MemberUPN = ${member}.MemberUPN
            MemberType = ${member}.MemberType
            MemberDepartment = ${member}.MemberDepartment
            MemberJobTitle = ${member}.MemberJobTitle
            MemberCompanyName = ${member}.MemberCompanyName
            MemberAccountEnabled = ${member}.MemberAccountEnabled
            MembershipMode = ${member}.MembershipMode
            IsAssignableToRole = $(if (${group}) { ${group}.IsAssignableToRole } else { ${member}.GroupIsAssignableToRole })
            IsDynamicGroup = $(if (${group}) { ${group}.IsDynamicGroup } else { '' })
            GroupId = ${member}.GroupId
            MemberId = ${member}.MemberId
        }
    }
}

if (@(${userGroupAssociations}).Count -eq 0 -and @(${members}).Count -gt 0) {
    Write-Stage 'Building fallback user-group association rows'
    ${ownedByPrincipal} = @{}
    foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.OwnerId } | Group-Object OwnerId)) { Set-MapValue ${ownedByPrincipal} ${ownerGroup}.Name @(${ownerGroup}.Group) }
    ${userGroupAssociations} = foreach (${memberGroup} in @(${members} | Where-Object { Test-Value $_.MemberId } | Group-Object MemberId)) {
        ${rows} = @(${memberGroup}.Group)
        if (${rows}.Count -eq 0) { continue }
        ${sample} = ${rows}[0]
        ${ownedGroups} = @(Get-MapValue ${ownedByPrincipal} ${memberGroup}.Name)
        ${riskRows} = foreach (${row} in ${rows}) { ${r} = Get-MapValue ${riskMap} ${row}.GroupId; if (${r}) { ${r} } }
        [pscustomobject]@{
            MemberDisplayName = ${sample}.MemberDisplayName
            MemberUPN = ${sample}.MemberUPN
            MemberType = ${sample}.MemberType
            MemberDepartment = ${sample}.MemberDepartment
            MemberJobTitle = ${sample}.MemberJobTitle
            MemberCompanyName = ${sample}.MemberCompanyName
            MemberAccountEnabled = ${sample}.MemberAccountEnabled
            GroupCount = @(${rows}).Count
            Groups = Join-Unique (@(${rows}) | ForEach-Object { $_.GroupName })
            RoleAssignableGroupCount = @(${rows} | Where-Object { $_.GroupIsAssignableToRole -eq 'True' -or $_.GroupIsAssignableToRole -eq $true }).Count
            HighRiskGroupCount = @(${riskRows} | Where-Object { (To-Int $_.RiskScore) -ge 50 }).Count
            HighestRiskScore = $(if (@(${riskRows}).Count -gt 0) { (@(${riskRows} | Sort-Object { To-Int $_.RiskScore } -Descending | Select-Object -First 1).RiskScore) } else { 0 })
            OwnedGroupCount = @(${ownedGroups}).Count
            OwnedGroups = Join-Unique (@(${ownedGroups}) | ForEach-Object { $_.GroupName })
            MemberId = ${memberGroup}.Name
        }
    }
}

if (@(${groupMembershipSummary}).Count -eq 0 -and @(${members}).Count -gt 0) {
    Write-Stage 'Building fallback group summary rows'
    ${groupMembershipSummary} = foreach (${grouped} in @(${members} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) {
        ${rows} = @(${grouped}.Group)
        ${sample} = ${rows}[0]
        ${group} = Get-MapValue ${groupMap} ${grouped}.Name
        ${risk} = Get-MapValue ${riskMap} ${grouped}.Name
        ${ownersForGroup} = @(Get-MapValue ${ownerByGroup} ${grouped}.Name)
        [pscustomobject]@{
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
}

if (@(${ownerGroupAssociations}).Count -eq 0 -and @(${owners}).Count -gt 0) {
    Write-Stage 'Building fallback owner association rows'
    ${ownerGroupAssociations} = foreach (${ownerGroup} in @(${owners} | Where-Object { Test-Value $_.OwnerId } | Group-Object OwnerId)) {
        ${rows} = @(${ownerGroup}.Group)
        ${sample} = ${rows}[0]
        [pscustomobject]@{
            OwnerDisplayName = ${sample}.OwnerDisplayName
            OwnerUPN = ${sample}.OwnerUPN
            OwnerType = ${sample}.OwnerType
            OwnedGroupCount = @(${rows}).Count
            OwnedGroups = Join-Unique (@(${rows}) | ForEach-Object { $_.GroupName })
            OwnerId = ${ownerGroup}.Name
        }
    }
}

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
    highRiskGroupCount = @(${riskScores} | Where-Object { (To-Int $_.RiskScore) -ge 50 }).Count
}

${payload} = [pscustomobject]@{
    summary = ${summary}
    manifest = ${manifest}
    users = ${users}
    groups = ${groups}
    members = ${members}
    owners = ${owners}
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

Write-Stage 'Embedding data into static HTML'
${json} = ${payload} | ConvertTo-Json -Depth 12
${jsonBytes} = [System.Text.Encoding]::UTF8.GetBytes(${json})
${jsonBase64} = [Convert]::ToBase64String(${jsonBytes})

${htmlTemplate} = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Identity Audit Static Dashboard</title>
<style>
:root{font-family:Segoe UI,Arial,sans-serif;color:#172033;background:#f5f7fb}body{margin:0;background:#f5f7fb}.app{max-width:1680px;margin:0 auto;padding:28px}.hero,.panel,.card{background:#fff;border:1px solid #d9e0ec;border-radius:16px;padding:16px;margin:14px 0}.hero{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.muted{color:#667085}.small{font-size:12px}.nav{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 18px}.nav button,.btn{border:1px solid #cbd5e1;background:#fff;color:#172033;border-radius:10px;padding:9px 12px;cursor:pointer}.nav button.active,.btn.primary{background:#174ea6;color:#fff;border-color:#174ea6}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}.value{font-size:30px;font-weight:750}.label{color:#52627a;font-size:13px}.sectionHead{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.pill{display:inline-block;color:#52627a;background:#eef3fb;border:1px solid #d8e1ee;padding:6px 10px;border-radius:999px;font-size:12px;white-space:nowrap}.search,input,select{border:1px solid #ccd6e5;border-radius:10px;padding:10px 12px;box-sizing:border-box;background:#fff}.search{width:100%;margin:8px 0 12px}.tableWrap{max-height:620px;overflow:auto;border:1px solid #edf1f7;border-radius:10px}table{width:100%;border-collapse:collapse;background:#fff;font-size:13px}td,th{border-bottom:1px solid #edf1f7;padding:7px 9px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8fafc;z-index:1}.toolbar,.filters{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}.filters label{display:flex;flex-direction:column;gap:5px;color:#52627a;font-size:13px}.graphLayout{display:grid;grid-template-columns:minmax(0,1fr)330px;gap:12px}.graphBox{min-height:650px;border:1px solid #d9e0ec;border-radius:14px;background:#fff;overflow:hidden}.side{border:1px solid #d9e0ec;border-radius:14px;background:#f8fafc;padding:12px;max-height:650px;overflow:auto}.pathBox{border:1px solid #f5c46b;background:#fff7ed;color:#7c2d12;border-radius:12px;padding:10px 12px;margin:10px 0;word-break:break-word}svg text{font-size:10px;pointer-events:none}pre{white-space:pre-wrap;background:#0f172a;color:#e2e8f0;padding:14px;border-radius:12px;overflow:auto}@media(max-width:1100px){.graphLayout{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
  <div class="hero"><div><h1>Identity Audit Static Dashboard</h1><p id="source" class="muted"></p></div><span id="generated" class="pill"></span></div>
  <div class="nav" id="nav"></div>
  <div id="content"></div>
</div>
<script>
const encoded='__DATA_BASE64__';
const bytes=Uint8Array.from(atob(encoded), c=>c.charCodeAt(0));
const DATA=JSON.parse(new TextDecoder().decode(bytes));
const pages=[['summary','Summary'],['groups','Groups'],['users','Users'],['owners','Owners'],['findings','Findings'],['graph','Graph'],['exports','Exports'],['manifest','Manifest']];
let current='summary';
let graphState={query:'',nodeType:'All',edgeType:'All',minRisk:0,highOnly:false,selected:null,expanded:new Set(),pathNodes:new Set(),pathEdges:new Set(),pathText:''};
const q=(id)=>document.getElementById(id);
const arr=(v)=>Array.isArray(v)?v:[];
const s=(v)=>v===null||v===undefined?'':String(v);
const n=(v)=>{const x=parseInt(s(v),10);return Number.isFinite(x)?x:0};
const esc=(v)=>s(v).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const highPattern=/(admin|privileged|break.?glass|global administrator|role|security administrator|application administrator|owner|tier.?0)/i;
function isHigh(node){return s(node.Type)==='Group'&&(n(node.RiskScore)>=50||highPattern.test(s(node.Label))||highPattern.test(s(node.RiskDrivers)))}
function csv(rows){if(!rows.length)return'';const cols=[...rows.reduce((set,r)=>{Object.keys(r).forEach(k=>set.add(k));return set},new Set())];const e=(v)=>{const raw=s(v);return /[",\n\r]/.test(raw)?'"'+raw.replace(/"/g,'""')+'"':raw};return [cols.join(','),...rows.map(r=>cols.map(c=>e(r[c])).join(','))].join('\n')}
function download(name,rows){const blob=new Blob([csv(rows)],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();URL.revokeObjectURL(url)}
function table(title,desc,rows,cols,name,max=500){rows=arr(rows);const id='t'+Math.random().toString(16).slice(2);const body=rows.slice(0,max).map(r=>'<tr>'+cols.map(c=>'<td>'+esc(r[c])+'</td>').join('')+'</tr>').join('');setTimeout(()=>{const input=q(id+'s'),tbl=q(id);if(input&&tbl){input.addEventListener('input',()=>{const v=input.value.toLowerCase();tbl.querySelectorAll('tbody tr').forEach(tr=>tr.style.display=tr.innerText.toLowerCase().includes(v)?'':'none')})}},0);return `<section class="panel"><div class="sectionHead"><div><h2>${esc(title)}</h2><p class="muted">${esc(desc||'')}</p></div><div><span class="pill">${rows.length.toLocaleString()} rows</span> <button class="btn" onclick="download('${name}.csv',window.__tables['${id}'])">Export CSV</button></div></div><input id="${id}s" class="search" placeholder="Search this table..."/><div class="tableWrap"><table id="${id}"><thead><tr>${cols.map(c=>'<th>'+esc(c)+'</th>').join('')}</tr></thead><tbody>${body}</tbody></table></div>${rows.length>max?'<p class="muted small">Showing first '+max.toLocaleString()+' rows. Export includes all rows.</p>':''}</section>`}
window.__tables={};function addTable(id,rows){window.__tables[id]=rows}
const cols={group:['GroupName','GroupCategory','RiskScore','RiskDrivers','MemberRows','UserMembers','GroupMembers','ServicePrincipalMembers','DeviceMembers','Departments','Owners','IsAssignableToRole','IsDynamicGroup'],groupUser:['GroupName','GroupCategory','RiskScore','RiskDrivers','Owners','MemberDisplayName','MemberUPN','MemberType','MemberDepartment','MemberJobTitle','MemberCompanyName','MemberAccountEnabled','MembershipMode','IsAssignableToRole','IsDynamicGroup'],user:['MemberDisplayName','MemberUPN','MemberType','MemberDepartment','MemberJobTitle','MemberCompanyName','MemberAccountEnabled','GroupCount','HighestRiskScore','HighRiskGroupCount','RoleAssignableGroupCount','OwnedGroupCount','Groups','OwnedGroups'],owner:['OwnerDisplayName','OwnerUPN','OwnerType','OwnedGroupCount','OwnedGroups'],risk:['GroupName','RiskScore','RiskDrivers','MemberCount','OwnerCount','DepartmentCount','MembershipDensityPct','PrivilegedPathCount','IsAssignableToRole','IsDynamicGroup'],path:['StartLabel','StartType','EntryGroup','TargetGroup','HopCount','Path','Risk']};
function render(){q('source').innerText='Source: '+s(DATA.summary?.sourceFolder);q('generated').innerText='Generated '+s(DATA.summary?.generatedAt);q('nav').innerHTML=pages.map(p=>`<button class="${current===p[0]?'active':''}" onclick="current='${p[0]}';render()">${p[1]}</button>`).join('');if(current==='summary')summary();if(current==='groups')groups();if(current==='users')users();if(current==='owners')owners();if(current==='findings')findings();if(current==='exports')exportsPage();if(current==='manifest')manifest();if(current==='graph')graph()}
function cards(items){return '<div class="grid">'+items.map(i=>`<div class="card"><div class="label">${esc(i[0])}</div><div class="value">${esc(i[1])}</div><div class="muted small">${esc(i[2]||'')}</div></div>`).join('')+'</div>'}
function summary(){const sm=DATA.summary||{};const topRisk=arr(DATA.riskScores).slice().sort((a,b)=>n(b.RiskScore)-n(a.RiskScore)).slice(0,10);const topUsers=arr(DATA.userGroupAssociations).slice().sort((a,b)=>n(b.GroupCount)-n(a.GroupCount)).slice(0,10);q('content').innerHTML=cards([['Groups',sm.groupCount||0],['Users',sm.userCount||0],['Membership rows',sm.membershipRowCount||0],['Ownerless groups',sm.ownerlessGroupCount||0],['High-risk groups',sm.highRiskGroupCount||0,'RiskScore >= 50'],['Privileged paths',sm.privilegedPathCount||0],['Circular nesting',sm.cycleCount||0],['Graph edges',sm.edgeCount||0]])+table('Top risk-scored groups','Highest risk groups.',topRisk,cols.risk,'top-risk')+table('Top users/members by group count','Members with the most group associations.',topUsers,cols.user,'top-users')}
function groups(){q('content').innerHTML=table('Group membership summaries','One row per group with counts, owners, departments, and risk.',arr(DATA.groupMembershipSummary),cols.group,'group-summary')+table('Group to users and members','Every membership row.',arr(DATA.groupUserDetail),cols.groupUser,'group-user-detail')}
function users(){q('content').innerHTML=table('User/member to group associations','One row per user/member showing all group associations.',arr(DATA.userGroupAssociations),cols.user,'user-group-associations')}
function owners(){q('content').innerHTML=table('Owner to group associations','One row per owner showing all owned groups.',arr(DATA.ownerGroupAssociations),cols.owner,'owner-associations')}
function findings(){q('content').innerHTML=table('Risk-scored groups','Risk scoring output.',arr(DATA.riskScores),cols.risk,'risk-scores')+table('Privileged paths','Path candidates from graph analysis.',arr(DATA.paths),cols.path,'privileged-paths')+table('Circular nesting','Detected group nesting cycles.',arr(DATA.cycles),['Length','Cycle'],'circular-nesting')+table('Nested chokepoints','Groups involved in nesting edges.',arr(DATA.nestingStats),['GroupName','NestedGroupMemberCount','NestedIntoGroupCount','MemberCount','DepartmentCount'],'nested-chokepoints')}
function exportsPage(){q('content').innerHTML=`<section class="panel"><h2>Filtered exports</h2><p class="muted">Exports run in the browser against embedded evidence data.</p><div class="filters"><label>Minimum risk<input id="minRisk" type="number" value="50"></label><label>Department contains<input id="dept" placeholder="Finance, IT..."></label></div><div class="toolbar"><button class="btn" onclick="exportRisk()">Export risk filter</button><button class="btn" onclick="exportDept()">Export department membership filter</button><button class="btn" onclick="download('all-group-user-detail.csv',arr(DATA.groupUserDetail))">Export all group-user detail</button><button class="btn" onclick="download('all-user-group-associations.csv',arr(DATA.userGroupAssociations))">Export all user-group associations</button></div></section>`}
function exportRisk(){const min=n(q('minRisk').value);download('groups-risk-'+min+'-plus.csv',arr(DATA.riskScores).filter(r=>n(r.RiskScore)>=min))}
function exportDept(){const d=s(q('dept').value).toLowerCase();download('department-membership-filter.csv',arr(DATA.groupUserDetail).filter(r=>!d||s(r.MemberDepartment).toLowerCase().includes(d)))}
function manifest(){q('content').innerHTML='<section class="panel"><h2>Manifest</h2><pre>'+esc(DATA.manifest||'No manifest loaded.')+'</pre></section>'}
function graph(){q('content').innerHTML=`<section class="panel"><div class="sectionHead"><div><h2>Graph explorer</h2><p class="muted">No libraries, no Node. SVG graph with filters, neighbor expansion, path search, and visible CSV exports.</p></div><span id="graphCount" class="pill"></span></div><div class="filters"><label>Node type<select id="nodeType"><option>All</option><option>Group</option><option>User</option><option>ServicePrincipal</option><option>Device</option><option>DirectoryObject</option></select></label><label>Edge type<select id="edgeType"><option>All</option><option>MemberOf</option><option>OwnsGroup</option></select></label><label>Min risk<input id="gRisk" type="number" value="0"></label><label><input id="gHigh" type="checkbox"> High-value groups only</label></div><div class="toolbar"><input id="gSearch" class="search" placeholder="Search node label, UPN, group name, or ID"><button class="btn" onclick="findNode()">Find node</button><button class="btn" onclick="expandNode()">Expand neighbors</button><button class="btn" onclick="pathToHigh()">Path to high-value group</button><button class="btn" onclick="drawGraph()">Refresh graph</button><button class="btn" onclick="download('visible-static-graph-nodes.csv',window.visibleNodes||[])">Export visible nodes</button><button class="btn" onclick="download('visible-static-graph-edges.csv',window.visibleEdges||[])">Export visible edges</button><button class="btn" onclick="resetGraph()">Reset</button></div><div id="pathBox" class="pathBox" style="display:none"></div><div class="graphLayout"><div class="graphBox"><svg id="graphSvg" width="100%" height="650" viewBox="0 0 1200 650"></svg></div><div class="side"><h3>Selected node</h3><div id="nodeDetail" class="muted">Select a node or search for one.</div></div></div></section>`;['nodeType','edgeType','gRisk','gHigh'].forEach(id=>q(id).addEventListener('change',drawGraph));q('gSearch').addEventListener('keydown',e=>{if(e.key==='Enter')findNode()});drawGraph()}
function graphEdges(){return arr(DATA.edges).map((e,i)=>Object.assign({},e,{_id:'e'+i,_source:s(e.SourceId),_target:s(e.TargetId),_type:s(e.EdgeType)}))}
function nodeMap(){return new Map(arr(DATA.nodes).map(nd=>[s(nd.Id),nd]))}
function neighborMap(edges){const m=new Map();edges.forEach(e=>{if(!m.has(e._source))m.set(e._source,new Set());if(!m.has(e._target))m.set(e._target,new Set());m.get(e._source).add(e._target);m.get(e._target).add(e._source)});return m}
function targetIds(){return new Set(arr(DATA.nodes).filter(isHigh).map(nd=>s(nd.Id)))}
function getVisible(){const nodes=arr(DATA.nodes),edges=graphEdges(),map=nodeMap(),neighbors=neighborMap(edges),search=s(q('gSearch')?.value).toLowerCase(),type=s(q('nodeType')?.value||'All'),edgeType=s(q('edgeType')?.value||'All'),min=n(q('gRisk')?.value),high=!!q('gHigh')?.checked;let seeds=new Set([...graphState.expanded,...graphState.pathNodes]);[...graphState.expanded].forEach(id=>(neighbors.get(id)||new Set()).forEach(x=>seeds.add(x)));let filtered=nodes.filter(nd=>{const id=s(nd.Id),label=s(nd.Label),nt=s(nd.Type),risk=n(nd.RiskScore);return (!search||id.toLowerCase().includes(search)||label.toLowerCase().includes(search)||s(nd.RiskDrivers).toLowerCase().includes(search))&&(type==='All'||nt===type)&&(min<=0||risk>=min||seeds.has(id))&&(!high||isHigh(nd)||seeds.has(id))});filtered.sort((a,b)=>n(b.RiskScore)-n(a.RiskScore));let selected=(search||type!=='All'||min>0||high||seeds.size)?filtered.slice(0,550):filtered.slice(0,300);const byId=new Map(selected.map(nd=>[s(nd.Id),nd]));seeds.forEach(id=>{const nd=map.get(id);if(nd)byId.set(id,nd)});selected=[...byId.values()].slice(0,650);const allowed=new Set(selected.map(nd=>s(nd.Id)));let visibleEdges=edges.filter(e=>((allowed.has(e._source)&&allowed.has(e._target)&&(edgeType==='All'||e._type===edgeType))||graphState.pathEdges.has(e._id))).slice(0,1500);return {nodes:selected,edges:visibleEdges,map}}
function drawGraph(){const svg=q('graphSvg');if(!svg)return;const {nodes,edges,map}=getVisible();window.visibleNodes=nodes;window.visibleEdges=edges;q('graphCount').innerText=nodes.length+' nodes / '+edges.length+' edges';svg.innerHTML='';const W=1200,H=650,cx=W/2,cy=H/2,r=Math.min(W,H)*0.42;const pos=new Map();nodes.forEach((nd,i)=>{const angle=2*Math.PI*i/Math.max(1,nodes.length);const risk=n(nd.RiskScore);const rr=r-(Math.min(risk,100)*1.8);pos.set(s(nd.Id),{x:cx+Math.cos(angle)*rr,y:cy+Math.sin(angle)*rr})});edges.forEach(e=>{const a=pos.get(e._source),b=pos.get(e._target);if(!a||!b)return;const line=document.createElementNS('http://www.w3.org/2000/svg','line');line.setAttribute('x1',a.x);line.setAttribute('y1',a.y);line.setAttribute('x2',b.x);line.setAttribute('y2',b.y);line.setAttribute('stroke',graphState.pathEdges.has(e._id)?'#f59e0b':'#a7b1c2');line.setAttribute('stroke-width',graphState.pathEdges.has(e._id)?'4':'1');svg.appendChild(line)});nodes.forEach(nd=>{const id=s(nd.Id),p=pos.get(id);if(!p)return;const g=document.createElementNS('http://www.w3.org/2000/svg','g');g.setAttribute('cursor','pointer');g.onclick=()=>selectNode(id);const c=document.createElementNS('http://www.w3.org/2000/svg','circle');c.setAttribute('cx',p.x);c.setAttribute('cy',p.y);c.setAttribute('r',graphState.pathNodes.has(id)?16:12);c.setAttribute('fill',s(nd.Type)==='Group'?'#7c3aed':s(nd.Type)==='User'?'#059669':s(nd.Type)==='ServicePrincipal'?'#d97706':'#4f6bed');c.setAttribute('stroke',graphState.selected===id?'#111827':n(nd.RiskScore)>=50?'#dc2626':graphState.pathNodes.has(id)?'#f59e0b':'#fff');c.setAttribute('stroke-width',graphState.selected===id?'5':n(nd.RiskScore)>=50?'4':'2');const t=document.createElementNS('http://www.w3.org/2000/svg','text');t.setAttribute('x',p.x+14);t.setAttribute('y',p.y+4);t.textContent=s(nd.Label).slice(0,44);g.appendChild(c);g.appendChild(t);svg.appendChild(g)});if(graphState.selected)showNode(graphState.selected,map)}
function selectNode(id){graphState.selected=id;drawGraph()}
function showNode(id,map){const nd=map.get(id);q('nodeDetail').innerHTML=nd?Object.entries(nd).map(([k,v])=>'<p><b>'+esc(k)+':</b> '+esc(v)+'</p>').join(''):'No node selected.'}
function findNode(){const term=s(q('gSearch').value).toLowerCase();if(!term)return;const nd=arr(DATA.nodes).find(x=>s(x.Id).toLowerCase().includes(term)||s(x.Label).toLowerCase().includes(term));if(nd){graphState.selected=s(nd.Id);graphState.expanded.add(s(nd.Id));drawGraph()}}
function expandNode(){if(!graphState.selected)return;graphState.expanded.add(graphState.selected);const neighbors=neighborMap(graphEdges());(neighbors.get(graphState.selected)||new Set()).forEach(x=>graphState.expanded.add(x));drawGraph()}
function shortest(start){const edges=graphEdges(),targets=targetIds(),adj=new Map();edges.forEach(e=>{if(!adj.has(e._source))adj.set(e._source,[]);adj.get(e._source).push({next:e._target,edge:e._id})});const queue=[{id:start,nodes:[start],edgeIds:[]}],seen=new Set([start]);while(queue.length){const cur=queue.shift();if(targets.has(cur.id)&&cur.id!==start)return cur;if(cur.nodes.length>10)continue;(adj.get(cur.id)||[]).forEach(nxt=>{if(seen.has(nxt.next))return;seen.add(nxt.next);queue.push({id:nxt.next,nodes:[...cur.nodes,nxt.next],edgeIds:[...cur.edgeIds,nxt.edge]})})}return null}
function pathToHigh(){if(!graphState.selected)return;const res=shortest(graphState.selected);graphState.pathNodes=new Set();graphState.pathEdges=new Set();if(!res){q('pathBox').style.display='block';q('pathBox').innerHTML='<b>Path result:</b> No directed path to a high-risk/high-value group was found within 10 hops.';drawGraph();return}res.nodes.forEach(x=>{graphState.pathNodes.add(x);graphState.expanded.add(x)});res.edgeIds.forEach(x=>graphState.pathEdges.add(x));const map=nodeMap();q('pathBox').style.display='block';q('pathBox').innerHTML='<b>Path result:</b> '+res.nodes.map(id=>esc(s(map.get(id)?.Label)||id)).join(' -> ');drawGraph()}
function resetGraph(){graphState={query:'',nodeType:'All',edgeType:'All',minRisk:0,highOnly:false,selected:null,expanded:new Set(),pathNodes:new Set(),pathEdges:new Set(),pathText:''};graph()}
render();
</script>
</body>
</html>
'@

${html} = ${htmlTemplate}.Replace('__DATA_BASE64__', ${jsonBase64})
${outDir} = Split-Path -Parent ${OutputFile}
if (Test-Value ${outDir}) { New-Item -ItemType Directory -Path ${outDir} -Force | Out-Null }
${html} | Out-File ${OutputFile} -Encoding utf8 -Force
Write-Stage "Wrote ${OutputFile}"

if (${OpenDashboard}) { Invoke-Item ${OutputFile} }
