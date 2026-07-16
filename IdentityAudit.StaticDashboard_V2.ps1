<#
.SYNOPSIS
Builds Identity Audit Static Dashboard v2 with job-title and department correlation analysis.

.DESCRIPTION
No Node, npm, React, Vite, CDN, internet access, or web server required.
Reads IdentityAudit evidence CSVs and writes one self-contained HTML file.

V2 focus:
- Job title to group correlation
- Compare two job titles
- Click department or job title as a vector against the dataset
- Shift-click users to create a multi-user comparison set
- Export correlation results from the browser
#>

[CmdletBinding()]
param(
    [string] ${OutputRoot} = '.\IdentityAudit-Evidence',
    [string] ${OutputFolder},
    [string] ${OutputFile},
    [switch] ${OpenDashboard}
)

$ErrorActionPreference = 'Stop'

function Write-Stage([string] ${Message}) { Write-Host "[IdentityAudit][StaticV2] ${Message}" -ForegroundColor Cyan }
function Test-Value($Value) { return -not [string]::IsNullOrWhiteSpace([string] ${Value}) }
function Import-CsvSafe([string] ${Path}) { if (-not (Test-Path ${Path})) { return @() }; return @(Import-Csv ${Path} | Where-Object { $null -ne $_ }) }
function Export-CsvSafe($Rows, [string] ${Path}) { ${items} = @(${Rows} | Where-Object { $null -ne $_ }); if (${items}.Count -eq 0) { New-Item -ItemType File -Path ${Path} -Force | Out-Null } else { ${items} | Export-Csv ${Path} -NoTypeInformation } }
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
if (-not (Test-Value ${OutputFile})) { ${OutputFile} = Join-Path ${OutputFolder} 'IdentityAudit-StaticDashboard-v2.html' }

Write-Stage "Reading ${OutputFolder}"

${users} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Users.csv')
${groups} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Groups.csv')
${members} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupMembers.csv')
${owners} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-GroupOwners.csv')
${exceptions} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-Exceptions.csv')
${riskScores} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-RiskScores.csv')
${paths} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-PrivilegedPaths.csv')
${cycles} = Import-CsvSafe (Join-Path ${OutputFolder} 'IdentityAudit-CircularNesting.csv')
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
    Export-CsvSafe ${groupUserDetail} (Join-Path ${OutputFolder} 'IdentityAudit-GroupUserDetail.csv')
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
    Export-CsvSafe ${userGroupAssociations} (Join-Path ${OutputFolder} 'IdentityAudit-UserGroupAssociations.csv')
}

${summary} = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('o')
    dashboardVersion = 'StaticDashboard_v2'
    sourceFolder = ${OutputFolder}
    userCount = @(${users}).Count
    groupCount = @(${groups}).Count
    membershipRowCount = @(${members}).Count
    groupUserDetailRows = @(${groupUserDetail}).Count
    jobTitleCount = @(${groupUserDetail} | Where-Object { Test-Value $_.MemberJobTitle } | Select-Object -ExpandProperty MemberJobTitle -Unique).Count
    departmentCount = @(${groupUserDetail} | Where-Object { Test-Value $_.MemberDepartment } | Select-Object -ExpandProperty MemberDepartment -Unique).Count
    privilegedPathCount = @(${paths}).Count
    cycleCount = @(${cycles}).Count
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
    paths = ${paths}
    cycles = ${cycles}
    groupUserDetail = ${groupUserDetail}
    userGroupAssociations = ${userGroupAssociations}
    groupMembershipSummary = ${groupMembershipSummary}
    ownerGroupAssociations = ${ownerGroupAssociations}
}

Write-Stage 'Embedding data into static V2 HTML'
${json} = ${payload} | ConvertTo-Json -Depth 12
${jsonBytes} = [System.Text.Encoding]::UTF8.GetBytes(${json})
${jsonBase64} = [Convert]::ToBase64String(${jsonBytes})

${htmlTemplate} = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Identity Audit Static Dashboard v2</title>
<style>
:root{font-family:Segoe UI,Arial,sans-serif;color:#172033;background:#f5f7fb}body{margin:0;background:#f5f7fb}.app{max-width:1680px;margin:0 auto;padding:28px}.hero,.panel,.card{background:#fff;border:1px solid #d9e0ec;border-radius:16px;padding:16px;margin:14px 0}.hero{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.muted{color:#667085}.small{font-size:12px}.nav,.toolbar,.chips{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.nav button,.btn,.chip{border:1px solid #cbd5e1;background:#fff;color:#172033;border-radius:10px;padding:9px 12px;cursor:pointer}.nav button.active,.btn.primary,.chip.active{background:#174ea6;color:#fff;border-color:#174ea6}.chip.department.active{background:#047857;border-color:#047857}.chip.title.active{background:#7c3aed;border-color:#7c3aed}.chip.user.active{background:#d97706;border-color:#d97706}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}.twoCol{display:grid;grid-template-columns:1fr 1fr;gap:14px}.threeCol{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.value{font-size:30px;font-weight:750}.label{color:#52627a;font-size:13px}.sectionHead{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.pill{display:inline-block;color:#52627a;background:#eef3fb;border:1px solid #d8e1ee;padding:6px 10px;border-radius:999px;font-size:12px;white-space:nowrap}.search,input,select{border:1px solid #ccd6e5;border-radius:10px;padding:10px 12px;box-sizing:border-box;background:#fff}.search{width:100%;margin:8px 0 12px}.tableWrap{max-height:620px;overflow:auto;border:1px solid #edf1f7;border-radius:10px}table{width:100%;border-collapse:collapse;background:#fff;font-size:13px}td,th{border-bottom:1px solid #edf1f7;padding:7px 9px;text-align:left;vertical-align:top}th{position:sticky;top:0;background:#f8fafc;z-index:1}.clickable{color:#174ea6;text-decoration:underline;cursor:pointer}.selectedRow{background:#fff7ed}.compareBox{border:1px solid #d9e0ec;border-radius:12px;background:#f8fafc;padding:12px;margin:10px 0}.warn{border:1px solid #f5c46b;background:#fff7ed;color:#7c2d12;border-radius:12px;padding:10px 12px;margin:10px 0}pre{white-space:pre-wrap;background:#0f172a;color:#e2e8f0;padding:14px;border-radius:12px;overflow:auto}@media(max-width:1100px){.twoCol,.threeCol{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
  <div class="hero"><div><h1>Identity Audit Static Dashboard v2</h1><p id="source" class="muted"></p></div><span id="generated" class="pill"></span></div>
  <div class="nav" id="nav"></div>
  <div id="content"></div>
</div>
<script>
const encoded='__DATA_BASE64__';
const bytes=Uint8Array.from(atob(encoded), c=>c.charCodeAt(0));
const DATA=JSON.parse(new TextDecoder().decode(bytes));
const pages=[['summary','Summary'],['correlation','Correlation'],['groups','Groups'],['users','Users'],['findings','Findings'],['exports','Exports'],['manifest','Manifest']];
let current='summary';
const state={departments:new Set(),titles:new Set(),users:new Set(),lastUserIndex:null,titleA:'',titleB:''};
const q=(id)=>document.getElementById(id);
const arr=(v)=>Array.isArray(v)?v:[];
const s=(v)=>v===null||v===undefined?'':String(v);
const n=(v)=>{const x=parseInt(s(v),10);return Number.isFinite(x)?x:0};
const norm=(v)=>s(v).trim()||'(blank)';
const esc=(v)=>s(v).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const detail=arr(DATA.groupUserDetail).filter(r=>s(r.MemberType)==='User'||s(r.MemberId));
const groupMap=new Map(arr(DATA.groups).map(g=>[s(g.GroupId),g]));
const riskMap=new Map(arr(DATA.riskScores).map(r=>[s(r.GroupId),r]));
function unique(vals){return [...new Set(vals.map(norm))].filter(x=>x&&x!=='(blank)').sort((a,b)=>a.localeCompare(b))}
function byTitle(title){return detail.filter(r=>norm(r.MemberJobTitle)===norm(title))}
function byDept(dept){return detail.filter(r=>norm(r.MemberDepartment)===norm(dept))}
function byUserIds(ids){return detail.filter(r=>ids.has(s(r.MemberId)))}
function groupsFor(rows){const m=new Map();rows.forEach(r=>{const gid=s(r.GroupId);if(!gid)return;if(!m.has(gid))m.set(gid,[]);m.get(gid).push(r)});return m}
function groupVectorRows(rows){const gm=groupsFor(rows);return [...gm.entries()].map(([gid,rs])=>{const risk=riskMap.get(gid)||{};const g=groupMap.get(gid)||{};const users=[...new Set(rs.map(r=>s(r.MemberId)).filter(Boolean))];const titles=unique(rs.map(r=>r.MemberJobTitle));const depts=unique(rs.map(r=>r.MemberDepartment));return {GroupName:s(rs[0].GroupName)||s(g.GroupName),GroupId:gid,GroupCategory:s(rs[0].GroupCategory)||s(g.GroupCategory),RiskScore:s(risk.RiskScore),RiskDrivers:s(risk.RiskDrivers),VectorUserCount:users.length,VectorMembershipRows:rs.length,Departments:depts.join('; '),JobTitles:titles.join('; '),IsAssignableToRole:s(rs[0].IsAssignableToRole||g.IsAssignableToRole),IsDynamicGroup:s(rs[0].IsDynamicGroup||g.IsDynamicGroup)}}).sort((a,b)=>n(b.RiskScore)-n(a.RiskScore)||n(b.VectorUserCount)-n(a.VectorUserCount)||s(a.GroupName).localeCompare(s(b.GroupName)))}
function titleStats(){return unique(detail.map(r=>r.MemberJobTitle)).map(t=>{const rows=byTitle(t);const groups=new Set(rows.map(r=>s(r.GroupId)).filter(Boolean));const users=new Set(rows.map(r=>s(r.MemberId)).filter(Boolean));return {JobTitle:t,UserCount:users.size,GroupCount:groups.size,MembershipRows:rows.length,HighRiskGroupCount:groupVectorRows(rows).filter(g=>n(g.RiskScore)>=50).length}}).sort((a,b)=>b.GroupCount-a.GroupCount||b.UserCount-a.UserCount)}
function deptStats(){return unique(detail.map(r=>r.MemberDepartment)).map(d=>{const rows=byDept(d);const groups=new Set(rows.map(r=>s(r.GroupId)).filter(Boolean));const users=new Set(rows.map(r=>s(r.MemberId)).filter(Boolean));return {Department:d,UserCount:users.size,GroupCount:groups.size,MembershipRows:rows.length,HighRiskGroupCount:groupVectorRows(rows).filter(g=>n(g.RiskScore)>=50).length}}).sort((a,b)=>b.GroupCount-a.GroupCount||b.UserCount-a.UserCount)}
function csv(rows){if(!rows.length)return'';const cols=[...rows.reduce((set,r)=>{Object.keys(r).forEach(k=>set.add(k));return set},new Set())];const e=(v)=>{const raw=s(v);return /[",\n\r]/.test(raw)?'"'+raw.replace(/"/g,'""')+'"':raw};return [cols.join(','),...rows.map(r=>cols.map(c=>e(r[c])).join(','))].join('\n')}
function download(name,rows){const blob=new Blob([csv(rows)],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();URL.revokeObjectURL(url)}
function table(title,desc,rows,cols,name,max=600){rows=arr(rows);const id='t'+Math.random().toString(16).slice(2);window.__tables[id]=rows;setTimeout(()=>{const input=q(id+'s'),tbl=q(id);if(input&&tbl){input.addEventListener('input',()=>{const v=input.value.toLowerCase();tbl.querySelectorAll('tbody tr').forEach(tr=>tr.style.display=tr.innerText.toLowerCase().includes(v)?'':'none')})}},0);const body=rows.slice(0,max).map((r,idx)=>'<tr data-idx="'+idx+'">'+cols.map(c=>'<td>'+cell(c,r,idx)+'</td>').join('')+'</tr>').join('');return `<section class="panel"><div class="sectionHead"><div><h2>${esc(title)}</h2><p class="muted">${esc(desc||'')}</p></div><div><span class="pill">${rows.length.toLocaleString()} rows</span> <button class="btn" onclick="download('${name}.csv',window.__tables['${id}'])">Export CSV</button></div></div><input id="${id}s" class="search" placeholder="Search this table..."/><div class="tableWrap"><table id="${id}"><thead><tr>${cols.map(c=>'<th>'+esc(c)+'</th>').join('')}</tr></thead><tbody>${body}</tbody></table></div>${rows.length>max?'<p class="muted small">Showing first '+max.toLocaleString()+' rows. Export includes all rows.</p>':''}</section>`}
function cell(c,r,idx){const v=s(r[c]);if(c==='MemberDepartment'||c==='Department')return `<span class="clickable" onclick="toggleDepartment('${encodeURIComponent(v)}')">${esc(v)}</span>`;if(c==='MemberJobTitle'||c==='JobTitle')return `<span class="clickable" onclick="toggleTitle('${encodeURIComponent(v)}')">${esc(v)}</span>`;if(c==='MemberDisplayName'||c==='MemberUPN')return `<span class="clickable" onclick="selectUserFromRow(event,'${encodeURIComponent(s(r.MemberId))}',${idx})">${esc(v)}</span>`;return esc(v)}
window.__tables={};
const cols={groupUser:['GroupName','GroupCategory','RiskScore','RiskDrivers','Owners','MemberDisplayName','MemberUPN','MemberDepartment','MemberJobTitle','MemberType','MemberAccountEnabled','IsAssignableToRole','IsDynamicGroup'],user:['MemberDisplayName','MemberUPN','MemberDepartment','MemberJobTitle','GroupCount','HighestRiskScore','HighRiskGroupCount','RoleAssignableGroupCount','Groups','OwnedGroups'],groupVector:['GroupName','GroupCategory','RiskScore','RiskDrivers','VectorUserCount','VectorMembershipRows','Departments','JobTitles','IsAssignableToRole','IsDynamicGroup'],title:['JobTitle','UserCount','GroupCount','MembershipRows','HighRiskGroupCount'],dept:['Department','UserCount','GroupCount','MembershipRows','HighRiskGroupCount'],compare:['GroupName','Presence','RiskScore','UsersA','UsersB','OnlyAUsers','OnlyBUsers','GroupId']};
function render(){q('source').innerText='Source: '+s(DATA.summary?.sourceFolder);q('generated').innerText='Generated '+s(DATA.summary?.generatedAt)+' | '+s(DATA.summary?.dashboardVersion);q('nav').innerHTML=pages.map(p=>`<button class="${current===p[0]?'active':''}" onclick="current='${p[0]}';render()">${p[1]}</button>`).join('');if(current==='summary')summary();if(current==='correlation')correlation();if(current==='groups')groups();if(current==='users')users();if(current==='findings')findings();if(current==='exports')exportsPage();if(current==='manifest')manifest()}
function cards(items){return '<div class="grid">'+items.map(i=>`<div class="card"><div class="label">${esc(i[0])}</div><div class="value">${esc(i[1])}</div><div class="muted small">${esc(i[2]||'')}</div></div>`).join('')+'</div>'}
function summary(){const sm=DATA.summary||{};q('content').innerHTML=cards([['Users',sm.userCount||0],['Groups',sm.groupCount||0],['Membership rows',sm.membershipRowCount||0],['Detail rows',sm.groupUserDetailRows||0],['Job titles',sm.jobTitleCount||0],['Departments',sm.departmentCount||0],['High-risk groups',sm.highRiskGroupCount||0],['Privileged paths',sm.privilegedPathCount||0]])+table('Top job title vectors','Job titles ranked by group coverage.',titleStats().slice(0,25),cols.title,'top-job-title-vectors')+table('Top department vectors','Departments ranked by group coverage.',deptStats().slice(0,25),cols.dept,'top-department-vectors')}
function vectorRows(){let rows=[];state.titles.forEach(t=>rows=rows.concat(byTitle(t)));state.departments.forEach(d=>rows=rows.concat(byDept(d)));if(state.users.size)rows=rows.concat(byUserIds(state.users));const seen=new Set();return rows.filter(r=>{const k=[s(r.GroupId),s(r.MemberId),s(r.MemberJobTitle),s(r.MemberDepartment)].join('|');if(seen.has(k))return false;seen.add(k);return true})}
function renderChips(){return `<div class="chips"><b>Active vectors:</b> ${[...state.departments].map(d=>`<button class="chip department active" onclick="toggleDepartment('${encodeURIComponent(d)}')">Department: ${esc(d)} x</button>`).join('')} ${[...state.titles].map(t=>`<button class="chip title active" onclick="toggleTitle('${encodeURIComponent(t)}')">Title: ${esc(t)} x</button>`).join('')} ${[...state.users].map(u=>`<button class="chip user active" onclick="toggleUser('${encodeURIComponent(u)}')">UserId: ${esc(u.slice(0,8))} x</button>`).join('')} <button class="btn" onclick="clearVectors()">Clear vectors</button></div>`}
function titleOptions(selected){return titleStats().map(t=>`<option ${t.JobTitle===selected?'selected':''}>${esc(t.JobTitle)}</option>`).join('')}
function correlation(){if(!state.titleA){state.titleA=titleStats()[0]?.JobTitle||''}if(!state.titleB){state.titleB=titleStats()[1]?.JobTitle||state.titleA||''}const active=vectorRows();const gv=groupVectorRows(active);const comp=compareTitles(state.titleA,state.titleB);q('content').innerHTML=`<section class="panel"><h2>Correlation: job title / department / user vectors</h2><p class="muted">Click a department or job title anywhere in the dashboard to add it as a vector. Shift-click users in the tables to build a multi-user comparison set. This is for job title to group correlation and access-pattern review.</p>${renderChips()}<div class="twoCol"><div class="compareBox"><h3>Compare two job titles</h3><label>Job title A<br><select id="titleA" onchange="state.titleA=this.value;correlation()">${titleOptions(state.titleA)}</select></label><br><br><label>Job title B<br><select id="titleB" onchange="state.titleB=this.value;correlation()">${titleOptions(state.titleB)}</select></label><div class="toolbar"><button class="btn primary" onclick="download('job-title-compare.csv',compareTitles(state.titleA,state.titleB))">Export title comparison</button></div><p class="muted small">Similarity: ${similarityLabel(comp)}</p></div><div class="compareBox"><h3>Shift-click user compare</h3><p class="muted">Shift-click names/UPNs in membership tables. Selected users are added as a vector and compared against the same group universe.</p><p><b>Selected users:</b> ${state.users.size}</p><button class="btn" onclick="download('selected-user-vector-groups.csv',groupVectorRows(byUserIds(state.users)))">Export selected-user groups</button></div></div></section>`+table('Job title comparison: '+state.titleA+' vs '+state.titleB,'Common, title-only, and differential group presence.',comp,cols.compare,'job-title-comparison')+table('Active vector group correlation','Groups correlated to selected departments, job titles, and shift-clicked users.',gv,cols.groupVector,'active-vector-group-correlation')+table('Active vector membership rows','Raw rows behind the active vector.',active,cols.groupUser,'active-vector-membership-rows')+table('Clickable job title vectors','Click a job title to add/remove it as a vector.',titleStats(),cols.title,'job-title-vectors')+table('Clickable department vectors','Click a department to add/remove it as a vector.',deptStats(),cols.dept,'department-vectors')}
function compareTitles(a,b){const rowsA=byTitle(a),rowsB=byTitle(b),ga=groupsFor(rowsA),gb=groupsFor(rowsB),ids=new Set([...ga.keys(),...gb.keys()]);return [...ids].map(id=>{const A=ga.get(id)||[],B=gb.get(id)||[],risk=riskMap.get(id)||{};const name=s(A[0]?.GroupName||B[0]?.GroupName||id);const usersA=[...new Set(A.map(r=>s(r.MemberUPN||r.MemberDisplayName)).filter(Boolean))];const usersB=[...new Set(B.map(r=>s(r.MemberUPN||r.MemberDisplayName)).filter(Boolean))];return {GroupName:name,Presence:A.length&&B.length?'Common':A.length?'Only A':'Only B',RiskScore:s(risk.RiskScore),UsersA:usersA.length,UsersB:usersB.length,OnlyAUsers:usersA.filter(x=>!usersB.includes(x)).join('; '),OnlyBUsers:usersB.filter(x=>!usersA.includes(x)).join('; '),GroupId:id}}).sort((x,y)=>presenceRank(x.Presence)-presenceRank(y.Presence)||n(y.RiskScore)-n(x.RiskScore)||s(x.GroupName).localeCompare(s(y.GroupName)))}
function presenceRank(p){return p==='Common'?0:p==='Only A'?1:2}
function similarityLabel(rows){const common=rows.filter(r=>r.Presence==='Common').length,union=rows.length;return union?`${common}/${union} common groups (${Math.round((common/union)*100)}%)`:'No groups found'}
function toggleSet(set,value){if(set.has(value))set.delete(value);else set.add(value)}
function toggleDepartment(v){toggleSet(state.departments,decodeURIComponent(v));current='correlation';render()}
function toggleTitle(v){toggleSet(state.titles,decodeURIComponent(v));current='correlation';render()}
function toggleUser(v){toggleSet(state.users,decodeURIComponent(v));current='correlation';render()}
function selectUserFromRow(event,userId,index){const id=decodeURIComponent(userId);if(!id)return;if(event.shiftKey&&state.lastUserIndex!==null){const rows=detail;const a=Math.min(state.lastUserIndex,index),b=Math.max(state.lastUserIndex,index);for(let i=a;i<=b;i++){const uid=s(rows[i]?.MemberId);if(uid)state.users.add(uid)}}else{toggleSet(state.users,id);state.lastUserIndex=index}current='correlation';render()}
function clearVectors(){state.departments.clear();state.titles.clear();state.users.clear();state.lastUserIndex=null;render()}
function groups(){q('content').innerHTML=table('Group to users and members','Click MemberDepartment or MemberJobTitle to add vectors. Shift-click MemberDisplayName or MemberUPN to add users.',detail,cols.groupUser,'group-user-detail')+table('Group membership summaries','Group summary rows.',arr(DATA.groupMembershipSummary),['GroupName','GroupCategory','RiskScore','RiskDrivers','MemberRows','UserMembers','GroupMembers','Departments','Owners','IsAssignableToRole','IsDynamicGroup'],'group-summary')}
function users(){q('content').innerHTML=table('User/member to group associations','Click departments/titles; shift-click users from membership rows.',arr(DATA.userGroupAssociations),cols.user,'user-group-associations')}
function findings(){q('content').innerHTML=table('Risk-scored groups','Risk scoring output.',arr(DATA.riskScores),['GroupName','RiskScore','RiskDrivers','MemberCount','OwnerCount','DepartmentCount','PrivilegedPathCount','IsAssignableToRole','IsDynamicGroup'],'risk-scores')+table('Privileged paths','Path candidates.',arr(DATA.paths),['StartLabel','StartType','EntryGroup','TargetGroup','HopCount','Path','Risk'],'privileged-paths')+table('Circular nesting','Detected group nesting cycles.',arr(DATA.cycles),['Length','Cycle'],'circular-nesting')}
function exportsPage(){q('content').innerHTML=`<section class="panel"><h2>Exports</h2><p class="muted">Quick exports for correlation review.</p><div class="toolbar"><button class="btn" onclick="download('all-group-user-detail.csv',detail)">Export all group-user detail</button><button class="btn" onclick="download('job-title-vectors.csv',titleStats())">Export job-title vectors</button><button class="btn" onclick="download('department-vectors.csv',deptStats())">Export department vectors</button><button class="btn" onclick="download('active-vector-groups.csv',groupVectorRows(vectorRows()))">Export active vector groups</button><button class="btn" onclick="download('active-vector-membership.csv',vectorRows())">Export active vector membership</button></div></section>`}
function manifest(){q('content').innerHTML='<section class="panel"><h2>Manifest</h2><pre>'+esc(DATA.manifest||'No manifest loaded.')+'</pre></section>'}
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
