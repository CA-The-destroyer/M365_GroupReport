<#
.SYNOPSIS
Exports Microsoft Entra ID group membership, department mapping, group density, owners, and an HTML dashboard.

.DESCRIPTION
Cache-backed Microsoft Graph identity audit. The script refreshes users, groups, group memberships, and owners into JSON cache files, then builds CSV evidence and a local HTML dashboard from cache. This avoids re-reading the full tenant every time.
#>

[CmdletBinding()]
param(
    [string] ${TenantId},
    [string] ${ClientId},
    [string] ${CertificateThumbprint},
    [string] ${OutputRoot} = ".\IdentityAudit-Evidence",
    [string] ${CacheRoot} = ".\IdentityAudit-Cache",
    [string] ${GroupIdsFile},
    [switch] ${IncludeTransitiveMembership},
    [switch] ${SecurityOnly},
    [switch] ${MailEnabledSecurityOnly},
    [switch] ${DistributionListOnly},
    [switch] ${Microsoft365Only},
    [switch] ${IsEmpty},
    [int] ${MinGroupMembersCount} = 0,
    [decimal] ${HighDensityPctThreshold} = 5.0,
    [int] ${CacheMaxAgeHours} = 168,
    [switch] ${Menu},
    [switch] ${RefreshUsers},
    [switch] ${RefreshGroups},
    [switch] ${RefreshMemberships},
    [switch] ${RefreshOwners},
    [switch] ${RefreshAll},
    [switch] ${UseCacheOnly},
    [switch] ${SkipOwners},
    [switch] ${InstallModules},
    [switch] ${OpenDashboard}
)

${ErrorActionPreference} = "Stop"

function Write-Stage { param([string] ${Message}) Write-Host "[IdentityAudit] ${Message}" -ForegroundColor Cyan }
function Write-Warn { param([string] ${Message}) Write-Host "[IdentityAudit][WARN] ${Message}" -ForegroundColor Yellow }

function Ensure-Module {
    param([Parameter(Mandatory)] [string] ${Name})
    if (-not (Get-Module -ListAvailable -Name ${Name})) {
        if (${InstallModules}.IsPresent) {
            Write-Stage "Installing ${Name}"
            Install-Module ${Name} -Scope CurrentUser -Force -AllowClobber
        }
        else { throw "Required module '${Name}' is not installed. Re-run with -InstallModules or install it manually." }
    }
    Import-Module ${Name} -ErrorAction Stop
}

function Get-P {
    param(${Object}, [string] ${Name})
    if ($null -eq ${Object}) { return $null }
    try {
        ${Prop} = ${Object}.PSObject.Properties[${Name}]
        if ($null -ne ${Prop}) { return ${Prop}.Value }
    } catch { }
    try { return ${Object}[${Name}] } catch { }
    try {
        ${Additional} = ${Object}.AdditionalProperties
        if ($null -ne ${Additional}) {
            try { return ${Additional}[${Name}] } catch { }
        }
    } catch { }
    return $null
}

function Invoke-GraphGetJson {
    param([Parameter(Mandatory)] [string] ${Uri})
    try {
        ${Response} = Invoke-MgGraphRequest -Method GET -Uri ${Uri} -OutputType HttpResponseMessage
        ${Content} = ${Response}.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if ([string]::IsNullOrWhiteSpace(${Content})) { return $null }
        return ${Content} | ConvertFrom-Json -Depth 100
    }
    catch {
        throw "Graph GET failed. Uri='${Uri}'. Error='$($_.Exception.Message)'"
    }
}

function Invoke-GraphPagedJson {
    param([Parameter(Mandatory)] [string] ${Uri})
    ${Rows} = New-Object System.Collections.Generic.List[object]
    ${NextUri} = ${Uri}
    while (-not [string]::IsNullOrWhiteSpace([string] ${NextUri})) {
        ${Page} = Invoke-GraphGetJson -Uri ${NextUri}
        ${Values} = Get-P -Object ${Page} -Name "value"
        foreach (${Item} in @(${Values})) { if ($null -ne ${Item}) { ${Rows}.Add(${Item}) } }
        ${NextUri} = Get-P -Object ${Page} -Name "@odata.nextLink"
    }
    return ${Rows}.ToArray()
}

function Save-CacheJson { param(${Rows}, [string] ${Path}) @(${Rows}) | ConvertTo-Json -Depth 100 | Out-File -FilePath ${Path} -Encoding utf8 }
function Read-CacheJson { param([string] ${Path}) if (-not (Test-Path ${Path})) { return @() } ${Raw}=Get-Content -Path ${Path} -Raw; if ([string]::IsNullOrWhiteSpace(${Raw})) { return @() } return @(${Raw} | ConvertFrom-Json -Depth 100) }

function Test-CacheFresh {
    param([string] ${Path})
    if (-not (Test-Path ${Path})) { return $false }
    if (${CacheMaxAgeHours} -le 0) { return $true }
    ${AgeHours} = ((Get-Date) - (Get-Item ${Path}).LastWriteTime).TotalHours
    return (${AgeHours} -le ${CacheMaxAgeHours})
}

function Get-DirectoryObjectType {
    param(${Object})
    ${TypeValue} = [string](Get-P -Object ${Object} -Name "@odata.type")
    switch -Regex (${TypeValue}) {
        "user$"             { return "User" }
        "group$"            { return "Group" }
        "servicePrincipal$" { return "ServicePrincipal" }
        "device$"           { return "Device" }
        "orgContact$"       { return "Contact" }
        "directoryRole$"    { return "DirectoryRole" }
        default             { return "DirectoryObject" }
    }
}

function Get-DisplayName { param(${Object}) ${V}=Get-P ${Object} "displayName"; if([string]::IsNullOrWhiteSpace([string]${V})){${V}=Get-P ${Object} "userPrincipalName"}; if([string]::IsNullOrWhiteSpace([string]${V})){${V}=Get-P ${Object} "mail"}; return ${V} }

function Get-GroupCategory {
    param(${Group})
    ${GroupTypes} = Get-P ${Group} "groupTypes"
    ${GroupTypesText} = if (${GroupTypes}) { @(${GroupTypes}) -join ";" } else { "" }
    ${SecurityEnabled} = Get-P ${Group} "securityEnabled"
    ${MailEnabled} = Get-P ${Group} "mailEnabled"
    if (${GroupTypesText} -match "Unified") { return "Microsoft365" }
    if (${SecurityEnabled} -eq $true -and ${MailEnabled} -eq $true) { return "MailEnabledSecurity" }
    if (${SecurityEnabled} -eq $true -and ${MailEnabled} -ne $true) { return "Security" }
    if (${MailEnabled} -eq $true -and ${SecurityEnabled} -ne $true) { return "DistributionList" }
    return "Other"
}

function ConvertTo-Percent { param([decimal]${Numerator},[decimal]${Denominator}) if(${Denominator} -le 0){return [decimal]0}; return [math]::Round((${Numerator}/${Denominator})*100,2) }
function HtmlEncode { param(${Value}) if($null -eq ${Value}){return ""}; return [System.Net.WebUtility]::HtmlEncode([string]${Value}) }

function HtmlTable {
    param([string]${Title}, [object[]]${Rows}, [string[]]${Columns}, [int]${MaxRows}=25)
    ${Html}="<section class='panel'><h2>$(HtmlEncode ${Title})</h2>"
    if(-not ${Rows} -or ${Rows}.Count -eq 0){ return ${Html}+"<p class='muted'>No records found.</p></section>" }
    ${Html}+="<table><thead><tr>"
    foreach(${C} in ${Columns}){ ${Html}+="<th>$(HtmlEncode ${C})</th>" }
    ${Html}+="</tr></thead><tbody>"
    foreach(${R} in (${Rows}|Select-Object -First ${MaxRows})){
        ${Html}+="<tr>"
        foreach(${C} in ${Columns}){ ${V}=""; if(${R}.PSObject.Properties.Name -contains ${C}){${V}=${R}.${C}}; ${Html}+="<td>$(HtmlEncode ${V})</td>" }
        ${Html}+="</tr>"
    }
    return ${Html}+"</tbody></table></section>"
}

function Show-IdentityAuditMenu {
    Write-Host ""
    Write-Host "Identity Audit v1" -ForegroundColor Cyan
    Write-Host "1. Generate report from cache; refresh missing or expired cache"
    Write-Host "2. Refresh users cache, then generate report"
    Write-Host "3. Refresh groups cache, memberships, and owners, then generate report"
    Write-Host "4. Refresh memberships cache, then generate report"
    Write-Host "5. Refresh owners cache, then generate report"
    Write-Host "6. Refresh all cache, then generate report"
    Write-Host "7. Generate report from cache only"
    Write-Host "Q. Quit"
    return (Read-Host "Select an option")
}

function Connect-IdentityAuditGraph {
    Ensure-Module -Name Microsoft.Graph.Authentication
    Write-Stage "Connecting to Microsoft Graph"
    if (-not [string]::IsNullOrWhiteSpace(${TenantId}) -and -not [string]::IsNullOrWhiteSpace(${ClientId}) -and -not [string]::IsNullOrWhiteSpace(${CertificateThumbprint})) {
        Connect-MgGraph -TenantId ${TenantId} -ClientId ${ClientId} -CertificateThumbprint ${CertificateThumbprint} -NoWelcome
        return "AppOnlyCertificate"
    }
    Connect-MgGraph -Scopes @("User.Read.All","Group.Read.All","GroupMember.Read.All","Directory.Read.All") -NoWelcome
    return "Delegated"
}

function Get-UsersFromGraph {
    Write-Stage "Refreshing users cache"
    ${Uri}='/v1.0/users?$select=id,displayName,userPrincipalName,mail,department,jobTitle,companyName,accountEnabled,userType,employeeId,createdDateTime&$top=999'
    return @(Invoke-GraphPagedJson -Uri ${Uri})
}

function Get-GroupsFromGraph {
    Write-Stage "Refreshing groups cache"
    ${AdvancedUri}='/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility&$top=999'
    try { return @(Invoke-GraphPagedJson -Uri ${AdvancedUri}) }
    catch {
        Write-Warn "Advanced group properties failed. Retrying with baseline properties. $($_.Exception.Message)"
        ${BasicUri}='/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime&$top=999'
        return @(Invoke-GraphPagedJson -Uri ${BasicUri})
    }
}

function Get-MembersFromGraph {
    param([object[]]${Groups})
    Write-Stage "Refreshing group membership cache"
    ${Rows}=New-Object System.Collections.Generic.List[object]
    ${Mode}=if(${IncludeTransitiveMembership}.IsPresent){"Transitive"}else{"Direct"}
    ${Count}=0
    foreach(${G} in ${Groups}){
        ${Count}++
        ${Gid}=Get-P ${G} "id"; ${Gname}=Get-P ${G} "displayName"
        Write-Progress -Activity "Refreshing memberships" -Status ${Gname} -PercentComplete ((${Count}/[math]::Max(1,${Groups}.Count))*100)
        ${Endpoint}=if(${IncludeTransitiveMembership}.IsPresent){"transitiveMembers"}else{"members"}
        try { ${Members}=@(Invoke-GraphPagedJson -Uri "/v1.0/groups/${Gid}/${Endpoint}?`$select=id,displayName,userPrincipalName,mail&`$top=999") }
        catch { Write-Warn "Member read failed for group '${Gname}' (${Gid}): $($_.Exception.Message)"; ${Members}=@() }
        foreach(${M} in ${Members}){
            ${Rows}.Add([pscustomobject]@{
                GroupId=${Gid}; GroupName=${Gname}; GroupMail=(Get-P ${G} "mail"); GroupCategory=(Get-GroupCategory ${G});
                GroupSecurityEnabled=(Get-P ${G} "securityEnabled"); GroupMailEnabled=(Get-P ${G} "mailEnabled");
                GroupTypes=(if((Get-P ${G} "groupTypes")){@(Get-P ${G} "groupTypes") -join ";"}else{""});
                GroupIsAssignableToRole=(Get-P ${G} "isAssignableToRole"); GroupMembershipRule=(Get-P ${G} "membershipRule");
                MembershipRuleState=(Get-P ${G} "membershipRuleProcessingState"); MembershipMode=${Mode};
                MemberId=(Get-P ${M} "id"); MemberDisplayName=(Get-DisplayName ${M}); MemberUPN=(Get-P ${M} "userPrincipalName");
                MemberMail=(Get-P ${M} "mail"); MemberType=(Get-DirectoryObjectType ${M})
            })
        }
    }
    Write-Progress -Activity "Refreshing memberships" -Completed
    return ${Rows}.ToArray()
}

function Get-OwnersFromGraph {
    param([object[]]${Groups})
    Write-Stage "Refreshing group owner cache"
    ${Rows}=New-Object System.Collections.Generic.List[object]
    ${Count}=0
    foreach(${G} in ${Groups}){
        ${Count}++
        ${Gid}=Get-P ${G} "id"; ${Gname}=Get-P ${G} "displayName"
        Write-Progress -Activity "Refreshing owners" -Status ${Gname} -PercentComplete ((${Count}/[math]::Max(1,${Groups}.Count))*100)
        try { ${Owners}=@(Invoke-GraphPagedJson -Uri "/v1.0/groups/${Gid}/owners?`$select=id,displayName,userPrincipalName,mail&`$top=999") }
        catch { Write-Warn "Owner read failed for group '${Gname}' (${Gid}): $($_.Exception.Message)"; ${Owners}=@() }
        foreach(${O} in ${Owners}){
            ${Rows}.Add([pscustomobject]@{ GroupId=${Gid}; GroupName=${Gname}; GroupCategory=(Get-GroupCategory ${G}); OwnerId=(Get-P ${O} "id"); OwnerDisplayName=(Get-DisplayName ${O}); OwnerUPN=(Get-P ${O} "userPrincipalName"); OwnerType=(Get-DirectoryObjectType ${O}) })
        }
    }
    Write-Progress -Activity "Refreshing owners" -Completed
    return ${Rows}.ToArray()
}

New-Item -ItemType Directory -Path ${CacheRoot} -Force | Out-Null
${UsersCache}=Join-Path ${CacheRoot} "users.json"
${GroupsCache}=Join-Path ${CacheRoot} "groups.json"
${MembersCache}=Join-Path ${CacheRoot} (if(${IncludeTransitiveMembership}.IsPresent){"members.transitive.json"}else{"members.direct.json"})
${OwnersCache}=Join-Path ${CacheRoot} "owners.json"

${DoRefreshUsers}=[bool]${RefreshUsers}.IsPresent
${DoRefreshGroups}=[bool]${RefreshGroups}.IsPresent
${DoRefreshMembers}=[bool]${RefreshMemberships}.IsPresent
${DoRefreshOwners}=[bool]${RefreshOwners}.IsPresent
${DoRefreshAll}=[bool]${RefreshAll}.IsPresent
${UseCacheOnlyLocal}=[bool]${UseCacheOnly}.IsPresent

if(${Menu}.IsPresent){
    ${Choice}=Show-IdentityAuditMenu
    switch -Regex (${Choice}) {
        '^1$' { }
        '^2$' { ${DoRefreshUsers}=$true }
        '^3$' { ${DoRefreshGroups}=$true; ${DoRefreshMembers}=$true; ${DoRefreshOwners}=$true }
        '^4$' { ${DoRefreshMembers}=$true }
        '^5$' { ${DoRefreshOwners}=$true }
        '^6$' { ${DoRefreshAll}=$true }
        '^7$' { ${UseCacheOnlyLocal}=$true }
        '^[qQ]$' { return }
        default { throw "Invalid menu option: ${Choice}" }
    }
}

if(${DoRefreshAll}){ ${DoRefreshUsers}=$true; ${DoRefreshGroups}=$true; ${DoRefreshMembers}=$true; ${DoRefreshOwners}=$true }
if(-not (Test-CacheFresh ${UsersCache})){ ${DoRefreshUsers}=$true }
if(-not (Test-CacheFresh ${GroupsCache})){ ${DoRefreshGroups}=$true }
if(-not (Test-CacheFresh ${MembersCache})){ ${DoRefreshMembers}=$true }
if(-not ${SkipOwners}.IsPresent -and -not (Test-CacheFresh ${OwnersCache})){ ${DoRefreshOwners}=$true }
if(${DoRefreshGroups}){ ${DoRefreshMembers}=$true; if(-not ${SkipOwners}.IsPresent){${DoRefreshOwners}=$true} }

${ExecutionMode}="CacheOnly"
try{
    if(${UseCacheOnlyLocal}){
        if(-not (Test-Path ${UsersCache}) -or -not (Test-Path ${GroupsCache}) -or -not (Test-Path ${MembersCache})){ throw "Cache-only mode requested, but required cache files do not exist." }
    }
    elseif(${DoRefreshUsers} -or ${DoRefreshGroups} -or ${DoRefreshMembers} -or ${DoRefreshOwners}){
        ${ExecutionMode}=Connect-IdentityAuditGraph
    }

    if(${DoRefreshUsers} -and -not ${UseCacheOnlyLocal}){ Save-CacheJson -Rows (Get-UsersFromGraph) -Path ${UsersCache} }
    if(${DoRefreshGroups} -and -not ${UseCacheOnlyLocal}){ Save-CacheJson -Rows (Get-GroupsFromGraph) -Path ${GroupsCache} }

    ${RawUsers}=Read-CacheJson ${UsersCache}
    ${RawGroups}=Read-CacheJson ${GroupsCache}

    if(${DoRefreshMembers} -and -not ${UseCacheOnlyLocal}){ Save-CacheJson -Rows (Get-MembersFromGraph -Groups ${RawGroups}) -Path ${MembersCache} }
    if(-not ${SkipOwners}.IsPresent -and ${DoRefreshOwners} -and -not ${UseCacheOnlyLocal}){ Save-CacheJson -Rows (Get-OwnersFromGraph -Groups ${RawGroups}) -Path ${OwnersCache} }

    ${Users}=@(foreach(${U} in ${RawUsers}){[pscustomobject]@{Id=(Get-P ${U} "id");DisplayName=(Get-P ${U} "displayName");UserPrincipalName=(Get-P ${U} "userPrincipalName");Mail=(Get-P ${U} "mail");Department=(Get-P ${U} "department");JobTitle=(Get-P ${U} "jobTitle");CompanyName=(Get-P ${U} "companyName");AccountEnabled=(Get-P ${U} "accountEnabled");UserType=(Get-P ${U} "userType");EmployeeId=(Get-P ${U} "employeeId");CreatedDateTime=(Get-P ${U} "createdDateTime")}})
    ${Groups}=@(foreach(${G} in ${RawGroups}){${G}})

    if(-not [string]::IsNullOrWhiteSpace(${GroupIdsFile})){
        ${Wanted}=@{}; Get-Content ${GroupIdsFile} | Where-Object { -not [string]::IsNullOrWhiteSpace([string]${_}) } | ForEach-Object { ${Wanted}[${_}.Trim()]=$true }
        ${Groups}=@(${Groups}|Where-Object{${Wanted}.ContainsKey((Get-P ${_} "id"))})
    }

    ${FilteredGroupIds}=@{}
    ${FilteredGroups}=@(foreach(${G} in ${Groups}){${Cat}=Get-GroupCategory ${G}; if(${SecurityOnly}.IsPresent -and ${Cat} -ne "Security"){continue}; if(${MailEnabledSecurityOnly}.IsPresent -and ${Cat} -ne "MailEnabledSecurity"){continue}; if(${DistributionListOnly}.IsPresent -and ${Cat} -ne "DistributionList"){continue}; if(${Microsoft365Only}.IsPresent -and ${Cat} -ne "Microsoft365"){continue}; ${FilteredGroupIds}[(Get-P ${G} "id")]=$true; ${G}})

    ${UsersById}=@{}; foreach(${U} in ${Users}){ if(-not [string]::IsNullOrWhiteSpace([string]${U}.Id)){${UsersById}[${U}.Id]=${U}} }
    ${MemberRowsCached}=Read-CacheJson ${MembersCache}
    ${OwnerRows}=if(${SkipOwners}.IsPresent){@()}else{@(Read-CacheJson ${OwnersCache}|Where-Object{${FilteredGroupIds}.ContainsKey(${_}.GroupId)})}

    ${MemberRows}=@(foreach(${M} in ${MemberRowsCached}){
        if(-not ${FilteredGroupIds}.ContainsKey(${M}.GroupId)){continue}
        ${Profile}=if(${UsersById}.ContainsKey(${M}.MemberId)){${UsersById}[${M}.MemberId]}else{$null}
        [pscustomobject]@{
            GroupId=${M}.GroupId; GroupName=${M}.GroupName; GroupMail=${M}.GroupMail; GroupCategory=${M}.GroupCategory; GroupSecurityEnabled=${M}.GroupSecurityEnabled; GroupMailEnabled=${M}.GroupMailEnabled; GroupTypes=${M}.GroupTypes; GroupIsAssignableToRole=${M}.GroupIsAssignableToRole; GroupMembershipRule=${M}.GroupMembershipRule; MembershipRuleState=${M}.MembershipRuleState; MembershipMode=${M}.MembershipMode;
            MemberId=${M}.MemberId; MemberDisplayName=${M}.MemberDisplayName; MemberUPN=(if([string]::IsNullOrWhiteSpace([string]${M}.MemberUPN) -and $null -ne ${Profile}){${Profile}.UserPrincipalName}else{${M}.MemberUPN}); MemberMail=(if([string]::IsNullOrWhiteSpace([string]${M}.MemberMail) -and $null -ne ${Profile}){${Profile}.Mail}else{${M}.MemberMail}); MemberType=${M}.MemberType;
            MemberDepartment=(if($null -ne ${Profile}){${Profile}.Department}else{""}); MemberJobTitle=(if($null -ne ${Profile}){${Profile}.JobTitle}else{""}); MemberCompanyName=(if($null -ne ${Profile}){${Profile}.CompanyName}else{""}); MemberAccountEnabled=(if($null -ne ${Profile}){${Profile}.AccountEnabled}else{$null}); MemberUserType=(if($null -ne ${Profile}){${Profile}.UserType}else{""}); MemberEmployeeId=(if($null -ne ${Profile}){${Profile}.EmployeeId}else{""})
        }
    })

    ${RunId}=Get-Date -Format "yyyyMMdd-HHmmss"; ${OutputDir}=Join-Path ${OutputRoot} ${RunId}; New-Item -ItemType Directory -Path ${OutputDir} -Force|Out-Null
    ${TotalMembershipRows}=@(${MemberRows}).Count; ${EnabledUserCount}=@(${Users}|Where-Object{${_}.AccountEnabled -eq $true}).Count; ${TotalUserCount}=@(${Users}).Count; ${NoDepartmentUserCount}=@(${Users}|Where-Object{[string]::IsNullOrWhiteSpace([string]${_}.Department)}).Count
    ${OwnerCountByGroupId}=@{}; foreach(${OG} in (@(${OwnerRows})|Group-Object GroupId)){${OwnerCountByGroupId}[${OG}.Name]=${OG}.Count}
    ${GroupRows}=New-Object System.Collections.Generic.List[object]

    foreach(${G} in ${FilteredGroups}){
        ${Gid}=Get-P ${G} "id"; ${RowsForGroup}=@(${MemberRows}|Where-Object{${_}.GroupId -eq ${Gid}}); ${UserMembers}=@(${RowsForGroup}|Where-Object{${_}.MemberType -eq "User"}); ${Depts}=@(${UserMembers}|Where-Object{-not [string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment)}|Select-Object -ExpandProperty MemberDepartment -Unique); ${DeptGroups}=@(${UserMembers}|Where-Object{-not [string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment)}|Group-Object MemberDepartment|Sort-Object -Property @{Expression="Count";Descending=$true}); ${PrimaryDept}=""; ${PrimaryDeptCount}=0; if(${DeptGroups}.Count -gt 0){${PrimaryDept}=${DeptGroups}[0].Name;${PrimaryDeptCount}=${DeptGroups}[0].Count}
        ${MemberCount}=${RowsForGroup}.Count; ${UniqueUserCount}=@(${UserMembers}|Select-Object -ExpandProperty MemberId -Unique).Count; ${ActiveUserCount}=@(${UserMembers}|Where-Object{${_}.MemberAccountEnabled -eq $true}|Select-Object -ExpandProperty MemberId -Unique).Count; ${Density}=ConvertTo-Percent ${MemberCount} ${TotalMembershipRows}; ${Coverage}=ConvertTo-Percent ${ActiveUserCount} ${EnabledUserCount}; ${OwnerCount}=if(${OwnerCountByGroupId}.ContainsKey(${Gid})){${OwnerCountByGroupId}[${Gid}]}else{0}; if(${IsEmpty}.IsPresent -and ${MemberCount} -ne 0){continue}; if(${MinGroupMembersCount} -gt 0 -and ${MemberCount} -lt ${MinGroupMembersCount}){continue}
        ${GroupRows}.Add([pscustomobject]@{GroupId=${Gid};GroupName=(Get-P ${G} "displayName");GroupMail=(Get-P ${G} "mail");GroupCategory=(Get-GroupCategory ${G});SecurityEnabled=(Get-P ${G} "securityEnabled");MailEnabled=(Get-P ${G} "mailEnabled");GroupTypes=(if((Get-P ${G} "groupTypes")){@(Get-P ${G} "groupTypes") -join ";"}else{""});IsDynamicGroup=(-not [string]::IsNullOrWhiteSpace([string](Get-P ${G} "membershipRule")));MembershipRule=(Get-P ${G} "membershipRule");MembershipRuleState=(Get-P ${G} "membershipRuleProcessingState");IsAssignableToRole=(Get-P ${G} "isAssignableToRole");OwnerCount=${OwnerCount};MemberCount=${MemberCount};UserMemberCount=${UniqueUserCount};ActiveUserMemberCount=${ActiveUserCount};DepartmentCount=@(${Depts}).Count;BlankDepartmentMemberCount=@(${UserMembers}|Where-Object{[string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment)}).Count;PrimaryDepartment=${PrimaryDept};PrimaryDepartmentCount=${PrimaryDeptCount};PrimaryDepartmentPctOfGroup=(ConvertTo-Percent ${PrimaryDeptCount} @(${UserMembers}).Count);MembershipDensityPct=${Density};UserCoveragePct=${Coverage};IsCrossDepartmentGroup=(@(${Depts}).Count -gt 1);IsHighDensityGroup=(${Density} -ge ${HighDensityPctThreshold});CreatedDateTime=(Get-P ${G} "createdDateTime");Visibility=(Get-P ${G} "visibility")})
    }

    ${GroupRowsForOutput}=@(${GroupRows}|Sort-Object -Property @{Expression="MembershipDensityPct";Descending=$true},@{Expression="GroupName";Ascending=$true})
    ${AllowedGroupIds}=@{}; foreach(${GR} in ${GroupRowsForOutput}){${AllowedGroupIds}[${GR}.GroupId]=$true}
    ${MemberRowsForOutput}=if(${IsEmpty}.IsPresent){@()}else{@(${MemberRows}|Where-Object{${AllowedGroupIds}.ContainsKey(${_}.GroupId)})}

    ${DepartmentUserTotals}=@{}; foreach(${DG} in (@(${Users})|Group-Object Department)){${DN}=if([string]::IsNullOrWhiteSpace([string]${DG}.Name)){"(blank)"}else{${DG}.Name};${DepartmentUserTotals}[${DN}]=${DG}.Count}
    ${DepartmentGroupRows}=New-Object System.Collections.Generic.List[object]
    foreach(${GD} in (@(${MemberRowsForOutput}|Where-Object{${_}.MemberType -eq "User"})|Group-Object GroupId,MemberDepartment)){
        ${Sample}=${GD}.Group[0]; ${DN}=if([string]::IsNullOrWhiteSpace([string]${Sample}.MemberDepartment)){"(blank)"}else{${Sample}.MemberDepartment}; ${Match}=@(${GroupRowsForOutput}|Where-Object{${_}.GroupId -eq ${Sample}.GroupId}|Select-Object -First 1); ${UsersInGroup}=@(${GD}.Group|Select-Object -ExpandProperty MemberId -Unique).Count; ${DeptTotal}=if(${DepartmentUserTotals}.ContainsKey(${DN})){${DepartmentUserTotals}[${DN}]}else{0}; ${GroupUserTotal}=if(${Match}.Count -gt 0){${Match}[0].UserMemberCount}else{0}; ${GroupDensity}=if(${Match}.Count -gt 0){${Match}[0].MembershipDensityPct}else{0}
        ${DepartmentGroupRows}.Add([pscustomobject]@{Department=${DN};GroupId=${Sample}.GroupId;GroupName=${Sample}.GroupName;GroupCategory=${Sample}.GroupCategory;UsersInDepartmentInGroup=${UsersInGroup};DepartmentTotalUsers=${DeptTotal};DepartmentCoveragePctForGroup=(ConvertTo-Percent ${UsersInGroup} ${DeptTotal});GroupShareFromDepartmentPct=(ConvertTo-Percent ${UsersInGroup} ${GroupUserTotal});GroupMembershipDensityPct=${GroupDensity};IsRoleAssignableGroup=${Sample}.GroupIsAssignableToRole;IsDynamicGroup=(-not [string]::IsNullOrWhiteSpace([string]${Sample}.GroupMembershipRule))})
    }
    ${DepartmentGroupRowsForOutput}=@(${DepartmentGroupRows}|Sort-Object -Property @{Expression="GroupMembershipDensityPct";Descending=$true},@{Expression="GroupName";Ascending=$true},@{Expression="Department";Ascending=$true})

    ${ExceptionRows}=New-Object System.Collections.Generic.List[object]
    foreach(${GR} in ${GroupRowsForOutput}){
        if(-not ${SkipOwners}.IsPresent -and ${GR}.OwnerCount -eq 0){${ExceptionRows}.Add([pscustomobject]@{Severity="Medium";Finding="Group has no owner";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="OwnerCount=0"})}
        if(${GR}.MemberCount -eq 0){${ExceptionRows}.Add([pscustomobject]@{Severity="Low";Finding="Empty group";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="MemberCount=0"})}
        if(${GR}.IsCrossDepartmentGroup){${ExceptionRows}.Add([pscustomobject]@{Severity="Review";Finding="Cross-department group";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="DepartmentCount=$(${GR}.DepartmentCount); PrimaryDepartment=$(${GR}.PrimaryDepartment); PrimaryDepartmentPctOfGroup=$(${GR}.PrimaryDepartmentPctOfGroup)"})}
        if(${GR}.BlankDepartmentMemberCount -gt 0){${ExceptionRows}.Add([pscustomobject]@{Severity="Review";Finding="Members missing department";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="BlankDepartmentMemberCount=$(${GR}.BlankDepartmentMemberCount)"})}
        if(${GR}.IsHighDensityGroup){${ExceptionRows}.Add([pscustomobject]@{Severity="Review";Finding="High group density";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="MembershipDensityPct=$(${GR}.MembershipDensityPct); Threshold=${HighDensityPctThreshold}"})}
        if(${GR}.IsAssignableToRole -eq $true){${ExceptionRows}.Add([pscustomobject]@{Severity="High";Finding="Role-assignable group";GroupId=${GR}.GroupId;GroupName=${GR}.GroupName;Detail="IsAssignableToRole=True"})}
    }

    Write-Stage "Writing CSV evidence"
    ${Users}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-Users.csv") -NoTypeInformation
    ${GroupRowsForOutput}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-Groups.csv") -NoTypeInformation
    ${MemberRowsForOutput}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-GroupMembers.csv") -NoTypeInformation
    ${OwnerRows}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-GroupOwners.csv") -NoTypeInformation
    ${DepartmentGroupRowsForOutput}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-DepartmentGroupMatrix.csv") -NoTypeInformation
    ${GroupRowsForOutput}|Select-Object GroupName,GroupId,GroupCategory,MemberCount,UserMemberCount,ActiveUserMemberCount,MembershipDensityPct,UserCoveragePct,DepartmentCount,PrimaryDepartment,PrimaryDepartmentPctOfGroup,OwnerCount,IsDynamicGroup,IsAssignableToRole,IsHighDensityGroup|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-GroupDensity.csv") -NoTypeInformation
    ${ExceptionRows}|Export-Csv (Join-Path ${OutputDir} "IdentityAudit-Exceptions.csv") -NoTypeInformation

    ${TopDensity}=@(${GroupRowsForOutput}|Sort-Object -Property @{Expression="MembershipDensityPct";Descending=$true},@{Expression="GroupName";Ascending=$true}|Select-Object -First 25); ${TopCoverage}=@(${GroupRowsForOutput}|Sort-Object -Property @{Expression="UserCoveragePct";Descending=$true},@{Expression="GroupName";Ascending=$true}|Select-Object -First 25); ${CrossDept}=@(${GroupRowsForOutput}|Where-Object{$_.IsCrossDepartmentGroup}|Sort-Object -Property @{Expression="DepartmentCount";Descending=$true},@{Expression="MembershipDensityPct";Descending=$true}|Select-Object -First 25); ${Ownerless}=@(${GroupRowsForOutput}|Where-Object{$_.OwnerCount -eq 0}|Select-Object -First 25); ${Dynamic}=@(${GroupRowsForOutput}|Where-Object{$_.IsDynamicGroup}|Select-Object -First 25); ${Hotspots}=@(${DepartmentGroupRowsForOutput}|Select-Object -First 25)
    ${Css}="<style>body{margin:0;padding:32px;background:#f5f7fb;color:#172033;font-family:Segoe UI,Arial,sans-serif}.meta,.muted,footer{color:#62708a;font-size:13px}.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin:18px 0 24px}.card,.panel{background:#fff;border:1px solid #d9e0ec;border-radius:14px;box-shadow:0 3px 10px rgba(23,32,51,.04)}.card{padding:16px}.label{color:#62708a;font-size:12px;text-transform:uppercase;letter-spacing:.05em}.value{font-size:28px;font-weight:700;margin-top:8px}.panel{padding:18px;margin:16px 0;overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px}th{text-align:left;color:#34415a;border-bottom:1px solid #d9e0ec;padding:8px;white-space:nowrap}td{border-bottom:1px solid #edf1f7;padding:8px;vertical-align:top}tr:hover td{background:#f9fbff}</style>"
    ${DashboardPath}=Join-Path ${OutputDir} "IdentityAudit-Dashboard.html"
    ${Html}=@"
<!doctype html><html><head><meta charset="utf-8"><title>Identity Audit Dashboard</title>${Css}</head><body>
<header><h1>Identity Audit Dashboard</h1><div class="meta">Run ID: $(HtmlEncode ${RunId}) | Mode: $(HtmlEncode ${ExecutionMode}) | Membership: $(if(${IncludeTransitiveMembership}.IsPresent){"Transitive"}else{"Direct"}) | Cache: $(HtmlEncode ${CacheRoot})</div></header>
<div class="cards"><div class="card"><div class="label">Groups</div><div class="value">$(@(${GroupRowsForOutput}).Count)</div></div><div class="card"><div class="label">Users</div><div class="value">${TotalUserCount}</div></div><div class="card"><div class="label">Membership rows</div><div class="value">${TotalMembershipRows}</div></div><div class="card"><div class="label">Ownerless groups</div><div class="value">$(@(${GroupRowsForOutput}|Where-Object{$_.OwnerCount -eq 0}).Count)</div></div><div class="card"><div class="label">Cross-dept groups</div><div class="value">$(@(${GroupRowsForOutput}|Where-Object{$_.IsCrossDepartmentGroup}).Count)</div></div><div class="card"><div class="label">High-density groups</div><div class="value">$(@(${GroupRowsForOutput}|Where-Object{$_.IsHighDensityGroup}).Count)</div></div><div class="card"><div class="label">Dynamic groups</div><div class="value">$(@(${GroupRowsForOutput}|Where-Object{$_.IsDynamicGroup}).Count)</div></div><div class="card"><div class="label">Users missing dept</div><div class="value">${NoDepartmentUserCount}</div></div></div>
$(HtmlTable "Top group density by membership percentage" ${TopDensity} @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","UserCoveragePct","DepartmentCount","PrimaryDepartment","PrimaryDepartmentPctOfGroup","OwnerCount","IsDynamicGroup","IsAssignableToRole"))
$(HtmlTable "Top groups by enabled user coverage percentage" ${TopCoverage} @("GroupName","GroupCategory","ActiveUserMemberCount","UserCoveragePct","MembershipDensityPct","DepartmentCount","PrimaryDepartment","OwnerCount"))
$(HtmlTable "Department / group hotspots" ${Hotspots} @("Department","GroupName","GroupCategory","UsersInDepartmentInGroup","DepartmentTotalUsers","DepartmentCoveragePctForGroup","GroupShareFromDepartmentPct","GroupMembershipDensityPct","IsRoleAssignableGroup","IsDynamicGroup"))
$(HtmlTable "Cross-department groups" ${CrossDept} @("GroupName","GroupCategory","MemberCount","UserMemberCount","DepartmentCount","PrimaryDepartment","PrimaryDepartmentPctOfGroup","MembershipDensityPct","OwnerCount"))
$(HtmlTable "Ownerless groups" ${Ownerless} @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","DepartmentCount","IsDynamicGroup","IsAssignableToRole"))
$(HtmlTable "Dynamic groups" ${Dynamic} @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","UserCoveragePct","DepartmentCount","PrimaryDepartment","MembershipRule","MembershipRuleState"))
<footer>Group density percentage = group membership rows divided by all membership rows observed in this run. User coverage percentage = enabled user members divided by all enabled users observed in this run.</footer></body></html>
"@
    ${Html}|Out-File ${DashboardPath} -Encoding utf8
    @("# Identity Audit Evidence Manifest","","Run ID: ${RunId}","Execution mode: ${ExecutionMode}","Cache root: ${CacheRoot}","Membership mode: $(if(${IncludeTransitiveMembership}.IsPresent){"Transitive"}else{"Direct"})","","Read-only collection. No users, groups, roles, policies, or memberships are modified.")|Out-File (Join-Path ${OutputDir} "IdentityAudit-Manifest.md") -Encoding utf8
    Write-Stage "Complete"; Write-Host "Output folder: ${OutputDir}"; Write-Host "Dashboard: ${DashboardPath}"; if(${OpenDashboard}.IsPresent){Invoke-Item ${DashboardPath}}
}
finally{ try{Disconnect-MgGraph|Out-Null}catch{} }
