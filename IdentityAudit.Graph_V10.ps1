<#
IdentityAudit.Graph_V10.ps1
Self-contained Microsoft Graph identity audit with cache, CSV evidence, dashboard, and graph analytics.
#>
[CmdletBinding()]
param(
    [string] ${TenantId},
    [string] ${ClientId},
    [string] ${CertificateThumbprint},
    [string] ${OutputRoot} = '.\IdentityAudit-Evidence',
    [string] ${CacheRoot} = '.\IdentityAudit-Cache',
    [int] ${CacheMaxAgeHours} = 168,
    [string] ${GroupIdsFile},
    [switch] ${IncludeTransitiveMembership},
    [switch] ${SecurityOnly},
    [switch] ${MailEnabledSecurityOnly},
    [switch] ${DistributionListOnly},
    [switch] ${Microsoft365Only},
    [switch] ${IsEmpty},
    [int] ${MinGroupMembersCount} = 0,
    [decimal] ${HighDensityPctThreshold} = 5.0,
    [switch] ${SkipOwners},
    [switch] ${InstallModules},
    [switch] ${OpenDashboard},
    [switch] ${UseCacheOnly},
    [switch] ${RefreshAll},
    [switch] ${RefreshUsers},
    [switch] ${RefreshGroups},
    [switch] ${RefreshMemberships},
    [switch] ${RefreshOwners},
    [int] ${MaxPathDepth} = 6,
    [string] ${HighValueGroupPattern} = '(?i)(admin|privileged|break.?glass|global administrator|role|security administrator|application administrator|owner)'
)

$ErrorActionPreference = 'Stop'

function Write-Stage([string] ${Message}) { Write-Host "[IdentityAudit] ${Message}" -ForegroundColor Cyan }
function Write-Warn([string] ${Message}) { Write-Host "[IdentityAudit][WARN] ${Message}" -ForegroundColor Yellow }
function Test-Value($Value) { return -not [string]::IsNullOrWhiteSpace([string] ${Value}) }
function HtmlSafe($Value) { if ($null -eq ${Value}) { return '' }; return [System.Net.WebUtility]::HtmlEncode([string] ${Value}) }
function Get-Prop($Object, [string] ${Name}) {
    if ($null -eq ${Object}) { return $null }
    ${p} = ${Object}.PSObject.Properties[${Name}]
    if ($null -ne ${p}) { return ${p}.Value }
    try { return ${Object}[${Name}] } catch { return $null }
}
function Set-MapValue($Map, $Key, $Value) { if (Test-Value ${Key}) { ${Map}[[string] ${Key}] = ${Value} } }
function Test-MapKey($Map, $Key) { if (-not (Test-Value ${Key})) { return $false }; return ${Map}.ContainsKey([string] ${Key}) }
function Get-MapValue($Map, $Key) { if (Test-MapKey ${Map} ${Key}) { return ${Map}[[string] ${Key}] }; return $null }

function Ensure-Module([string] ${Name}) {
    if (-not (Get-Module -ListAvailable -Name ${Name})) {
        if (${InstallModules}) { Install-Module ${Name} -Scope CurrentUser -Force -AllowClobber }
        else { throw "Missing module ${Name}. Use -InstallModules." }
    }
    Import-Module ${Name} -ErrorAction Stop
}

function Invoke-GraphGetJson([string] ${Uri}) {
    try {
        ${response} = Invoke-MgGraphRequest -Method GET -Uri ${Uri} -OutputType HttpResponseMessage
        ${content} = ${response}.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if ([string]::IsNullOrWhiteSpace(${content})) { return $null }
        return Microsoft.PowerShell.Utility\ConvertFrom-Json -InputObject ${content}
    } catch {
        throw "Graph GET failed. Uri='${Uri}'. Error='$($_.Exception.Message)'"
    }
}

function Invoke-GraphPagedJson([string] ${Uri}) {
    ${rows} = New-Object System.Collections.Generic.List[object]
    ${next} = ${Uri}
    while (-not [string]::IsNullOrWhiteSpace([string] ${next})) {
        ${page} = Invoke-GraphGetJson ${next}
        foreach (${item} in @((Get-Prop ${page} 'value'))) {
            if ($null -ne ${item}) { [void] ${rows}.Add(${item}) }
        }
        ${next} = Get-Prop ${page} '@odata.nextLink'
    }
    return ${rows}.ToArray()
}

function Save-JsonCache([string] ${Path}, $Rows) {
    ${items} = @(${Rows} | Where-Object { $null -ne $_ })
    if (${items}.Count -eq 0) { '[]' | Out-File ${Path} -Encoding utf8 -Force }
    else { ${items} | ConvertTo-Json -Depth 100 | Out-File ${Path} -Encoding utf8 -Force }
}
function Read-JsonCache([string] ${Path}) {
    if (-not (Test-Path ${Path})) { return @() }
    ${raw} = Get-Content ${Path} -Raw
    if ([string]::IsNullOrWhiteSpace(${raw})) { return @() }
    return @(${raw} | Microsoft.PowerShell.Utility\ConvertFrom-Json | Where-Object { $null -ne $_ })
}
function Test-CacheFresh([string] ${Path}) {
    if (-not (Test-Path ${Path})) { return $false }
    if (${CacheMaxAgeHours} -le 0) { return $true }
    return (((Get-Date) - (Get-Item ${Path}).LastWriteTime).TotalHours -le ${CacheMaxAgeHours})
}
function Export-CsvSafe($Rows, [string] ${Path}) {
    ${items} = @(${Rows} | Where-Object { $null -ne $_ })
    if (${items}.Count -eq 0) { New-Item -ItemType File -Path ${Path} -Force | Out-Null }
    else { ${items} | Export-Csv ${Path} -NoTypeInformation }
}

function Get-ObjectType($Object) {
    ${type} = [string] (Get-Prop ${Object} '@odata.type')
    if (${type} -match 'user$') { return 'User' }
    if (${type} -match 'group$') { return 'Group' }
    if (${type} -match 'servicePrincipal$') { return 'ServicePrincipal' }
    if (${type} -match 'device$') { return 'Device' }
    if (${type} -match 'orgContact$') { return 'Contact' }
    return 'DirectoryObject'
}
function Get-Display($Object) {
    ${v} = Get-Prop ${Object} 'displayName'
    if (-not (Test-Value ${v})) { ${v} = Get-Prop ${Object} 'userPrincipalName' }
    if (-not (Test-Value ${v})) { ${v} = Get-Prop ${Object} 'mail' }
    return ${v}
}
function Get-GroupTypesText($Group) {
    ${types} = Get-Prop ${Group} 'groupTypes'
    if ($null -eq ${types}) { return '' }
    return @(${types}) -join ';'
}
function Get-GroupCategory($Group) {
    ${types} = Get-GroupTypesText ${Group}
    ${securityEnabled} = Get-Prop ${Group} 'securityEnabled'
    ${mailEnabled} = Get-Prop ${Group} 'mailEnabled'
    if (${types} -match 'Unified') { return 'Microsoft365' }
    if (${securityEnabled} -eq $true -and ${mailEnabled} -eq $true) { return 'MailEnabledSecurity' }
    if (${securityEnabled} -eq $true) { return 'Security' }
    if (${mailEnabled} -eq $true) { return 'DistributionList' }
    return 'Other'
}
function Convert-Percent([decimal] ${Numerator}, [decimal] ${Denominator}) {
    if (${Denominator} -le 0) { return [decimal] 0 }
    return [math]::Round((${Numerator} / ${Denominator}) * 100, 2)
}
function New-HtmlTable([string] ${Title}, $Rows, [string[]] ${Columns}, [int] ${MaxRows} = 25) {
    ${items} = @(${Rows} | Where-Object { $null -ne $_ })
    ${html} = "<section class='panel'><h2>$(HtmlSafe ${Title})</h2>"
    if (${items}.Count -eq 0) { return ${html} + '<p class="muted">No records found.</p></section>' }
    ${html} += '<table><thead><tr>'
    foreach (${column} in ${Columns}) { ${html} += "<th>$(HtmlSafe ${column})</th>" }
    ${html} += '</tr></thead><tbody>'
    foreach (${row} in @(${items} | Select-Object -First ${MaxRows})) {
        ${html} += '<tr>'
        foreach (${column} in ${Columns}) {
            ${value} = ''
            if (${row}.PSObject.Properties.Name -contains ${column}) { ${value} = ${row}.${column} }
            ${html} += "<td>$(HtmlSafe ${value})</td>"
        }
        ${html} += '</tr>'
    }
    return ${html} + '</tbody></table></section>'
}
function Add-Node($Map, $Id, $Label, $Type, $SubType = '') {
    if (-not (Test-Value ${Id})) { return }
    if (-not (Test-MapKey ${Map} ${Id})) {
        ${Map}[[string] ${Id}] = [pscustomobject]@{ Id = ${Id}; Label = ${Label}; Type = ${Type}; SubType = ${SubType}; RiskScore = 0; RiskDrivers = '' }
    }
}
function New-Edge($SourceId, $TargetId, $EdgeType, $SourceLabel, $TargetLabel, $SourceType, $TargetType) {
    if (-not (Test-Value ${SourceId}) -or -not (Test-Value ${TargetId})) { return $null }
    return [pscustomobject]@{ SourceId = ${SourceId}; SourceLabel = ${SourceLabel}; SourceType = ${SourceType}; TargetId = ${TargetId}; TargetLabel = ${TargetLabel}; TargetType = ${TargetType}; EdgeType = ${EdgeType} }
}
function Find-ShortestPath($Adjacency, [string] ${Start}, $Targets, [int] ${DepthLimit}) {
    if (-not (Test-Value ${Start})) { return @() }
    ${queue} = New-Object System.Collections.Queue
    ${seen} = @{}
    ${queue}.Enqueue(@(${Start}, @(${Start})))
    Set-MapValue ${seen} ${Start} $true
    while (${queue}.Count -gt 0) {
        ${item} = ${queue}.Dequeue()
        ${node} = ${item}[0]
        ${path} = @(${item}[1])
        if (-not (Test-Value ${node})) { continue }
        if ((Test-MapKey ${Targets} ${node}) -and ${node} -ne ${Start}) { return ${path} }
        if (${path}.Count -gt (${DepthLimit} + 1)) { continue }
        foreach (${next} in @(${Adjacency}[[string] ${node}])) {
            if ((Test-Value ${next}) -and -not (Test-MapKey ${seen} ${next})) {
                Set-MapValue ${seen} ${next} $true
                ${queue}.Enqueue(@(${next}, @(${path} + ${next})))
            }
        }
    }
    return @()
}
function Find-Cycles($Adjacency, [int] ${DepthLimit}) {
    ${found} = @{}
    ${cycles} = @()
    foreach (${start} in ${Adjacency}.Keys) {
        if (-not (Test-Value ${start})) { continue }
        ${stack} = @(@(${start}, @(${start})))
        while (${stack}.Count -gt 0) {
            ${current} = ${stack}[-1]
            if (${stack}.Count -eq 1) { ${stack} = @() } else { ${stack} = ${stack}[0..(${stack}.Count - 2)] }
            ${node} = ${current}[0]
            ${path} = @(${current}[1])
            if (${path}.Count -gt (${DepthLimit} + 1)) { continue }
            foreach (${next} in @(${Adjacency}[[string] ${node}])) {
                if (-not (Test-Value ${next})) { continue }
                if (${next} -eq ${start} -and ${path}.Count -gt 1) {
                    ${cycle} = @(${path} + ${start})
                    ${key} = (@(${cycle} | Sort-Object) -join '|')
                    if (-not (Test-MapKey ${found} ${key})) {
                        Set-MapValue ${found} ${key} $true
                        ${cycles} += [pscustomobject]@{ Length = (${cycle}.Count - 1); Cycle = (${cycle} -join ' -> ') }
                    }
                } elseif (${path} -notcontains ${next}) {
                    ${stack} += ,@(${next}, @(${path} + ${next}))
                }
            }
        }
    }
    return ${cycles}
}

New-Item -ItemType Directory -Path ${CacheRoot} -Force | Out-Null
${usersCache} = Join-Path ${CacheRoot} 'users.json'
${groupsCache} = Join-Path ${CacheRoot} 'groups.json'
${membersCache} = Join-Path ${CacheRoot} 'members.direct.json'
${membershipMode} = 'Direct'
if (${IncludeTransitiveMembership}) { ${membersCache} = Join-Path ${CacheRoot} 'members.transitive.json'; ${membershipMode} = 'Transitive' }
${ownersCache} = Join-Path ${CacheRoot} 'owners.json'

${doUsers} = [bool] ${RefreshUsers}
${doGroups} = [bool] ${RefreshGroups}
${doMembers} = [bool] ${RefreshMemberships}
${doOwners} = [bool] ${RefreshOwners}
if (${RefreshAll}) { ${doUsers} = $true; ${doGroups} = $true; ${doMembers} = $true; ${doOwners} = $true }
if (-not (Test-CacheFresh ${usersCache})) { ${doUsers} = $true }
if (-not (Test-CacheFresh ${groupsCache})) { ${doGroups} = $true }
if (-not (Test-CacheFresh ${membersCache})) { ${doMembers} = $true }
if (-not ${SkipOwners} -and -not (Test-CacheFresh ${ownersCache})) { ${doOwners} = $true }
if (${doGroups}) { ${doMembers} = $true; if (-not ${SkipOwners}) { ${doOwners} = $true } }
if (${UseCacheOnly}) {
    foreach (${path} in @(${usersCache}, ${groupsCache}, ${membersCache})) { if (-not (Test-Path ${path})) { throw "Cache-only requested but missing cache file: ${path}" } }
    if (-not ${SkipOwners} -and -not (Test-Path ${ownersCache})) { throw "Cache-only requested but missing cache file: ${ownersCache}" }
    ${doUsers} = $false; ${doGroups} = $false; ${doMembers} = $false; ${doOwners} = $false
}

${needsGraph} = (${doUsers} -or ${doGroups} -or ${doMembers} -or ${doOwners})
${connected} = $false
try {
    if (${needsGraph}) {
        Write-Stage 'Preparing modules'
        Ensure-Module 'Microsoft.Graph.Authentication'
        Write-Stage 'Connecting to Microsoft Graph'
        if (${TenantId} -and ${ClientId} -and ${CertificateThumbprint}) { Connect-MgGraph -TenantId ${TenantId} -ClientId ${ClientId} -CertificateThumbprint ${CertificateThumbprint} -NoWelcome }
        else { Connect-MgGraph -Scopes @('User.Read.All','Group.Read.All','GroupMember.Read.All','Directory.Read.All') -NoWelcome }
        ${connected} = $true
        if (-not ${TenantId}) { ${TenantId} = (Get-MgContext).TenantId }
    } else { Write-Stage 'Using cache only; Graph connection skipped' }

    if (${doUsers}) {
        Write-Stage 'Refreshing users cache'
        Save-JsonCache ${usersCache} (Invoke-GraphPagedJson '/v1.0/users?$select=id,displayName,userPrincipalName,mail,department,jobTitle,companyName,accountEnabled,userType,employeeId,createdDateTime&$top=999')
    } else { Write-Stage 'Using users cache' }

    if (${doGroups}) {
        Write-Stage 'Refreshing groups cache'
        if (${GroupIdsFile}) {
            ${rawGroups} = @()
            foreach (${gidRaw} in (Get-Content ${GroupIdsFile} | Where-Object { $_ })) {
                ${gid} = ${gidRaw}.Trim()
                if (-not (Test-Value ${gid})) { continue }
                try { ${rawGroups} += Invoke-GraphGetJson "/v1.0/groups/${gid}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility" }
                catch { ${rawGroups} += Invoke-GraphGetJson "/v1.0/groups/${gid}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime" }
            }
            Save-JsonCache ${groupsCache} ${rawGroups}
        } else {
            try { Save-JsonCache ${groupsCache} (Invoke-GraphPagedJson '/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility&$top=999') }
            catch { Write-Warn "Advanced group properties failed; using baseline. $($_.Exception.Message)"; Save-JsonCache ${groupsCache} (Invoke-GraphPagedJson '/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime&$top=999') }
        }
    } else { Write-Stage 'Using groups cache' }

    ${rawGroupsForCollection} = Read-JsonCache ${groupsCache}
    if (${doMembers}) {
        Write-Stage 'Refreshing membership cache'
        ${memberCacheRows} = @()
        ${endpoint} = 'members'
        if (${IncludeTransitiveMembership}) { ${endpoint} = 'transitiveMembers' }
        ${i} = 0
        foreach (${group} in @(${rawGroupsForCollection})) {
            ${i}++
            ${gid} = Get-Prop ${group} 'id'
            ${gname} = Get-Prop ${group} 'displayName'
            if (-not (Test-Value ${gid})) { continue }
            Write-Progress -Activity 'Refreshing memberships' -Status ${gname} -PercentComplete ((${i} / [math]::Max(1, @(${rawGroupsForCollection}).Count)) * 100)
            try { ${members} = Invoke-GraphPagedJson "/v1.0/groups/${gid}/${endpoint}?`$select=id,displayName,userPrincipalName,mail&`$top=999" }
            catch { Write-Warn "Members failed for ${gname}. $($_.Exception.Message)"; ${members} = @() }
            foreach (${member} in @(${members})) {
                ${mid} = Get-Prop ${member} 'id'
                if (-not (Test-Value ${mid})) { continue }
                ${memberCacheRows} += [pscustomobject]@{ GroupId=${gid}; GroupName=${gname}; GroupMail=Get-Prop ${group} 'mail'; GroupCategory=Get-GroupCategory ${group}; GroupSecurityEnabled=Get-Prop ${group} 'securityEnabled'; GroupMailEnabled=Get-Prop ${group} 'mailEnabled'; GroupTypes=Get-GroupTypesText ${group}; GroupIsAssignableToRole=Get-Prop ${group} 'isAssignableToRole'; GroupMembershipRule=Get-Prop ${group} 'membershipRule'; MembershipRuleState=Get-Prop ${group} 'membershipRuleProcessingState'; MembershipMode=${membershipMode}; MemberId=${mid}; MemberDisplayName=Get-Display ${member}; MemberUPN=Get-Prop ${member} 'userPrincipalName'; MemberMail=Get-Prop ${member} 'mail'; MemberType=Get-ObjectType ${member} }
            }
        }
        Write-Progress -Activity 'Refreshing memberships' -Completed
        Save-JsonCache ${membersCache} ${memberCacheRows}
    } else { Write-Stage 'Using membership cache' }

    if (-not ${SkipOwners}) {
        if (${doOwners}) {
            Write-Stage 'Refreshing owners cache'
            ${ownerCacheRows} = @()
            foreach (${group} in @(${rawGroupsForCollection})) {
                ${gid} = Get-Prop ${group} 'id'
                ${gname} = Get-Prop ${group} 'displayName'
                if (-not (Test-Value ${gid})) { continue }
                try { ${owners} = Invoke-GraphPagedJson "/v1.0/groups/${gid}/owners?`$select=id,displayName,userPrincipalName,mail&`$top=999" }
                catch { Write-Warn "Owners failed for ${gname}. $($_.Exception.Message)"; ${owners} = @() }
                foreach (${owner} in @(${owners})) {
                    ${oid} = Get-Prop ${owner} 'id'
                    if (-not (Test-Value ${oid})) { continue }
                    ${ownerCacheRows} += [pscustomobject]@{ GroupId=${gid}; GroupName=${gname}; GroupCategory=Get-GroupCategory ${group}; OwnerId=${oid}; OwnerDisplayName=Get-Display ${owner}; OwnerUPN=Get-Prop ${owner} 'userPrincipalName'; OwnerType=Get-ObjectType ${owner} }
                }
            }
            Save-JsonCache ${ownersCache} ${ownerCacheRows}
        } else { Write-Stage 'Using owners cache' }
    }
} finally {
    if (${connected}) { try { Disconnect-MgGraph | Out-Null } catch {} }
}

Write-Stage 'Building report from cache'
${rawUsers} = Read-JsonCache ${usersCache}
${rawGroups} = Read-JsonCache ${groupsCache}
${memberCacheRows} = Read-JsonCache ${membersCache}
${ownerRows} = @(); if (-not ${SkipOwners}) { ${ownerRows} = Read-JsonCache ${ownersCache} }

${users} = @(); ${usersById} = @{}
foreach (${user} in @(${rawUsers})) {
    ${uid} = Get-Prop ${user} 'id'
    ${row} = [pscustomobject]@{ Id=${uid}; DisplayName=Get-Prop ${user} 'displayName'; UserPrincipalName=Get-Prop ${user} 'userPrincipalName'; Mail=Get-Prop ${user} 'mail'; Department=Get-Prop ${user} 'department'; JobTitle=Get-Prop ${user} 'jobTitle'; CompanyName=Get-Prop ${user} 'companyName'; AccountEnabled=Get-Prop ${user} 'accountEnabled'; UserType=Get-Prop ${user} 'userType'; EmployeeId=Get-Prop ${user} 'employeeId'; CreatedDateTime=Get-Prop ${user} 'createdDateTime' }
    ${users} += ${row}; Set-MapValue ${usersById} ${uid} ${row}
}

${groups} = @()
foreach (${group} in @(${rawGroups})) {
    ${gid} = Get-Prop ${group} 'id'; if (-not (Test-Value ${gid})) { continue }
    ${cat} = Get-GroupCategory ${group}
    if (${SecurityOnly} -and ${cat} -ne 'Security') { continue }
    if (${MailEnabledSecurityOnly} -and ${cat} -ne 'MailEnabledSecurity') { continue }
    if (${DistributionListOnly} -and ${cat} -ne 'DistributionList') { continue }
    if (${Microsoft365Only} -and ${cat} -ne 'Microsoft365') { continue }
    ${groups} += ${group}
}

${filteredIds} = @{}
foreach (${group} in @(${groups})) { Set-MapValue ${filteredIds} (Get-Prop ${group} 'id') $true }

${memberRows} = @()
foreach (${member} in @(${memberCacheRows})) {
    if (-not (Test-MapKey ${filteredIds} ${member}.GroupId)) { continue }
    ${profile} = Get-MapValue ${usersById} ${member}.MemberId
    ${dept}=''; ${job}=''; ${company}=''; ${enabled}=$null; ${userType}=''; ${employeeId}=''; ${upn}=${member}.MemberUPN; ${mail}=${member}.MemberMail
    if ($null -ne ${profile}) {
        ${dept}=${profile}.Department; ${job}=${profile}.JobTitle; ${company}=${profile}.CompanyName; ${enabled}=${profile}.AccountEnabled; ${userType}=${profile}.UserType; ${employeeId}=${profile}.EmployeeId
        if (-not (Test-Value ${upn})) { ${upn}=${profile}.UserPrincipalName }
        if (-not (Test-Value ${mail})) { ${mail}=${profile}.Mail }
    }
    ${memberRows} += [pscustomobject]@{ GroupId=${member}.GroupId; GroupName=${member}.GroupName; GroupMail=${member}.GroupMail; GroupCategory=${member}.GroupCategory; GroupSecurityEnabled=${member}.GroupSecurityEnabled; GroupMailEnabled=${member}.GroupMailEnabled; GroupTypes=${member}.GroupTypes; GroupIsAssignableToRole=${member}.GroupIsAssignableToRole; GroupMembershipRule=${member}.GroupMembershipRule; MembershipRuleState=${member}.MembershipRuleState; MembershipMode=${member}.MembershipMode; MemberId=${member}.MemberId; MemberDisplayName=${member}.MemberDisplayName; MemberUPN=${upn}; MemberMail=${mail}; MemberType=${member}.MemberType; MemberDepartment=${dept}; MemberJobTitle=${job}; MemberCompanyName=${company}; MemberAccountEnabled=${enabled}; MemberUserType=${userType}; MemberEmployeeId=${employeeId} }
}

${enabledUsers} = @(${users} | Where-Object { $_.AccountEnabled -eq $true }).Count
${totalMembershipRows} = @(${memberRows}).Count
${ownerCounts} = @{}
foreach (${ownerGroup} in @(${ownerRows} | Where-Object { Test-Value $_.GroupId } | Group-Object GroupId)) { Set-MapValue ${ownerCounts} ${ownerGroup}.Name ${ownerGroup}.Count }

${groupRows} = @()
foreach (${group} in @(${groups})) {
    ${gid}=Get-Prop ${group} 'id'; if (-not (Test-Value ${gid})) { continue }
    ${gname}=Get-Prop ${group} 'displayName'
    ${rows}=@(${memberRows} | Where-Object { $_.GroupId -eq ${gid} })
    ${userMembers}=@(${rows} | Where-Object { $_.MemberType -eq 'User' })
    ${depts}=@(${userMembers} | Where-Object { Test-Value $_.MemberDepartment } | Select-Object -ExpandProperty MemberDepartment -Unique)
    ${deptGroups}=@(${userMembers} | Where-Object { Test-Value $_.MemberDepartment } | Group-Object MemberDepartment | Sort-Object @{Expression='Count';Descending=$true})
    ${primaryDept}=''; ${primaryDeptCount}=0
    if (${deptGroups}.Count -gt 0) { ${primaryDept}=${deptGroups}[0].Name; ${primaryDeptCount}=${deptGroups}[0].Count }
    ${memberCount}=${rows}.Count
    ${uniqueUserCount}=@(${userMembers} | Select-Object -ExpandProperty MemberId -Unique).Count
    ${activeUserCount}=@(${userMembers} | Where-Object { $_.MemberAccountEnabled -eq $true } | Select-Object -ExpandProperty MemberId -Unique).Count
    ${ownerCount}=0; if (Test-MapKey ${ownerCounts} ${gid}) { ${ownerCount}=Get-MapValue ${ownerCounts} ${gid} }
    ${density}=Convert-Percent ${memberCount} ${totalMembershipRows}
    ${coverage}=Convert-Percent ${activeUserCount} ${enabledUsers}
    if (${IsEmpty} -and ${memberCount} -ne 0) { continue }
    if (${MinGroupMembersCount} -gt 0 -and ${memberCount} -lt ${MinGroupMembersCount}) { continue }
    ${rule}=Get-Prop ${group} 'membershipRule'
    ${isDynamic}=Test-Value ${rule}
    ${groupRows} += [pscustomobject]@{ GroupId=${gid}; GroupName=${gname}; GroupMail=Get-Prop ${group} 'mail'; GroupCategory=Get-GroupCategory ${group}; SecurityEnabled=Get-Prop ${group} 'securityEnabled'; MailEnabled=Get-Prop ${group} 'mailEnabled'; GroupTypes=Get-GroupTypesText ${group}; IsDynamicGroup=${isDynamic}; MembershipRule=${rule}; MembershipRuleState=Get-Prop ${group} 'membershipRuleProcessingState'; IsAssignableToRole=Get-Prop ${group} 'isAssignableToRole'; OwnerCount=${ownerCount}; MemberCount=${memberCount}; UserMemberCount=${uniqueUserCount}; ActiveUserMemberCount=${activeUserCount}; DepartmentCount=@(${depts}).Count; BlankDepartmentMemberCount=@(${userMembers} | Where-Object { -not (Test-Value $_.MemberDepartment) }).Count; PrimaryDepartment=${primaryDept}; PrimaryDepartmentCount=${primaryDeptCount}; PrimaryDepartmentPctOfGroup=Convert-Percent ${primaryDeptCount} @(${userMembers}).Count; MembershipDensityPct=${density}; UserCoveragePct=${coverage}; IsCrossDepartmentGroup=(@(${depts}).Count -gt 1); IsHighDensityGroup=(${density} -ge ${HighDensityPctThreshold}); CreatedDateTime=Get-Prop ${group} 'createdDateTime'; Visibility=Get-Prop ${group} 'visibility' }
}
${groupRows}=@(${groupRows} | Sort-Object @{Expression='MembershipDensityPct';Descending=$true},@{Expression='GroupName';Ascending=$true})
${allowedGroups}=@{}; foreach (${group} in @(${groupRows})) { Set-MapValue ${allowedGroups} ${group}.GroupId $true }
${memberOut}=@(${memberRows} | Where-Object { Test-MapKey ${allowedGroups} $_.GroupId })

${deptTotals}=@{}; foreach (${deptGroup} in @(${users} | Group-Object Department)) { ${dn}=${deptGroup}.Name; if (-not (Test-Value ${dn})) { ${dn}='(blank)' }; Set-MapValue ${deptTotals} ${dn} ${deptGroup}.Count }
${deptRows}=@()
foreach (${gd} in @(${memberOut} | Where-Object { $_.MemberType -eq 'User' } | Group-Object GroupId,MemberDepartment)) {
    ${sample}=${gd}.Group[0]
    ${dn}=${sample}.MemberDepartment; if (-not (Test-Value ${dn})) { ${dn}='(blank)' }
    ${count}=@(${gd}.Group | Select-Object -ExpandProperty MemberId -Unique).Count
    ${total}=0; if (Test-MapKey ${deptTotals} ${dn}) { ${total}=Get-MapValue ${deptTotals} ${dn} }
    ${gr}=@(${groupRows} | Where-Object { $_.GroupId -eq ${sample}.GroupId } | Select-Object -First 1)
    ${groupUserCount}=0; ${groupDensity}=0
    if (${gr}) { ${groupUserCount}=${gr}[0].UserMemberCount; ${groupDensity}=${gr}[0].MembershipDensityPct }
    ${deptRows}+=[pscustomobject]@{ Department=${dn}; GroupId=${sample}.GroupId; GroupName=${sample}.GroupName; GroupCategory=${sample}.GroupCategory; UsersInDepartmentInGroup=${count}; DepartmentTotalUsers=${total}; DepartmentCoveragePctForGroup=Convert-Percent ${count} ${total}; GroupShareFromDepartmentPct=Convert-Percent ${count} ${groupUserCount}; GroupMembershipDensityPct=${groupDensity}; IsRoleAssignableGroup=${sample}.GroupIsAssignableToRole; IsDynamicGroup=(Test-Value ${sample}.GroupMembershipRule) }
}
${deptRows}=@(${deptRows} | Sort-Object @{Expression='GroupMembershipDensityPct';Descending=$true},@{Expression='GroupName';Ascending=$true})

${exceptions}=@()
foreach (${group} in @(${groupRows})) {
    if (-not ${SkipOwners} -and [int]${group}.OwnerCount -eq 0) { ${exceptions}+=[pscustomobject]@{Severity='Medium';Finding='Group has no owner';GroupId=${group}.GroupId;GroupName=${group}.GroupName;Detail='OwnerCount=0'} }
    if (${group}.IsHighDensityGroup) { ${exceptions}+=[pscustomobject]@{Severity='Review';Finding='High group density';GroupId=${group}.GroupId;GroupName=${group}.GroupName;Detail="MembershipDensityPct=$($group.MembershipDensityPct)"} }
    if (${group}.IsCrossDepartmentGroup) { ${exceptions}+=[pscustomobject]@{Severity='Review';Finding='Cross-department group';GroupId=${group}.GroupId;GroupName=${group}.GroupName;Detail="DepartmentCount=$($group.DepartmentCount)"} }
    if (${group}.IsAssignableToRole) { ${exceptions}+=[pscustomobject]@{Severity='High';Finding='Role-assignable group';GroupId=${group}.GroupId;GroupName=${group}.GroupName;Detail='IsAssignableToRole=True'} }
}

${runId}=Get-Date -Format 'yyyyMMdd-HHmmss'
${out}=Join-Path ${OutputRoot} ${runId}
New-Item -ItemType Directory -Path ${out} -Force | Out-Null
Write-Stage 'Writing CSV evidence'
Export-CsvSafe ${users} (Join-Path ${out} 'IdentityAudit-Users.csv')
Export-CsvSafe ${groupRows} (Join-Path ${out} 'IdentityAudit-Groups.csv')
Export-CsvSafe ${memberOut} (Join-Path ${out} 'IdentityAudit-GroupMembers.csv')
Export-CsvSafe ${ownerRows} (Join-Path ${out} 'IdentityAudit-GroupOwners.csv')
Export-CsvSafe ${deptRows} (Join-Path ${out} 'IdentityAudit-DepartmentGroupMatrix.csv')
Export-CsvSafe (${groupRows} | Select-Object GroupName,GroupId,GroupCategory,MemberCount,UserMemberCount,ActiveUserMemberCount,MembershipDensityPct,UserCoveragePct,DepartmentCount,PrimaryDepartment,PrimaryDepartmentPctOfGroup,OwnerCount,IsDynamicGroup,IsAssignableToRole,IsHighDensityGroup) (Join-Path ${out} 'IdentityAudit-GroupDensity.csv')
Export-CsvSafe ${exceptions} (Join-Path ${out} 'IdentityAudit-Exceptions.csv')

Write-Stage 'Building graph analytics'
${nodes}=@{}; ${edges}=@(); ${groupById}=@{}
foreach (${user} in @(${users})) { Add-Node ${nodes} ${user}.Id ($(if(${user}.UserPrincipalName){${user}.UserPrincipalName}else{${user}.DisplayName})) 'User' ${user}.UserType }
foreach (${group} in @(${groupRows})) { Add-Node ${nodes} ${group}.GroupId ${group}.GroupName 'Group' ${group}.GroupCategory; Set-MapValue ${groupById} ${group}.GroupId ${group} }
foreach (${member} in @(${memberOut})) { Add-Node ${nodes} ${member}.MemberId ($(if(${member}.MemberUPN){${member}.MemberUPN}else{${member}.MemberDisplayName})) ${member}.MemberType ''; ${edge}=New-Edge ${member}.MemberId ${member}.GroupId 'MemberOf' ($(if(${member}.MemberUPN){${member}.MemberUPN}else{${member}.MemberDisplayName})) ${member}.GroupName ${member}.MemberType 'Group'; if ($null -ne ${edge}) { ${edges}+=${edge} } }
foreach (${owner} in @(${ownerRows})) { Add-Node ${nodes} ${owner}.OwnerId ($(if(${owner}.OwnerUPN){${owner}.OwnerUPN}else{${owner}.OwnerDisplayName})) ${owner}.OwnerType ''; ${edge}=New-Edge ${owner}.OwnerId ${owner}.GroupId 'OwnsGroup' ($(if(${owner}.OwnerUPN){${owner}.OwnerUPN}else{${owner}.OwnerDisplayName})) ${owner}.GroupName ${owner}.OwnerType 'Group'; if ($null -ne ${edge}) { ${edges}+=${edge} } }
${groupEdges}=@(${edges} | Where-Object { $_.SourceType -eq 'Group' -and $_.EdgeType -eq 'MemberOf' -and (Test-Value $_.SourceId) -and (Test-Value $_.TargetId) })
${adjacency}=@{}; foreach (${edge} in @(${groupEdges})) { if (-not (Test-MapKey ${adjacency} ${edge}.SourceId)) { ${adjacency}[[string]${edge}.SourceId]=@() }; ${adjacency}[[string]${edge}.SourceId]+=${edge}.TargetId }
${highTargets}=@{}; foreach (${group} in @(${groupRows})) { if ((Test-Value ${group}.GroupId) -and (${group}.IsAssignableToRole -eq $true -or ${group}.GroupName -match ${HighValueGroupPattern})) { Set-MapValue ${highTargets} ${group}.GroupId $true } }
${paths}=@(); foreach (${edge} in @(${edges} | Where-Object { ($_.SourceType -eq 'User' -or $_.SourceType -eq 'Group' -or $_.SourceType -eq 'ServicePrincipal') -and (Test-Value $_.TargetId) })) { if (Test-MapKey ${highTargets} ${edge}.TargetId) { continue }; ${path}=Find-ShortestPath ${adjacency} ${edge}.TargetId ${highTargets} ${MaxPathDepth}; if (@(${path}).Count -gt 0) { ${labels}=@(); foreach (${id} in @(${path})) { ${g}=Get-MapValue ${groupById} ${id}; if ($null -ne ${g}) { ${labels}+=${g}.GroupName } else { ${labels}+=${id} } }; ${paths}+=[pscustomobject]@{StartId=${edge}.SourceId;StartLabel=${edge}.SourceLabel;StartType=${edge}.SourceType;EntryGroup=${edge}.TargetLabel;TargetGroup=${labels}[-1];HopCount=@(${path}).Count;Path=(${edge}.SourceLabel+' -> '+(${labels} -join ' -> '));Risk='High'} } }
${cycles}=Find-Cycles ${adjacency} ${MaxPathDepth}
${nestedStats}=@(); foreach (${group} in @(${groupRows})) { ${gid}=${group}.GroupId; if (-not (Test-Value ${gid})) { continue }; ${nestedStats}+=[pscustomobject]@{GroupId=${gid};GroupName=${group}.GroupName;NestedGroupMemberCount=@(${groupEdges}|Where-Object{$_.TargetId -eq ${gid}}).Count;NestedIntoGroupCount=@(${groupEdges}|Where-Object{$_.SourceId -eq ${gid}}).Count;MemberCount=[int]${group}.MemberCount;DepartmentCount=[int]${group}.DepartmentCount} }
${riskRows}=@(); foreach (${group} in @(${groupRows})) { ${score}=0; ${drivers}=@(); if (${group}.IsAssignableToRole -eq $true) { ${score}+=50; ${drivers}+='Role-assignable' }; if ([int]${group}.OwnerCount -eq 0) { ${score}+=20; ${drivers}+='Ownerless' }; if (${group}.IsHighDensityGroup -eq $true) { ${score}+=20; ${drivers}+='High density' }; if (${group}.IsCrossDepartmentGroup -eq $true) { ${score}+=15; ${drivers}+='Cross-department' }; if (${group}.IsDynamicGroup -eq $true) { ${score}+=5; ${drivers}+='Dynamic' }; if ([int]${group}.DepartmentCount -gt 5) { ${score}+=10; ${drivers}+='Many departments' }; if ([int]${group}.MemberCount -gt 500) { ${score}+=10; ${drivers}+='Large membership' }; ${ns}=@(${nestedStats}|Where-Object{$_.GroupId -eq ${group}.GroupId}|Select-Object -First 1); if (${ns} -and (${ns}[0].NestedGroupMemberCount -gt 0 -or ${ns}[0].NestedIntoGroupCount -gt 0)) { ${score}+=10; ${drivers}+='Nested group' }; ${pc}=@(${paths}|Where-Object{$_.EntryGroup -eq ${group}.GroupName -or $_.TargetGroup -eq ${group}.GroupName}).Count; if (${pc} -gt 0) { ${score}+=25; ${drivers}+='Privileged path' }; ${riskRows}+=[pscustomobject]@{GroupId=${group}.GroupId;GroupName=${group}.GroupName;RiskScore=${score};RiskDrivers=(${drivers} -join '; ');MemberCount=${group}.MemberCount;OwnerCount=${group}.OwnerCount;DepartmentCount=${group}.DepartmentCount;MembershipDensityPct=${group}.MembershipDensityPct;UserCoveragePct=${group}.UserCoveragePct;PrivilegedPathCount=${pc};IsAssignableToRole=${group}.IsAssignableToRole;IsDynamicGroup=${group}.IsDynamicGroup} }
${riskRows}=@(${riskRows}|Sort-Object @{Expression='RiskScore';Descending=$true},@{Expression='GroupName';Ascending=$true})
${nodesOut}=@(${nodes}.Values | Where-Object { $null -ne $_ }); foreach (${node} in @(${nodesOut})) { if (${node}.Type -eq 'Group') { ${risk}=@(${riskRows}|Where-Object{$_.GroupId -eq ${node}.Id}|Select-Object -First 1); if (${risk}) { ${node}.RiskScore=${risk}[0].RiskScore; ${node}.RiskDrivers=${risk}[0].RiskDrivers } } }
Export-CsvSafe ${nodesOut} (Join-Path ${out} 'IdentityAudit-Nodes.csv')
Export-CsvSafe ${edges} (Join-Path ${out} 'IdentityAudit-Edges.csv')
Export-CsvSafe ${paths} (Join-Path ${out} 'IdentityAudit-PrivilegedPaths.csv')
Export-CsvSafe ${cycles} (Join-Path ${out} 'IdentityAudit-CircularNesting.csv')
Export-CsvSafe ${nestedStats} (Join-Path ${out} 'IdentityAudit-NestingStats.csv')
Export-CsvSafe ${riskRows} (Join-Path ${out} 'IdentityAudit-RiskScores.csv')
[pscustomobject]@{nodes=${nodesOut};edges=${edges};paths=${paths};cycles=${cycles};riskScores=${riskRows}} | ConvertTo-Json -Depth 6 | Out-File (Join-Path ${out} 'IdentityAudit-Graph.json') -Encoding utf8 -Force

Write-Stage 'Writing dashboards'
${css}="<style>body{font-family:Segoe UI,Arial;margin:32px;background:#f5f7fb;color:#172033}.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}.card,.panel{background:#fff;border:1px solid #d9e0ec;border-radius:12px;padding:14px;margin:14px 0}.value{font-size:28px;font-weight:700}.muted{color:#667}table{border-collapse:collapse;width:100%;font-size:13px}td,th{border-bottom:1px solid #edf1f7;padding:7px;text-align:left}</style>"
${stdDash}=Join-Path ${out} 'IdentityAudit-Dashboard.html'
${std}="<!doctype html><html><head><meta charset='utf-8'><title>Identity Audit Dashboard</title>${css}</head><body><h1>Identity Audit Dashboard</h1><p class='muted'>Run ${runId} | Tenant $(HtmlSafe ${TenantId}) | Cache root $(HtmlSafe ${CacheRoot})</p><div class='cards'><div class='card'>Groups<div class='value'>$(@(${groupRows}).Count)</div></div><div class='card'>Users<div class='value'>$(@(${users}).Count)</div></div><div class='card'>Membership rows<div class='value'>${totalMembershipRows}</div></div><div class='card'>Ownerless groups<div class='value'>$(@(${groupRows}|Where-Object{$_.OwnerCount -eq 0}).Count)</div></div></div>"
${std} += New-HtmlTable 'Top group density by membership percentage' (@(${groupRows}|Select-Object -First 25)) @('GroupName','GroupCategory','MemberCount','UserMemberCount','MembershipDensityPct','UserCoveragePct','DepartmentCount','PrimaryDepartment','OwnerCount','IsDynamicGroup','IsAssignableToRole')
${std} += New-HtmlTable 'Department / group hotspots' (@(${deptRows}|Select-Object -First 25)) @('Department','GroupName','UsersInDepartmentInGroup','DepartmentTotalUsers','DepartmentCoveragePctForGroup','GroupShareFromDepartmentPct','GroupMembershipDensityPct')
${std} += '</body></html>'; ${std} | Out-File ${stdDash} -Encoding utf8 -Force
${graphDash}=Join-Path ${out} 'IdentityAudit-GraphDashboard.html'
${graph}="<!doctype html><html><head><meta charset='utf-8'><title>Identity Graph Analysis</title>${css}</head><body><h1>Identity Graph Analysis</h1><p class='muted'>BloodHound-like graph analysis from Entra group, member, and owner evidence.</p><div class='cards'><div class='card'>Nodes<div class='value'>$(@(${nodesOut}).Count)</div></div><div class='card'>Edges<div class='value'>$(@(${edges}).Count)</div></div><div class='card'>High-value targets<div class='value'>$(${highTargets}.Count)</div></div><div class='card'>Privileged paths<div class='value'>$(@(${paths}).Count)</div></div><div class='card'>Circular nesting<div class='value'>$(@(${cycles}).Count)</div></div><div class='card'>Nested group edges<div class='value'>$(@(${groupEdges}).Count)</div></div></div>"
${graph} += New-HtmlTable 'Top risk-scored groups' (@(${riskRows}|Select-Object -First 25)) @('GroupName','RiskScore','RiskDrivers','MemberCount','OwnerCount','DepartmentCount','MembershipDensityPct','PrivilegedPathCount') 25
${graph} += New-HtmlTable 'Privileged path candidates' (@(${paths}|Select-Object -First 50)) @('StartLabel','StartType','EntryGroup','TargetGroup','HopCount','Path','Risk') 50
${graph} += New-HtmlTable 'Circular group nesting' ${cycles} @('Length','Cycle') 50
${graph} += New-HtmlTable 'Nested group chokepoints' (@(${nestedStats}|Sort-Object @{Expression='NestedGroupMemberCount';Descending=$true},@{Expression='NestedIntoGroupCount';Descending=$true}|Select-Object -First 25)) @('GroupName','NestedGroupMemberCount','NestedIntoGroupCount','MemberCount','DepartmentCount') 25
${graph} += '<p class="muted">Path findings are review candidates, not automatic violations.</p></body></html>'; ${graph} | Out-File ${graphDash} -Encoding utf8 -Force
@('# Identity Audit Evidence Manifest','',"Run ID: ${runId}","Tenant ID: ${TenantId}","Cache root: ${CacheRoot}","Cache max age hours: ${CacheMaxAgeHours}",'','Outputs: users, groups, members, owners, department matrix, density, exceptions, graph nodes, graph edges, paths, cycles, risk scores, dashboards') | Out-File (Join-Path ${out} 'IdentityAudit-Manifest.md') -Encoding utf8 -Force
Write-Stage 'Complete'
Write-Host "Output folder: ${out}"
Write-Host "Dashboard: ${stdDash}"
Write-Host "Graph dashboard: ${graphDash}"
if (${OpenDashboard}) { Invoke-Item ${graphDash} }
