<#
.SYNOPSIS
Exports Microsoft Entra ID group membership, department mapping, group density, owners, and an HTML dashboard.

.DESCRIPTION
Graph REST-backed identity audit report. The script uses Invoke-MgGraphRequest with raw JSON parsing to avoid typed Graph SDK deserialization issues.
It produces CSV evidence files and a local HTML dashboard.
#>

[CmdletBinding()]
param(
    [string] ${TenantId},
    [string] ${ClientId},
    [string] ${CertificateThumbprint},
    [string] ${OutputRoot} = ".\IdentityAudit-Evidence",
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
    [switch] ${OpenDashboard}
)

${ErrorActionPreference} = "Stop"

function Write-Stage {
    param([Parameter(Mandatory)] [string] ${Message})
    Write-Host "[IdentityAudit] ${Message}" -ForegroundColor Cyan
}

function Write-Warn {
    param([Parameter(Mandatory)] [string] ${Message})
    Write-Host "[IdentityAudit][WARN] ${Message}" -ForegroundColor Yellow
}

function Ensure-Module {
    param([Parameter(Mandatory)] [string] ${Name})

    if (-not (Get-Module -ListAvailable -Name ${Name})) {
        if (${InstallModules}.IsPresent) {
            Write-Stage "Installing ${Name}"
            Install-Module ${Name} -Scope CurrentUser -Force -AllowClobber
        }
        else {
            throw "Required module '${Name}' is not installed. Re-run with -InstallModules or install it manually."
        }
    }

    Import-Module ${Name} -ErrorAction Stop
}

function Get-P {
    param(
        ${Object},
        [Parameter(Mandatory)] [string] ${Name}
    )

    if ($null -eq ${Object}) { return $null }

    try {
        ${Prop} = ${Object}.PSObject.Properties[${Name}]
        if ($null -ne ${Prop}) { return ${Prop}.Value }
    }
    catch { }

    try {
        return ${Object}[${Name}]
    }
    catch { }

    try {
        ${Additional} = ${Object}.AdditionalProperties
        if ($null -ne ${Additional}) {
            return ${Additional}[${Name}]
        }
    }
    catch { }

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

        foreach (${Item} in @(${Values})) {
            if ($null -ne ${Item}) { ${Rows}.Add(${Item}) }
        }

        ${NextUri} = Get-P -Object ${Page} -Name "@odata.nextLink"
    }

    return ${Rows}.ToArray()
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

function Get-DisplayName {
    param(${Object})

    ${Value} = Get-P -Object ${Object} -Name "displayName"
    if ([string]::IsNullOrWhiteSpace([string] ${Value})) { ${Value} = Get-P -Object ${Object} -Name "userPrincipalName" }
    if ([string]::IsNullOrWhiteSpace([string] ${Value})) { ${Value} = Get-P -Object ${Object} -Name "mail" }
    return ${Value}
}

function Get-GroupTypesText {
    param(${Group})

    ${GroupTypes} = Get-P -Object ${Group} -Name "groupTypes"
    if ($null -eq ${GroupTypes}) { return "" }
    return @(${GroupTypes}) -join ";"
}

function Get-GroupCategory {
    param(${Group})

    ${GroupTypesText} = Get-GroupTypesText -Group ${Group}
    ${SecurityEnabled} = Get-P -Object ${Group} -Name "securityEnabled"
    ${MailEnabled} = Get-P -Object ${Group} -Name "mailEnabled"

    if (${GroupTypesText} -match "Unified") { return "Microsoft365" }
    if (${SecurityEnabled} -eq $true -and ${MailEnabled} -eq $true) { return "MailEnabledSecurity" }
    if (${SecurityEnabled} -eq $true -and ${MailEnabled} -ne $true) { return "Security" }
    if (${MailEnabled} -eq $true -and ${SecurityEnabled} -ne $true) { return "DistributionList" }
    return "Other"
}

function ConvertTo-Percent {
    param(
        [decimal] ${Numerator},
        [decimal] ${Denominator}
    )

    if (${Denominator} -le 0) { return [decimal]0 }
    return [math]::Round((${Numerator} / ${Denominator}) * 100, 2)
}

function HtmlEncode {
    param(${Value})
    if ($null -eq ${Value}) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string] ${Value})
}

function HtmlTable {
    param(
        [Parameter(Mandatory)] [string] ${Title},
        [Parameter(Mandatory)] [object[]] ${Rows},
        [Parameter(Mandatory)] [string[]] ${Columns},
        [int] ${MaxRows} = 25
    )

    ${Html} = "<section class='panel'><h2>$(HtmlEncode ${Title})</h2>"

    if (-not ${Rows} -or ${Rows}.Count -eq 0) {
        ${Html} += "<p class='muted'>No records found.</p></section>"
        return ${Html}
    }

    ${Html} += "<table><thead><tr>"
    foreach (${Column} in ${Columns}) { ${Html} += "<th>$(HtmlEncode ${Column})</th>" }
    ${Html} += "</tr></thead><tbody>"

    foreach (${Row} in (${Rows} | Select-Object -First ${MaxRows})) {
        ${Html} += "<tr>"
        foreach (${Column} in ${Columns}) {
            ${Value} = ""
            if (${Row}.PSObject.Properties.Name -contains ${Column}) { ${Value} = ${Row}.${Column} }
            ${Html} += "<td>$(HtmlEncode ${Value})</td>"
        }
        ${Html} += "</tr>"
    }

    ${Html} += "</tbody></table></section>"
    return ${Html}
}

function Export-CsvSafe {
    param(
        ${Rows},
        [Parameter(Mandatory)] [string] ${Path}
    )

    @(${Rows}) | Export-Csv -Path ${Path} -NoTypeInformation
}

Write-Stage "Preparing modules"
Ensure-Module -Name Microsoft.Graph.Authentication

${RunId} = Get-Date -Format "yyyyMMdd-HHmmss"
${OutputDir} = Join-Path ${OutputRoot} ${RunId}
New-Item -ItemType Directory -Path ${OutputDir} -Force | Out-Null

${ExecutionMode} = "Delegated"
${MembershipMode} = "Direct"
if (${IncludeTransitiveMembership}.IsPresent) { ${MembershipMode} = "Transitive" }

try {
    Write-Stage "Connecting to Microsoft Graph"

    if (-not [string]::IsNullOrWhiteSpace(${TenantId}) -and -not [string]::IsNullOrWhiteSpace(${ClientId}) -and -not [string]::IsNullOrWhiteSpace(${CertificateThumbprint})) {
        ${ExecutionMode} = "AppOnlyCertificate"
        Connect-MgGraph -TenantId ${TenantId} -ClientId ${ClientId} -CertificateThumbprint ${CertificateThumbprint} -NoWelcome
    }
    else {
        Connect-MgGraph -Scopes @("User.Read.All", "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All") -NoWelcome
    }

    ${Context} = Get-MgContext
    if ([string]::IsNullOrWhiteSpace(${TenantId})) { ${TenantId} = ${Context}.TenantId }

    Write-Stage "Collecting users"
    ${UsersUri} = '/v1.0/users?$select=id,displayName,userPrincipalName,mail,department,jobTitle,companyName,accountEnabled,userType,employeeId,createdDateTime&$top=999'
    ${RawUsers} = @(Invoke-GraphPagedJson -Uri ${UsersUri})

    ${Users} = @(
        foreach (${User} in ${RawUsers}) {
            [pscustomobject]@{
                Id                = Get-P -Object ${User} -Name "id"
                DisplayName       = Get-P -Object ${User} -Name "displayName"
                UserPrincipalName = Get-P -Object ${User} -Name "userPrincipalName"
                Mail              = Get-P -Object ${User} -Name "mail"
                Department        = Get-P -Object ${User} -Name "department"
                JobTitle          = Get-P -Object ${User} -Name "jobTitle"
                CompanyName       = Get-P -Object ${User} -Name "companyName"
                AccountEnabled    = Get-P -Object ${User} -Name "accountEnabled"
                UserType          = Get-P -Object ${User} -Name "userType"
                EmployeeId        = Get-P -Object ${User} -Name "employeeId"
                CreatedDateTime   = Get-P -Object ${User} -Name "createdDateTime"
            }
        }
    )

    ${UsersById} = @{}
    foreach (${User} in ${Users}) {
        if (-not [string]::IsNullOrWhiteSpace([string] ${User}.Id)) { ${UsersById}[${User}.Id] = ${User} }
    }

    ${EnabledUserCount} = @(${Users} | Where-Object { ${_}.AccountEnabled -eq $true }).Count
    ${TotalUserCount} = @(${Users}).Count
    ${NoDepartmentUserCount} = @(${Users} | Where-Object { [string]::IsNullOrWhiteSpace([string] ${_}.Department) }).Count

    Write-Stage "Collecting groups"
    ${AdvancedGroupsUri} = '/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility&$top=999'
    ${BasicGroupsUri} = '/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime&$top=999'

    if (-not [string]::IsNullOrWhiteSpace(${GroupIdsFile})) {
        if (-not (Test-Path -Path ${GroupIdsFile})) { throw "GroupIdsFile not found: ${GroupIdsFile}" }
        ${GroupIds} = @(Get-Content -Path ${GroupIdsFile} | Where-Object { -not [string]::IsNullOrWhiteSpace([string] ${_}) })
        ${RawGroups} = @()
        foreach (${GroupId} in ${GroupIds}) {
            ${CleanGroupId} = ${GroupId}.Trim()
            try {
                ${RawGroups} += Invoke-GraphGetJson -Uri "/v1.0/groups/${CleanGroupId}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility"
            }
            catch {
                Write-Warn "Advanced group read failed for ${CleanGroupId}. Retrying baseline properties. $($_.Exception.Message)"
                ${RawGroups} += Invoke-GraphGetJson -Uri "/v1.0/groups/${CleanGroupId}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime"
            }
        }
    }
    else {
        try {
            ${RawGroups} = @(Invoke-GraphPagedJson -Uri ${AdvancedGroupsUri})
        }
        catch {
            Write-Warn "Advanced group properties failed. Retrying baseline group properties. $($_.Exception.Message)"
            ${RawGroups} = @(Invoke-GraphPagedJson -Uri ${BasicGroupsUri})
        }
    }

    ${FilteredGroups} = @(
        foreach (${Group} in ${RawGroups}) {
            ${Category} = Get-GroupCategory -Group ${Group}
            if (${SecurityOnly}.IsPresent -and ${Category} -ne "Security") { continue }
            if (${MailEnabledSecurityOnly}.IsPresent -and ${Category} -ne "MailEnabledSecurity") { continue }
            if (${DistributionListOnly}.IsPresent -and ${Category} -ne "DistributionList") { continue }
            if (${Microsoft365Only}.IsPresent -and ${Category} -ne "Microsoft365") { continue }
            ${Group}
        }
    )

    ${OwnerRows} = New-Object System.Collections.Generic.List[object]
    ${MemberRows} = New-Object System.Collections.Generic.List[object]
    ${GroupRows} = New-Object System.Collections.Generic.List[object]
    ${GroupCount} = 0

    foreach (${Group} in ${FilteredGroups}) {
        ${GroupCount}++
        ${GroupId} = Get-P -Object ${Group} -Name "id"
        ${GroupName} = Get-P -Object ${Group} -Name "displayName"
        ${GroupMail} = Get-P -Object ${Group} -Name "mail"
        ${GroupCategory} = Get-GroupCategory -Group ${Group}
        ${GroupTypesText} = Get-GroupTypesText -Group ${Group}
        ${GroupSecurityEnabled} = Get-P -Object ${Group} -Name "securityEnabled"
        ${GroupMailEnabled} = Get-P -Object ${Group} -Name "mailEnabled"
        ${GroupRoleAssignable} = Get-P -Object ${Group} -Name "isAssignableToRole"
        ${GroupRule} = Get-P -Object ${Group} -Name "membershipRule"
        ${RuleState} = Get-P -Object ${Group} -Name "membershipRuleProcessingState"

        ${PercentComplete} = (${GroupCount} / [math]::Max(1, @(${FilteredGroups}).Count)) * 100
        Write-Progress -Activity "Collecting group membership" -Status ${GroupName} -PercentComplete ${PercentComplete}

        ${Endpoint} = "members"
        if (${IncludeTransitiveMembership}.IsPresent) { ${Endpoint} = "transitiveMembers" }

        try {
            ${Members} = @(Invoke-GraphPagedJson -Uri "/v1.0/groups/${GroupId}/${Endpoint}?`$select=id,displayName,userPrincipalName,mail&`$top=999")
        }
        catch {
            Write-Warn "Member read failed for group '${GroupName}' (${GroupId}): $($_.Exception.Message)"
            ${Members} = @()
        }

        if (-not ${SkipOwners}.IsPresent) {
            try {
                ${Owners} = @(Invoke-GraphPagedJson -Uri "/v1.0/groups/${GroupId}/owners?`$select=id,displayName,userPrincipalName,mail&`$top=999")
            }
            catch {
                Write-Warn "Owner read failed for group '${GroupName}' (${GroupId}): $($_.Exception.Message)"
                ${Owners} = @()
            }

            foreach (${Owner} in ${Owners}) {
                ${OwnerRows}.Add([pscustomobject]@{
                    GroupId          = ${GroupId}
                    GroupName        = ${GroupName}
                    GroupCategory    = ${GroupCategory}
                    OwnerId          = Get-P -Object ${Owner} -Name "id"
                    OwnerDisplayName = Get-DisplayName -Object ${Owner}
                    OwnerUPN         = Get-P -Object ${Owner} -Name "userPrincipalName"
                    OwnerType        = Get-DirectoryObjectType -Object ${Owner}
                })
            }
        }

        foreach (${Member} in ${Members}) {
            ${MemberId} = Get-P -Object ${Member} -Name "id"
            ${MemberType} = Get-DirectoryObjectType -Object ${Member}
            ${MemberDisplayName} = Get-DisplayName -Object ${Member}
            ${MemberUPN} = Get-P -Object ${Member} -Name "userPrincipalName"
            ${MemberMail} = Get-P -Object ${Member} -Name "mail"
            ${MemberDepartment} = ""
            ${MemberJobTitle} = ""
            ${MemberCompanyName} = ""
            ${MemberAccountEnabled} = $null
            ${MemberUserType} = ""
            ${MemberEmployeeId} = ""

            if (${MemberType} -eq "User" -and ${UsersById}.ContainsKey(${MemberId})) {
                ${Profile} = ${UsersById}[${MemberId}]
                ${MemberDepartment} = ${Profile}.Department
                ${MemberJobTitle} = ${Profile}.JobTitle
                ${MemberCompanyName} = ${Profile}.CompanyName
                ${MemberAccountEnabled} = ${Profile}.AccountEnabled
                ${MemberUserType} = ${Profile}.UserType
                ${MemberEmployeeId} = ${Profile}.EmployeeId
                if ([string]::IsNullOrWhiteSpace([string] ${MemberUPN})) { ${MemberUPN} = ${Profile}.UserPrincipalName }
                if ([string]::IsNullOrWhiteSpace([string] ${MemberMail})) { ${MemberMail} = ${Profile}.Mail }
            }

            ${MemberRows}.Add([pscustomobject]@{
                GroupId                 = ${GroupId}
                GroupName               = ${GroupName}
                GroupMail               = ${GroupMail}
                GroupCategory           = ${GroupCategory}
                GroupSecurityEnabled    = ${GroupSecurityEnabled}
                GroupMailEnabled        = ${GroupMailEnabled}
                GroupTypes              = ${GroupTypesText}
                GroupIsAssignableToRole = ${GroupRoleAssignable}
                GroupMembershipRule     = ${GroupRule}
                MembershipRuleState     = ${RuleState}
                MembershipMode          = ${MembershipMode}
                MemberId                = ${MemberId}
                MemberDisplayName       = ${MemberDisplayName}
                MemberUPN               = ${MemberUPN}
                MemberMail              = ${MemberMail}
                MemberType              = ${MemberType}
                MemberDepartment        = ${MemberDepartment}
                MemberJobTitle          = ${MemberJobTitle}
                MemberCompanyName       = ${MemberCompanyName}
                MemberAccountEnabled    = ${MemberAccountEnabled}
                MemberUserType          = ${MemberUserType}
                MemberEmployeeId        = ${MemberEmployeeId}
            })
        }
    }

    Write-Progress -Activity "Collecting group membership" -Completed

    ${TotalMembershipRows} = @(${MemberRows}).Count
    ${OwnerCountByGroupId} = @{}
    foreach (${OwnerGroup} in (@(${OwnerRows}) | Group-Object GroupId)) { ${OwnerCountByGroupId}[${OwnerGroup}.Name] = ${OwnerGroup}.Count }

    foreach (${Group} in ${FilteredGroups}) {
        ${GroupId} = Get-P -Object ${Group} -Name "id"
        ${GroupName} = Get-P -Object ${Group} -Name "displayName"
        ${GroupCategory} = Get-GroupCategory -Group ${Group}
        ${RowsForGroup} = @(${MemberRows} | Where-Object { ${_}.GroupId -eq ${GroupId} })
        ${MemberCount} = ${RowsForGroup}.Count
        ${UserMembers} = @(${RowsForGroup} | Where-Object { ${_}.MemberType -eq "User" })
        ${UniqueUserCount} = @(${UserMembers} | Select-Object -ExpandProperty MemberId -Unique).Count
        ${ActiveUserCount} = @(${UserMembers} | Where-Object { ${_}.MemberAccountEnabled -eq $true } | Select-Object -ExpandProperty MemberId -Unique).Count
        ${Departments} = @(${UserMembers} | Where-Object { -not [string]::IsNullOrWhiteSpace([string] ${_}.MemberDepartment) } | Select-Object -ExpandProperty MemberDepartment -Unique)
        ${DepartmentCount} = @(${Departments}).Count
        ${BlankDepartmentCount} = @(${UserMembers} | Where-Object { [string]::IsNullOrWhiteSpace([string] ${_}.MemberDepartment) }).Count
        ${OwnerCount} = 0
        if (${OwnerCountByGroupId}.ContainsKey(${GroupId})) { ${OwnerCount} = ${OwnerCountByGroupId}[${GroupId}] }

        ${DepartmentGroups} = @(${UserMembers} | Where-Object { -not [string]::IsNullOrWhiteSpace([string] ${_}.MemberDepartment) } | Group-Object MemberDepartment | Sort-Object -Property @{ Expression = "Count"; Descending = $true })
        ${PrimaryDepartment} = ""
        ${PrimaryDepartmentCount} = 0
        if (${DepartmentGroups}.Count -gt 0) {
            ${PrimaryDepartment} = ${DepartmentGroups}[0].Name
            ${PrimaryDepartmentCount} = ${DepartmentGroups}[0].Count
        }

        ${DensityPct} = ConvertTo-Percent -Numerator ${MemberCount} -Denominator ${TotalMembershipRows}
        ${CoveragePct} = ConvertTo-Percent -Numerator ${ActiveUserCount} -Denominator ${EnabledUserCount}
        ${PrimaryDepartmentPct} = ConvertTo-Percent -Numerator ${PrimaryDepartmentCount} -Denominator @(${UserMembers}).Count

        if (${IsEmpty}.IsPresent -and ${MemberCount} -ne 0) { continue }
        if (${MinGroupMembersCount} -gt 0 -and ${MemberCount} -lt ${MinGroupMembersCount}) { continue }

        ${MembershipRule} = Get-P -Object ${Group} -Name "membershipRule"
        ${IsDynamicGroup} = -not [string]::IsNullOrWhiteSpace([string] ${MembershipRule})
        ${IsHighDensityGroup} = ${DensityPct} -ge ${HighDensityPctThreshold}
        ${IsCrossDepartmentGroup} = ${DepartmentCount} -gt 1

        ${GroupRows}.Add([pscustomobject]@{
            GroupId                     = ${GroupId}
            GroupName                   = ${GroupName}
            GroupMail                   = Get-P -Object ${Group} -Name "mail"
            GroupCategory               = ${GroupCategory}
            SecurityEnabled             = Get-P -Object ${Group} -Name "securityEnabled"
            MailEnabled                 = Get-P -Object ${Group} -Name "mailEnabled"
            GroupTypes                  = Get-GroupTypesText -Group ${Group}
            IsDynamicGroup              = ${IsDynamicGroup}
            MembershipRule              = ${MembershipRule}
            MembershipRuleState         = Get-P -Object ${Group} -Name "membershipRuleProcessingState"
            IsAssignableToRole          = Get-P -Object ${Group} -Name "isAssignableToRole"
            OwnerCount                  = ${OwnerCount}
            MemberCount                 = ${MemberCount}
            UserMemberCount             = ${UniqueUserCount}
            ActiveUserMemberCount       = ${ActiveUserCount}
            DepartmentCount             = ${DepartmentCount}
            BlankDepartmentMemberCount  = ${BlankDepartmentCount}
            PrimaryDepartment           = ${PrimaryDepartment}
            PrimaryDepartmentCount      = ${PrimaryDepartmentCount}
            PrimaryDepartmentPctOfGroup = ${PrimaryDepartmentPct}
            MembershipDensityPct        = ${DensityPct}
            UserCoveragePct             = ${CoveragePct}
            IsCrossDepartmentGroup      = ${IsCrossDepartmentGroup}
            IsHighDensityGroup          = ${IsHighDensityGroup}
            CreatedDateTime             = Get-P -Object ${Group} -Name "createdDateTime"
            Visibility                  = Get-P -Object ${Group} -Name "visibility"
        })
    }

    ${GroupRowsForOutput} = @(${GroupRows} | Sort-Object -Property @{ Expression = "MembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true })

    if (${IsEmpty}.IsPresent) {
        ${MemberRowsForOutput} = @()
    }
    elseif (${MinGroupMembersCount} -gt 0) {
        ${AllowedGroupIds} = @{}
        foreach (${GroupRow} in ${GroupRowsForOutput}) { ${AllowedGroupIds}[${GroupRow}.GroupId] = $true }
        ${MemberRowsForOutput} = @(${MemberRows} | Where-Object { ${AllowedGroupIds}.ContainsKey(${_}.GroupId) })
    }
    else {
        ${MemberRowsForOutput} = @(${MemberRows})
    }

    Write-Stage "Building department/group matrix"
    ${DepartmentGroupRows} = New-Object System.Collections.Generic.List[object]
    ${DepartmentUserTotals} = @{}
    foreach (${DepartmentGroup} in (@(${Users}) | Group-Object Department)) {
        ${DepartmentName} = ${DepartmentGroup}.Name
        if ([string]::IsNullOrWhiteSpace([string] ${DepartmentName})) { ${DepartmentName} = "(blank)" }
        ${DepartmentUserTotals}[${DepartmentName}] = ${DepartmentGroup}.Count
    }

    foreach (${GroupDept} in (@(${MemberRowsForOutput}) | Where-Object { ${_}.MemberType -eq "User" } | Group-Object GroupId, MemberDepartment)) {
        ${Sample} = ${GroupDept}.Group[0]
        ${DepartmentName} = ${Sample}.MemberDepartment
        if ([string]::IsNullOrWhiteSpace([string] ${DepartmentName})) { ${DepartmentName} = "(blank)" }
        ${UsersInDepartmentInGroup} = @(${GroupDept}.Group | Select-Object -ExpandProperty MemberId -Unique).Count
        ${DepartmentTotalUsers} = 0
        if (${DepartmentUserTotals}.ContainsKey(${DepartmentName})) { ${DepartmentTotalUsers} = ${DepartmentUserTotals}[${DepartmentName}] }
        ${MatchingGroupRow} = @(${GroupRowsForOutput} | Where-Object { ${_}.GroupId -eq ${Sample}.GroupId } | Select-Object -First 1)
        ${GroupUserMemberCount} = 0
        ${GroupDensityPct} = 0
        if (${MatchingGroupRow}.Count -gt 0) {
            ${GroupUserMemberCount} = ${MatchingGroupRow}[0].UserMemberCount
            ${GroupDensityPct} = ${MatchingGroupRow}[0].MembershipDensityPct
        }

        ${DepartmentGroupRows}.Add([pscustomobject]@{
            Department                    = ${DepartmentName}
            GroupId                       = ${Sample}.GroupId
            GroupName                     = ${Sample}.GroupName
            GroupCategory                 = ${Sample}.GroupCategory
            UsersInDepartmentInGroup      = ${UsersInDepartmentInGroup}
            DepartmentTotalUsers          = ${DepartmentTotalUsers}
            DepartmentCoveragePctForGroup = ConvertTo-Percent -Numerator ${UsersInDepartmentInGroup} -Denominator ${DepartmentTotalUsers}
            GroupShareFromDepartmentPct   = ConvertTo-Percent -Numerator ${UsersInDepartmentInGroup} -Denominator ${GroupUserMemberCount}
            GroupMembershipDensityPct     = ${GroupDensityPct}
            IsRoleAssignableGroup         = ${Sample}.GroupIsAssignableToRole
            IsDynamicGroup                = -not [string]::IsNullOrWhiteSpace([string] ${Sample}.GroupMembershipRule)
        })
    }

    ${DepartmentGroupRowsForOutput} = @(${DepartmentGroupRows} | Sort-Object -Property @{ Expression = "GroupMembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true }, @{ Expression = "Department"; Ascending = $true })

    Write-Stage "Building exception review"
    ${ExceptionRows} = New-Object System.Collections.Generic.List[object]
    foreach (${GroupRow} in ${GroupRowsForOutput}) {
        if (-not ${SkipOwners}.IsPresent -and ${GroupRow}.OwnerCount -eq 0) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "Medium"; Finding = "Group has no owner"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "OwnerCount=0" }) }
        if (${GroupRow}.MemberCount -eq 0) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "Low"; Finding = "Empty group"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "MemberCount=0" }) }
        if (${GroupRow}.IsCrossDepartmentGroup -eq $true) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "Review"; Finding = "Cross-department group"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "DepartmentCount=$(${GroupRow}.DepartmentCount); PrimaryDepartment=$(${GroupRow}.PrimaryDepartment); PrimaryDepartmentPctOfGroup=$(${GroupRow}.PrimaryDepartmentPctOfGroup)" }) }
        if (${GroupRow}.BlankDepartmentMemberCount -gt 0) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "Review"; Finding = "Members missing department"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "BlankDepartmentMemberCount=$(${GroupRow}.BlankDepartmentMemberCount)" }) }
        if (${GroupRow}.IsHighDensityGroup -eq $true) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "Review"; Finding = "High group density"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "MembershipDensityPct=$(${GroupRow}.MembershipDensityPct); Threshold=${HighDensityPctThreshold}" }) }
        if (${GroupRow}.IsAssignableToRole -eq $true) { ${ExceptionRows}.Add([pscustomobject]@{ Severity = "High"; Finding = "Role-assignable group"; GroupId = ${GroupRow}.GroupId; GroupName = ${GroupRow}.GroupName; Detail = "IsAssignableToRole=True" }) }
    }

    Write-Stage "Writing CSV evidence"
    Export-CsvSafe -Rows ${Users} -Path (Join-Path ${OutputDir} "IdentityAudit-Users.csv")
    Export-CsvSafe -Rows ${GroupRowsForOutput} -Path (Join-Path ${OutputDir} "IdentityAudit-Groups.csv")
    Export-CsvSafe -Rows ${MemberRowsForOutput} -Path (Join-Path ${OutputDir} "IdentityAudit-GroupMembers.csv")
    Export-CsvSafe -Rows ${OwnerRows} -Path (Join-Path ${OutputDir} "IdentityAudit-GroupOwners.csv")
    Export-CsvSafe -Rows ${DepartmentGroupRowsForOutput} -Path (Join-Path ${OutputDir} "IdentityAudit-DepartmentGroupMatrix.csv")
    ${DensityRows} = ${GroupRowsForOutput} | Select-Object GroupName, GroupId, GroupCategory, MemberCount, UserMemberCount, ActiveUserMemberCount, MembershipDensityPct, UserCoveragePct, DepartmentCount, PrimaryDepartment, PrimaryDepartmentPctOfGroup, OwnerCount, IsDynamicGroup, IsAssignableToRole, IsHighDensityGroup
    Export-CsvSafe -Rows ${DensityRows} -Path (Join-Path ${OutputDir} "IdentityAudit-GroupDensity.csv")
    Export-CsvSafe -Rows ${ExceptionRows} -Path (Join-Path ${OutputDir} "IdentityAudit-Exceptions.csv")

    Write-Stage "Writing dashboard"
    ${TopDensityGroups} = @(${GroupRowsForOutput} | Sort-Object -Property @{ Expression = "MembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)
    ${TopCoverageGroups} = @(${GroupRowsForOutput} | Sort-Object -Property @{ Expression = "UserCoveragePct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)
    ${CrossDepartmentGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsCrossDepartmentGroup -eq $true } | Sort-Object -Property @{ Expression = "DepartmentCount"; Descending = $true }, @{ Expression = "MembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)
    ${OwnerlessGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.OwnerCount -eq 0 } | Sort-Object -Property @{ Expression = "MembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)
    ${DynamicGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsDynamicGroup -eq $true } | Sort-Object -Property @{ Expression = "MembershipDensityPct"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)
    ${DepartmentHotspots} = @(${DepartmentGroupRowsForOutput} | Sort-Object -Property @{ Expression = "GroupMembershipDensityPct"; Descending = $true }, @{ Expression = "UsersInDepartmentInGroup"; Descending = $true }, @{ Expression = "GroupName"; Ascending = $true } | Select-Object -First 25)

    ${TotalGroups} = @(${GroupRowsForOutput}).Count
    ${DynamicGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsDynamicGroup -eq $true }).Count
    ${RoleAssignableGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsAssignableToRole -eq $true }).Count
    ${OwnerlessGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.OwnerCount -eq 0 }).Count
    ${CrossDepartmentGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsCrossDepartmentGroup -eq $true }).Count
    ${HighDensityGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsHighDensityGroup -eq $true }).Count
    ${DepartmentCountCard} = @(${Users} | Where-Object { -not [string]::IsNullOrWhiteSpace([string] ${_}.Department) } | Select-Object -ExpandProperty Department -Unique).Count

    ${Css} = @"
<style>
body { margin:0; padding:32px; background:#f5f7fb; color:#172033; font-family:Segoe UI,Arial,sans-serif; }
header { margin-bottom:24px; }
h1 { margin:0 0 6px 0; font-size:28px; }
h2 { margin:0 0 14px 0; font-size:18px; }
.meta,.muted,footer { color:#62708a; font-size:13px; }
.cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin:18px 0 24px 0; }
.card,.panel { background:#fff; border:1px solid #d9e0ec; border-radius:14px; box-shadow:0 3px 10px rgba(23,32,51,.04); }
.card { padding:16px; }
.card .label { color:#62708a; font-size:12px; text-transform:uppercase; letter-spacing:.05em; }
.card .value { font-size:28px; font-weight:700; margin-top:8px; }
.panel { padding:18px; margin:16px 0; overflow:auto; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th { text-align:left; color:#34415a; border-bottom:1px solid #d9e0ec; padding:8px; white-space:nowrap; }
td { border-bottom:1px solid #edf1f7; padding:8px; vertical-align:top; }
tr:hover td { background:#f9fbff; }
footer { margin-top:26px; font-size:12px; }
</style>
"@

    ${DashboardPath} = Join-Path ${OutputDir} "IdentityAudit-Dashboard.html"
    ${DashboardHtml} = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Identity Audit Dashboard</title>
${Css}
</head>
<body>
<header>
  <h1>Identity Audit Dashboard</h1>
  <div class="meta">Run ID: $(HtmlEncode ${RunId}) | Tenant: $(HtmlEncode ${TenantId}) | Membership mode: $(HtmlEncode ${MembershipMode})</div>
</header>
<div class="cards">
  <div class="card"><div class="label">Groups</div><div class="value">${TotalGroups}</div></div>
  <div class="card"><div class="label">Users</div><div class="value">${TotalUserCount}</div></div>
  <div class="card"><div class="label">Membership rows</div><div class="value">${TotalMembershipRows}</div></div>
  <div class="card"><div class="label">Departments</div><div class="value">${DepartmentCountCard}</div></div>
  <div class="card"><div class="label">Ownerless groups</div><div class="value">${OwnerlessGroupCount}</div></div>
  <div class="card"><div class="label">Cross-department groups</div><div class="value">${CrossDepartmentGroupCount}</div></div>
  <div class="card"><div class="label">High-density groups</div><div class="value">${HighDensityGroupCount}</div></div>
  <div class="card"><div class="label">Role-assignable groups</div><div class="value">${RoleAssignableGroupCount}</div></div>
  <div class="card"><div class="label">Dynamic groups</div><div class="value">${DynamicGroupCount}</div></div>
  <div class="card"><div class="label">Users missing dept</div><div class="value">${NoDepartmentUserCount}</div></div>
</div>
$(HtmlTable -Title "Top group density by membership percentage" -Rows ${TopDensityGroups} -Columns @("GroupName", "GroupCategory", "MemberCount", "UserMemberCount", "MembershipDensityPct", "UserCoveragePct", "DepartmentCount", "PrimaryDepartment", "PrimaryDepartmentPctOfGroup", "OwnerCount", "IsDynamicGroup", "IsAssignableToRole"))
$(HtmlTable -Title "Top groups by enabled user coverage percentage" -Rows ${TopCoverageGroups} -Columns @("GroupName", "GroupCategory", "ActiveUserMemberCount", "UserCoveragePct", "MembershipDensityPct", "DepartmentCount", "PrimaryDepartment", "OwnerCount"))
$(HtmlTable -Title "Department / group hotspots" -Rows ${DepartmentHotspots} -Columns @("Department", "GroupName", "GroupCategory", "UsersInDepartmentInGroup", "DepartmentTotalUsers", "DepartmentCoveragePctForGroup", "GroupShareFromDepartmentPct", "GroupMembershipDensityPct", "IsRoleAssignableGroup", "IsDynamicGroup"))
$(HtmlTable -Title "Cross-department groups" -Rows ${CrossDepartmentGroups} -Columns @("GroupName", "GroupCategory", "MemberCount", "UserMemberCount", "DepartmentCount", "PrimaryDepartment", "PrimaryDepartmentPctOfGroup", "MembershipDensityPct", "OwnerCount"))
$(HtmlTable -Title "Ownerless groups" -Rows ${OwnerlessGroups} -Columns @("GroupName", "GroupCategory", "MemberCount", "UserMemberCount", "MembershipDensityPct", "DepartmentCount", "IsDynamicGroup", "IsAssignableToRole"))
$(HtmlTable -Title "Dynamic groups" -Rows ${DynamicGroups} -Columns @("GroupName", "GroupCategory", "MemberCount", "UserMemberCount", "MembershipDensityPct", "UserCoveragePct", "DepartmentCount", "PrimaryDepartment", "MembershipRule", "MembershipRuleState"))
<footer>Group density percentage = this group's membership rows divided by all membership rows observed in this run. User coverage percentage = enabled user members in the group divided by all enabled users observed in this run.</footer>
</body>
</html>
"@

    ${DashboardHtml} | Out-File -FilePath ${DashboardPath} -Encoding utf8

    @(
        "# Identity Audit Evidence Manifest",
        "",
        "Run ID: ${RunId}",
        "Run date/time: $((Get-Date).ToString('o'))",
        "Tenant ID: ${TenantId}",
        "Execution mode: ${ExecutionMode}",
        "Membership mode: ${MembershipMode}",
        "",
        "## Outputs",
        "- IdentityAudit-Users.csv",
        "- IdentityAudit-Groups.csv",
        "- IdentityAudit-GroupMembers.csv",
        "- IdentityAudit-GroupOwners.csv",
        "- IdentityAudit-DepartmentGroupMatrix.csv",
        "- IdentityAudit-GroupDensity.csv",
        "- IdentityAudit-Exceptions.csv",
        "- IdentityAudit-Dashboard.html"
    ) | Out-File -FilePath (Join-Path ${OutputDir} "IdentityAudit-Manifest.md") -Encoding utf8

    Write-Stage "Complete"
    Write-Host "Output folder: ${OutputDir}"
    Write-Host "Dashboard: ${DashboardPath}"

    if (${OpenDashboard}.IsPresent -and (Test-Path ${DashboardPath})) { Invoke-Item ${DashboardPath} }
}
finally {
    try { Disconnect-MgGraph | Out-Null } catch { }
}
