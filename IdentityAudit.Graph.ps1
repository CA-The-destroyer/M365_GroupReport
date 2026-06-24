<# 
.SYNOPSIS
Exports Microsoft Entra ID group membership, department mapping, group density, owners, and an HTML dashboard.

.DESCRIPTION
This script is a Microsoft Graph replacement/enrichment path for the legacy MSOnline-based M365GroupReport.
It produces CSV evidence files plus a dashboard that highlights group density by percentage, department spread,
dynamic groups, role-assignable groups, owner gaps, and cross-department access patterns.

.AUTHENTICATION
Interactive delegated:
  .\IdentityAudit.Graph.ps1

App-only certificate:
  .\IdentityAudit.Graph.ps1 -TenantId "<tenant-id>" -ClientId "<app-id>" -CertificateThumbprint "<thumbprint>"

.NOTES
Requires Microsoft.Graph PowerShell modules.
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

Set-StrictMode -Version Latest
${ErrorActionPreference} = "Stop"

function Write-Stage {
    param([Parameter(Mandatory)][string] ${Message})
    Write-Host "[IdentityAudit] ${Message}" -ForegroundColor Cyan
}

function Ensure-Module {
    param([Parameter(Mandatory)][string] ${Name})

    if (-not (Get-Module -ListAvailable -Name ${Name})) {
        if (${InstallModules}.IsPresent) {
            Write-Stage "Installing module ${Name}"
            Install-Module ${Name} -Scope CurrentUser -Force -AllowClobber
        }
        else {
            throw "Required module '${Name}' is not installed. Re-run with -InstallModules or install it manually."
        }
    }

    Import-Module ${Name} -ErrorAction Stop
}

function Get-AdditionalPropertyValue {
    param(
        [Parameter(Mandatory)] ${Object},
        [Parameter(Mandatory)] [string] ${Name}
    )

    if (${Object}.PSObject.Properties.Name -contains ${Name}) {
        return ${Object}.PSObject.Properties[${Name}].Value
    }

    if (${Object}.PSObject.Properties.Name -contains "AdditionalProperties") {
        if (${Object}.AdditionalProperties -and ${Object}.AdditionalProperties.ContainsKey(${Name})) {
            return ${Object}.AdditionalProperties[${Name}]
        }
    }

    return $null
}

function Get-DirectoryObjectType {
    param([Parameter(Mandatory)] ${Object})

    ${TypeValue} = Get-AdditionalPropertyValue -Object ${Object} -Name "@odata.type"

    switch -Regex ([string]${TypeValue}) {
        "user$"             { return "User" }
        "group$"            { return "Group" }
        "servicePrincipal$" { return "ServicePrincipal" }
        "device$"           { return "Device" }
        "orgContact$"       { return "Contact" }
        "directoryRole$"    { return "DirectoryRole" }
        default             { return "DirectoryObject" }
    }
}

function Get-DirectoryObjectDisplayName {
    param([Parameter(Mandatory)] ${Object})

    ${Value} = Get-AdditionalPropertyValue -Object ${Object} -Name "displayName"
    if ([string]::IsNullOrWhiteSpace([string]${Value})) {
        ${Value} = Get-AdditionalPropertyValue -Object ${Object} -Name "userPrincipalName"
    }
    if ([string]::IsNullOrWhiteSpace([string]${Value})) {
        ${Value} = Get-AdditionalPropertyValue -Object ${Object} -Name "mail"
    }
    return ${Value}
}

function Get-GroupCategory {
    param([Parameter(Mandatory)] ${Group})

    ${GroupTypesText} = ""
    if (${Group}.GroupTypes) {
        ${GroupTypesText} = (${Group}.GroupTypes -join ";")
    }

    if (${GroupTypesText} -match "Unified") {
        return "Microsoft365"
    }

    if (${Group}.SecurityEnabled -eq $true -and ${Group}.MailEnabled -eq $true) {
        return "MailEnabledSecurity"
    }

    if (${Group}.SecurityEnabled -eq $true -and ${Group}.MailEnabled -ne $true) {
        return "Security"
    }

    if (${Group}.MailEnabled -eq $true -and ${Group}.SecurityEnabled -ne $true) {
        return "DistributionList"
    }

    return "Other"
}

function ConvertTo-HtmlEncoded {
    param(${Value})

    if ($null -eq ${Value}) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]${Value})
}

function ConvertTo-Percent {
    param(
        [decimal] ${Numerator},
        [decimal] ${Denominator}
    )

    if (${Denominator} -le 0) {
        return [decimal]0
    }

    return [math]::Round((${Numerator} / ${Denominator}) * 100, 2)
}

function ConvertTo-IdentityAuditHtmlTable {
    param(
        [Parameter(Mandatory)] [string] ${Title},
        [Parameter(Mandatory)] [object[]] ${Rows},
        [Parameter(Mandatory)] [string[]] ${Columns},
        [int] ${MaxRows} = 25
    )

    ${Html} = "<section class='panel'><h2>$(ConvertTo-HtmlEncoded ${Title})</h2>"
    if (-not ${Rows} -or ${Rows}.Count -eq 0) {
        ${Html} += "<p class='muted'>No records found.</p></section>"
        return ${Html}
    }

    ${Html} += "<table><thead><tr>"
    foreach (${Column} in ${Columns}) {
        ${Html} += "<th>$(ConvertTo-HtmlEncoded ${Column})</th>"
    }
    ${Html} += "</tr></thead><tbody>"

    foreach (${Row} in (${Rows} | Select-Object -First ${MaxRows})) {
        ${Html} += "<tr>"
        foreach (${Column} in ${Columns}) {
            ${Value} = ""
            if (${Row}.PSObject.Properties.Name -contains ${Column}) {
                ${Value} = ${Row}.${Column}
            }
            ${Html} += "<td>$(ConvertTo-HtmlEncoded ${Value})</td>"
        }
        ${Html} += "</tr>"
    }

    ${Html} += "</tbody></table></section>"
    return ${Html}
}

function Write-Manifest {
    param(
        [Parameter(Mandatory)] [string] ${Path},
        [Parameter(Mandatory)] [hashtable] ${Values}
    )

    ${Lines} = @(
        "# Identity Audit Evidence Manifest",
        "",
        "Run ID: $(${Values}.RunId)",
        "Run date/time: $(${Values}.RunDateTime)",
        "Tenant ID: $(${Values}.TenantId)",
        "Execution mode: $(${Values}.ExecutionMode)",
        "Membership mode: $(${Values}.MembershipMode)",
        "",
        "## Outputs",
        "",
        "- IdentityAudit-Users.csv",
        "- IdentityAudit-Groups.csv",
        "- IdentityAudit-GroupMembers.csv",
        "- IdentityAudit-GroupOwners.csv",
        "- IdentityAudit-DepartmentGroupMatrix.csv",
        "- IdentityAudit-GroupDensity.csv",
        "- IdentityAudit-Exceptions.csv",
        "- IdentityAudit-Dashboard.html",
        "",
        "## Control notes",
        "",
        "- Read-only Microsoft Graph collection.",
        "- No users, groups, assignments, or policies are modified.",
        "- Group density percentage is calculated from observed membership records in this run.",
        "- User coverage percentage is calculated against enabled Entra user objects collected in this run.",
        "- Department association is derived from the Microsoft Entra user department attribute.",
        ""
    )

    ${Lines} | Out-File -FilePath ${Path} -Encoding utf8
}

Write-Stage "Preparing modules"
Ensure-Module -Name Microsoft.Graph.Authentication
Ensure-Module -Name Microsoft.Graph.Users
Ensure-Module -Name Microsoft.Graph.Groups

${RunId} = Get-Date -Format "yyyyMMdd-HHmmss"
${OutputDir} = Join-Path ${OutputRoot} ${RunId}
New-Item -ItemType Directory -Path ${OutputDir} -Force | Out-Null

${ExecutionMode} = "Delegated"
${GraphScopes} = @(
    "User.Read.All",
    "Group.Read.All",
    "GroupMember.Read.All",
    "Directory.Read.All",
    "RoleManagement.Read.Directory"
)

Write-Stage "Connecting to Microsoft Graph"
if (-not [string]::IsNullOrWhiteSpace(${TenantId}) -and
    -not [string]::IsNullOrWhiteSpace(${ClientId}) -and
    -not [string]::IsNullOrWhiteSpace(${CertificateThumbprint})) {

    ${ExecutionMode} = "AppOnlyCertificate"

    Connect-MgGraph `
        -TenantId ${TenantId} `
        -ClientId ${ClientId} `
        -CertificateThumbprint ${CertificateThumbprint} `
        -NoWelcome
}
else {
    Connect-MgGraph -Scopes ${GraphScopes} -NoWelcome
}

${Context} = Get-MgContext
if ([string]::IsNullOrWhiteSpace(${TenantId})) {
    ${TenantId} = ${Context}.TenantId
}

Write-Stage "Collecting users"
${UserProperties} = @(
    "id",
    "displayName",
    "userPrincipalName",
    "mail",
    "department",
    "jobTitle",
    "companyName",
    "accountEnabled",
    "userType",
    "employeeId",
    "createdDateTime"
)

${Users} = Get-MgUser -All -Property ${UserProperties} |
    Select-Object `
        Id,
        DisplayName,
        UserPrincipalName,
        Mail,
        Department,
        JobTitle,
        CompanyName,
        AccountEnabled,
        UserType,
        EmployeeId,
        CreatedDateTime

${UsersById} = @{}
foreach (${User} in ${Users}) {
    if (-not [string]::IsNullOrWhiteSpace(${User}.Id)) {
        ${UsersById}[${User}.Id] = ${User}
    }
}

${EnabledUserCount} = @(${Users} | Where-Object { ${_}.AccountEnabled -eq $true }).Count
${TotalUserCount} = @(${Users}).Count
${NoDepartmentUserCount} = @(${Users} | Where-Object { [string]::IsNullOrWhiteSpace([string]${_}.Department) }).Count

Write-Stage "Collecting groups"
${GroupProperties} = @(
    "id",
    "displayName",
    "description",
    "mail",
    "mailEnabled",
    "securityEnabled",
    "groupTypes",
    "membershipRule",
    "membershipRuleProcessingState",
    "isAssignableToRole",
    "createdDateTime",
    "visibility"
)

if (-not [string]::IsNullOrWhiteSpace(${GroupIdsFile})) {
    if (-not (Test-Path -Path ${GroupIdsFile})) {
        throw "GroupIdsFile not found: ${GroupIdsFile}"
    }

    ${GroupIds} = Get-Content -Path ${GroupIdsFile} | Where-Object { -not [string]::IsNullOrWhiteSpace(${_}) }
    ${Groups} = foreach (${GroupId} in ${GroupIds}) {
        Get-MgGroup -GroupId ${GroupId}.Trim() -Property ${GroupProperties}
    }
}
else {
    ${Groups} = Get-MgGroup -All -Property ${GroupProperties}
}

${FilteredGroups} = foreach (${Group} in ${Groups}) {
    ${Category} = Get-GroupCategory -Group ${Group}

    if (${SecurityOnly}.IsPresent -and ${Category} -ne "Security") { continue }
    if (${MailEnabledSecurityOnly}.IsPresent -and ${Category} -ne "MailEnabledSecurity") { continue }
    if (${DistributionListOnly}.IsPresent -and ${Category} -ne "DistributionList") { continue }
    if (${Microsoft365Only}.IsPresent -and ${Category} -ne "Microsoft365") { continue }

    ${Group}
}

${GroupRows} = New-Object System.Collections.Generic.List[object]
${OwnerRows} = New-Object System.Collections.Generic.List[object]
${MemberRows} = New-Object System.Collections.Generic.List[object]

${GroupCount} = 0
${MembershipMode} = if (${IncludeTransitiveMembership}.IsPresent) { "Transitive" } else { "Direct" }

foreach (${Group} in ${FilteredGroups}) {
    ${GroupCount}++
    ${Category} = Get-GroupCategory -Group ${Group}
    Write-Progress -Activity "Collecting group membership" -Status ${Group}.DisplayName -PercentComplete ((${GroupCount} / [math]::Max(1, @(${FilteredGroups}).Count)) * 100)

    if (${IncludeTransitiveMembership}.IsPresent) {
        ${Members} = @(Get-MgGroupTransitiveMember -GroupId ${Group}.Id -All -ErrorAction Stop)
    }
    else {
        ${Members} = @(Get-MgGroupMember -GroupId ${Group}.Id -All -ErrorAction Stop)
    }

    if (-not ${SkipOwners}.IsPresent) {
        try {
            ${Owners} = @(Get-MgGroupOwner -GroupId ${Group}.Id -All -ErrorAction Stop)
        }
        catch {
            ${Owners} = @()
            ${OwnerRows}.Add([pscustomobject]@{
                GroupId          = ${Group}.Id
                GroupName        = ${Group}.DisplayName
                GroupCategory    = ${Category}
                OwnerId          = ""
                OwnerDisplayName = "ERROR: $($_.Exception.Message)"
                OwnerUPN         = ""
                OwnerType        = "Error"
            })
        }

        foreach (${Owner} in ${Owners}) {
            ${OwnerType} = Get-DirectoryObjectType -Object ${Owner}
            ${OwnerDisplayName} = Get-DirectoryObjectDisplayName -Object ${Owner}
            ${OwnerUPN} = Get-AdditionalPropertyValue -Object ${Owner} -Name "userPrincipalName"

            ${OwnerRows}.Add([pscustomobject]@{
                GroupId          = ${Group}.Id
                GroupName        = ${Group}.DisplayName
                GroupCategory    = ${Category}
                OwnerId          = ${Owner}.Id
                OwnerDisplayName = ${OwnerDisplayName}
                OwnerUPN         = ${OwnerUPN}
                OwnerType        = ${OwnerType}
            })
        }
    }

    foreach (${Member} in ${Members}) {
        ${ObjectType} = Get-DirectoryObjectType -Object ${Member}
        ${MemberDisplayName} = Get-DirectoryObjectDisplayName -Object ${Member}
        ${MemberMail} = Get-AdditionalPropertyValue -Object ${Member} -Name "mail"
        ${MemberUPN} = Get-AdditionalPropertyValue -Object ${Member} -Name "userPrincipalName"
        ${Department} = ""
        ${JobTitle} = ""
        ${CompanyName} = ""
        ${AccountEnabled} = ""
        ${UserType} = ""
        ${EmployeeId} = ""

        if (${ObjectType} -eq "User" -and ${UsersById}.ContainsKey(${Member}.Id)) {
            ${UserProfile} = ${UsersById}[${Member}.Id]
            ${Department} = ${UserProfile}.Department
            ${JobTitle} = ${UserProfile}.JobTitle
            ${CompanyName} = ${UserProfile}.CompanyName
            ${AccountEnabled} = ${UserProfile}.AccountEnabled
            ${UserType} = ${UserProfile}.UserType
            ${EmployeeId} = ${UserProfile}.EmployeeId

            if ([string]::IsNullOrWhiteSpace([string]${MemberUPN})) {
                ${MemberUPN} = ${UserProfile}.UserPrincipalName
            }
            if ([string]::IsNullOrWhiteSpace([string]${MemberMail})) {
                ${MemberMail} = ${UserProfile}.Mail
            }
        }

        ${MemberRows}.Add([pscustomobject]@{
            GroupId                  = ${Group}.Id
            GroupName                = ${Group}.DisplayName
            GroupMail                = ${Group}.Mail
            GroupCategory            = ${Category}
            GroupSecurityEnabled     = ${Group}.SecurityEnabled
            GroupMailEnabled         = ${Group}.MailEnabled
            GroupTypes               = if (${Group}.GroupTypes) { ${Group}.GroupTypes -join ";" } else { "" }
            GroupIsAssignableToRole  = ${Group}.IsAssignableToRole
            GroupMembershipRule      = ${Group}.MembershipRule
            MembershipRuleState      = ${Group}.MembershipRuleProcessingState
            MembershipMode           = ${MembershipMode}
            MemberId                 = ${Member}.Id
            MemberDisplayName        = ${MemberDisplayName}
            MemberUPN                = ${MemberUPN}
            MemberMail               = ${MemberMail}
            MemberType               = ${ObjectType}
            MemberDepartment         = ${Department}
            MemberJobTitle           = ${JobTitle}
            MemberCompanyName        = ${CompanyName}
            MemberAccountEnabled     = ${AccountEnabled}
            MemberUserType           = ${UserType}
            MemberEmployeeId         = ${EmployeeId}
        })
    }
}

Write-Progress -Activity "Collecting group membership" -Completed

${TotalMembershipRows} = @(${MemberRows}).Count
${AllGroupIdsWithMembers} = @(${MemberRows} | Select-Object -ExpandProperty GroupId -Unique)
${OwnerCountByGroupId} = @{}
foreach (${OwnerGroup} in (${OwnerRows} | Group-Object GroupId)) {
    ${OwnerCountByGroupId}[${OwnerGroup}.Name] = ${OwnerGroup}.Count
}

foreach (${Group} in ${FilteredGroups}) {
    ${Category} = Get-GroupCategory -Group ${Group}
    ${RowsForGroup} = @(${MemberRows} | Where-Object { ${_}.GroupId -eq ${Group}.Id })
    ${MemberCount} = ${RowsForGroup}.Count
    ${UserMembers} = @(${RowsForGroup} | Where-Object { ${_}.MemberType -eq "User" })
    ${UniqueUserMembers} = @(${UserMembers} | Select-Object -ExpandProperty MemberId -Unique)
    ${ActiveUserMembers} = @(${UserMembers} | Where-Object { ${_}.MemberAccountEnabled -eq $true } | Select-Object -ExpandProperty MemberId -Unique)
    ${Departments} = @(${UserMembers} | Where-Object { -not [string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment) } | Select-Object -ExpandProperty MemberDepartment -Unique)
    ${BlankDepartmentCount} = @(${UserMembers} | Where-Object { [string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment) }).Count
    ${OwnerCount} = if (${OwnerCountByGroupId}.ContainsKey(${Group}.Id)) { ${OwnerCountByGroupId}[${Group}.Id] } else { 0 }
    ${PrimaryDepartment} = ""
    ${PrimaryDepartmentCount} = 0

    ${DepartmentGroups} = ${UserMembers} |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]${_}.MemberDepartment) } |
        Group-Object MemberDepartment |
        Sort-Object Count -Descending

    if (${DepartmentGroups} -and ${DepartmentGroups}.Count -gt 0) {
        ${PrimaryDepartment} = ${DepartmentGroups}[0].Name
        ${PrimaryDepartmentCount} = ${DepartmentGroups}[0].Count
    }

    ${UniqueUserMemberCount} = @(${UniqueUserMembers}).Count
    ${ActiveUserMemberCount} = @(${ActiveUserMembers}).Count
    ${DepartmentCountForGroup} = @(${Departments}).Count
    ${UserMemberRowsCount} = @(${UserMembers}).Count

    ${DensityPct} = ConvertTo-Percent -Numerator ${MemberCount} -Denominator ${TotalMembershipRows}
    ${UserCoveragePct} = ConvertTo-Percent -Numerator ${ActiveUserMemberCount} -Denominator ${EnabledUserCount}
    ${PrimaryDepartmentPctOfGroup} = ConvertTo-Percent -Numerator ${PrimaryDepartmentCount} -Denominator ${UserMemberRowsCount}

    if (${IsEmpty}.IsPresent -and ${MemberCount} -ne 0) { continue }
    if (${MinGroupMembersCount} -gt 0 -and ${MemberCount} -lt ${MinGroupMembersCount}) { continue }

    ${GroupRows}.Add([pscustomobject]@{
        GroupId                     = ${Group}.Id
        GroupName                   = ${Group}.DisplayName
        GroupMail                   = ${Group}.Mail
        GroupCategory               = ${Category}
        SecurityEnabled             = ${Group}.SecurityEnabled
        MailEnabled                 = ${Group}.MailEnabled
        GroupTypes                  = if (${Group}.GroupTypes) { ${Group}.GroupTypes -join ";" } else { "" }
        IsDynamicGroup              = -not [string]::IsNullOrWhiteSpace([string]${Group}.MembershipRule)
        MembershipRule              = ${Group}.MembershipRule
        MembershipRuleState         = ${Group}.MembershipRuleProcessingState
        IsAssignableToRole          = ${Group}.IsAssignableToRole
        OwnerCount                  = ${OwnerCount}
        MemberCount                 = ${MemberCount}
        UserMemberCount             = ${UniqueUserMemberCount}
        ActiveUserMemberCount       = ${ActiveUserMemberCount}
        DepartmentCount             = ${DepartmentCountForGroup}
        BlankDepartmentMemberCount  = ${BlankDepartmentCount}
        PrimaryDepartment           = ${PrimaryDepartment}
        PrimaryDepartmentCount      = ${PrimaryDepartmentCount}
        PrimaryDepartmentPctOfGroup = ${PrimaryDepartmentPctOfGroup}
        MembershipDensityPct        = ${DensityPct}
        UserCoveragePct             = ${UserCoveragePct}
        IsCrossDepartmentGroup      = (${DepartmentCountForGroup} -gt 1)
        IsHighDensityGroup          = (${DensityPct} -ge ${HighDensityPctThreshold})
        CreatedDateTime             = ${Group}.CreatedDateTime
        Visibility                  = ${Group}.Visibility
    })
}

${GroupRowsForOutput} = @(${GroupRows} | Sort-Object MembershipDensityPct -Descending)

if (${IsEmpty}.IsPresent) {
    ${MemberRowsForOutput} = @()
}
elseif (${MinGroupMembersCount} -gt 0) {
    ${AllowedGroupIds} = @{}
    foreach (${GroupRow} in ${GroupRowsForOutput}) {
        ${AllowedGroupIds}[${GroupRow}.GroupId] = $true
    }

    ${MemberRowsForOutput} = @(${MemberRows} | Where-Object { ${AllowedGroupIds}.ContainsKey(${_}.GroupId) })
}
else {
    ${MemberRowsForOutput} = @(${MemberRows})
}

Write-Stage "Building department/group matrix"
${DepartmentGroupRows} = New-Object System.Collections.Generic.List[object]
${UserMemberRowsOnly} = @(${MemberRowsForOutput} | Where-Object { ${_}.MemberType -eq "User" })

${DepartmentUserTotals} = @{}
foreach (${DepartmentGroup} in (${Users} | Group-Object Department)) {
    ${DepartmentName} = if ([string]::IsNullOrWhiteSpace([string]${DepartmentGroup}.Name)) { "(blank)" } else { ${DepartmentGroup}.Name }
    ${DepartmentUserTotals}[${DepartmentName}] = ${DepartmentGroup}.Count
}

foreach (${GroupDept} in (${UserMemberRowsOnly} | Group-Object GroupId, MemberDepartment)) {
    ${Sample} = ${GroupDept}.Group[0]
    ${DepartmentName} = if ([string]::IsNullOrWhiteSpace([string]${Sample}.MemberDepartment)) { "(blank)" } else { ${Sample}.MemberDepartment }
    ${GroupRow} = ${GroupRowsForOutput} | Where-Object { ${_}.GroupId -eq ${Sample}.GroupId } | Select-Object -First 1
    ${UsersInDepartmentInGroup} = @(${GroupDept}.Group | Select-Object -ExpandProperty MemberId -Unique).Count
    ${DepartmentTotalUsers} = if (${DepartmentUserTotals}.ContainsKey(${DepartmentName})) { ${DepartmentUserTotals}[${DepartmentName}] } else { 0 }

    ${DepartmentGroupRows}.Add([pscustomobject]@{
        Department                    = ${DepartmentName}
        GroupId                       = ${Sample}.GroupId
        GroupName                     = ${Sample}.GroupName
        GroupCategory                 = ${Sample}.GroupCategory
        UsersInDepartmentInGroup      = ${UsersInDepartmentInGroup}
        DepartmentTotalUsers          = ${DepartmentTotalUsers}
        DepartmentCoveragePctForGroup = ConvertTo-Percent -Numerator ${UsersInDepartmentInGroup} -Denominator ${DepartmentTotalUsers}
        GroupShareFromDepartmentPct   = if ($null -ne ${GroupRow}) { ConvertTo-Percent -Numerator ${UsersInDepartmentInGroup} -Denominator ${GroupRow}.UserMemberCount } else { 0 }
        GroupMembershipDensityPct     = if ($null -ne ${GroupRow}) { ${GroupRow}.MembershipDensityPct } else { 0 }
        IsRoleAssignableGroup         = ${Sample}.GroupIsAssignableToRole
        IsDynamicGroup                = -not [string]::IsNullOrWhiteSpace([string]${Sample}.GroupMembershipRule)
    })
}

Write-Stage "Building exception review"
${ExceptionRows} = New-Object System.Collections.Generic.List[object]

foreach (${GroupRow} in ${GroupRowsForOutput}) {
    if (${GroupRow}.OwnerCount -eq 0 -and -not ${SkipOwners}.IsPresent) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "Medium"
            Finding  = "Group has no owner"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "OwnerCount=0"
        })
    }

    if (${GroupRow}.MemberCount -eq 0) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "Low"
            Finding  = "Empty group"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "MemberCount=0"
        })
    }

    if (${GroupRow}.IsCrossDepartmentGroup -eq $true) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "Review"
            Finding  = "Cross-department group"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "DepartmentCount=$(${GroupRow}.DepartmentCount); PrimaryDepartment=$(${GroupRow}.PrimaryDepartment); PrimaryDepartmentPctOfGroup=$(${GroupRow}.PrimaryDepartmentPctOfGroup)"
        })
    }

    if (${GroupRow}.BlankDepartmentMemberCount -gt 0) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "Review"
            Finding  = "Members missing department"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "BlankDepartmentMemberCount=$(${GroupRow}.BlankDepartmentMemberCount)"
        })
    }

    if (${GroupRow}.IsHighDensityGroup -eq $true) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "Review"
            Finding  = "High group density"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "MembershipDensityPct=$(${GroupRow}.MembershipDensityPct); Threshold=${HighDensityPctThreshold}"
        })
    }

    if (${GroupRow}.IsAssignableToRole -eq $true) {
        ${ExceptionRows}.Add([pscustomobject]@{
            Severity = "High"
            Finding  = "Role-assignable group"
            GroupId  = ${GroupRow}.GroupId
            GroupName = ${GroupRow}.GroupName
            Detail   = "IsAssignableToRole=True"
        })
    }
}

Write-Stage "Writing CSV evidence"
${Users} | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-Users.csv") -NoTypeInformation
${GroupRowsForOutput} | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-Groups.csv") -NoTypeInformation
${MemberRowsForOutput} | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-GroupMembers.csv") -NoTypeInformation
${OwnerRows} | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-GroupOwners.csv") -NoTypeInformation
${DepartmentGroupRows} | Sort-Object GroupMembershipDensityPct -Descending, GroupName, Department | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-DepartmentGroupMatrix.csv") -NoTypeInformation
${GroupRowsForOutput} | Select-Object GroupName,GroupId,GroupCategory,MemberCount,UserMemberCount,ActiveUserMemberCount,MembershipDensityPct,UserCoveragePct,DepartmentCount,PrimaryDepartment,PrimaryDepartmentPctOfGroup,OwnerCount,IsDynamicGroup,IsAssignableToRole,IsHighDensityGroup | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-GroupDensity.csv") -NoTypeInformation
${ExceptionRows} | Export-Csv -Path (Join-Path ${OutputDir} "IdentityAudit-Exceptions.csv") -NoTypeInformation

Write-Stage "Writing dashboard"
${DashboardPath} = Join-Path ${OutputDir} "IdentityAudit-Dashboard.html"

${TopDensityGroups} = @(${GroupRowsForOutput} | Sort-Object MembershipDensityPct -Descending | Select-Object -First 25)
${TopCoverageGroups} = @(${GroupRowsForOutput} | Sort-Object UserCoveragePct -Descending | Select-Object -First 25)
${CrossDepartmentGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsCrossDepartmentGroup -eq $true } | Sort-Object DepartmentCount -Descending, MembershipDensityPct -Descending | Select-Object -First 25)
${OwnerlessGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.OwnerCount -eq 0 } | Sort-Object MembershipDensityPct -Descending | Select-Object -First 25)
${DynamicGroups} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsDynamicGroup -eq $true } | Sort-Object MembershipDensityPct -Descending | Select-Object -First 25)
${DepartmentHotspots} = @(${DepartmentGroupRows} | Sort-Object GroupMembershipDensityPct -Descending, UsersInDepartmentInGroup -Descending | Select-Object -First 25)

${TotalGroups} = @(${GroupRowsForOutput}).Count
${DynamicGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsDynamicGroup -eq $true }).Count
${RoleAssignableGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsAssignableToRole -eq $true }).Count
${OwnerlessGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.OwnerCount -eq 0 }).Count
${CrossDepartmentGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsCrossDepartmentGroup -eq $true }).Count
${HighDensityGroupCount} = @(${GroupRowsForOutput} | Where-Object { ${_}.IsHighDensityGroup -eq $true }).Count
${DepartmentCount} = @(${Users} | Where-Object { -not [string]::IsNullOrWhiteSpace([string]${_}.Department) } | Select-Object -ExpandProperty Department -Unique).Count

${Css} = @"
<style>
:root {
  --bg: #f5f7fb;
  --panel: #ffffff;
  --text: #172033;
  --muted: #62708a;
  --border: #d9e0ec;
  --accent: #1f5eff;
}
body {
  margin: 0;
  padding: 32px;
  background: var(--bg);
  color: var(--text);
  font-family: Segoe UI, Arial, sans-serif;
}
header {
  margin-bottom: 24px;
}
h1 {
  margin: 0 0 6px 0;
  font-size: 28px;
}
h2 {
  margin: 0 0 14px 0;
  font-size: 18px;
}
.meta {
  color: var(--muted);
  font-size: 13px;
}
.cards {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
  gap: 14px;
  margin: 18px 0 24px 0;
}
.card {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 16px;
  box-shadow: 0 3px 10px rgba(23,32,51,.04);
}
.card .label {
  color: var(--muted);
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: .05em;
}
.card .value {
  font-size: 28px;
  font-weight: 700;
  margin-top: 8px;
}
.panel {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 18px;
  margin: 16px 0;
  overflow: auto;
  box-shadow: 0 3px 10px rgba(23,32,51,.04);
}
table {
  border-collapse: collapse;
  width: 100%;
  font-size: 13px;
}
th {
  text-align: left;
  color: #34415a;
  border-bottom: 1px solid var(--border);
  padding: 8px;
  white-space: nowrap;
}
td {
  border-bottom: 1px solid #edf1f7;
  padding: 8px;
  vertical-align: top;
}
tr:hover td {
  background: #f9fbff;
}
.muted {
  color: var(--muted);
}
footer {
  margin-top: 26px;
  color: var(--muted);
  font-size: 12px;
}
</style>
"@

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
  <div class="meta">Run ID: $(ConvertTo-HtmlEncoded ${RunId}) | Tenant: $(ConvertTo-HtmlEncoded ${TenantId}) | Membership mode: $(ConvertTo-HtmlEncoded ${MembershipMode})</div>
</header>

<div class="cards">
  <div class="card"><div class="label">Groups</div><div class="value">${TotalGroups}</div></div>
  <div class="card"><div class="label">Users</div><div class="value">${TotalUserCount}</div></div>
  <div class="card"><div class="label">Membership rows</div><div class="value">${TotalMembershipRows}</div></div>
  <div class="card"><div class="label">Departments</div><div class="value">${DepartmentCount}</div></div>
  <div class="card"><div class="label">Ownerless groups</div><div class="value">${OwnerlessGroupCount}</div></div>
  <div class="card"><div class="label">Cross-department groups</div><div class="value">${CrossDepartmentGroupCount}</div></div>
  <div class="card"><div class="label">High-density groups</div><div class="value">${HighDensityGroupCount}</div></div>
  <div class="card"><div class="label">Role-assignable groups</div><div class="value">${RoleAssignableGroupCount}</div></div>
  <div class="card"><div class="label">Dynamic groups</div><div class="value">${DynamicGroupCount}</div></div>
  <div class="card"><div class="label">Users missing dept</div><div class="value">${NoDepartmentUserCount}</div></div>
</div>

$(ConvertTo-IdentityAuditHtmlTable -Title "Top group density by membership percentage" -Rows ${TopDensityGroups} -Columns @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","UserCoveragePct","DepartmentCount","PrimaryDepartment","PrimaryDepartmentPctOfGroup","OwnerCount","IsDynamicGroup","IsAssignableToRole") -MaxRows 25)

$(ConvertTo-IdentityAuditHtmlTable -Title "Top groups by enabled user coverage percentage" -Rows ${TopCoverageGroups} -Columns @("GroupName","GroupCategory","ActiveUserMemberCount","UserCoveragePct","MembershipDensityPct","DepartmentCount","PrimaryDepartment","OwnerCount") -MaxRows 25)

$(ConvertTo-IdentityAuditHtmlTable -Title "Department / group hotspots" -Rows ${DepartmentHotspots} -Columns @("Department","GroupName","GroupCategory","UsersInDepartmentInGroup","DepartmentTotalUsers","DepartmentCoveragePctForGroup","GroupShareFromDepartmentPct","GroupMembershipDensityPct","IsRoleAssignableGroup","IsDynamicGroup") -MaxRows 25)

$(ConvertTo-IdentityAuditHtmlTable -Title "Cross-department groups" -Rows ${CrossDepartmentGroups} -Columns @("GroupName","GroupCategory","MemberCount","UserMemberCount","DepartmentCount","PrimaryDepartment","PrimaryDepartmentPctOfGroup","MembershipDensityPct","OwnerCount") -MaxRows 25)

$(ConvertTo-IdentityAuditHtmlTable -Title "Ownerless groups" -Rows ${OwnerlessGroups} -Columns @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","DepartmentCount","IsDynamicGroup","IsAssignableToRole") -MaxRows 25)

$(ConvertTo-IdentityAuditHtmlTable -Title "Dynamic groups" -Rows ${DynamicGroups} -Columns @("GroupName","GroupCategory","MemberCount","UserMemberCount","MembershipDensityPct","UserCoveragePct","DepartmentCount","PrimaryDepartment","MembershipRule","MembershipRuleState") -MaxRows 25)

<footer>
  Group density percentage = this group's membership rows divided by all membership rows observed in this run.
  User coverage percentage = enabled user members in the group divided by all enabled users observed in this run.
</footer>
</body>
</html>
"@

${DashboardHtml} | Out-File -FilePath ${DashboardPath} -Encoding utf8

Write-Manifest -Path (Join-Path ${OutputDir} "IdentityAudit-Manifest.md") -Values @{
    RunId = ${RunId}
    RunDateTime = (Get-Date).ToString("o")
    TenantId = ${TenantId}
    ExecutionMode = ${ExecutionMode}
    MembershipMode = ${MembershipMode}
}

Write-Stage "Complete"
Write-Host "Output folder: ${OutputDir}"
Write-Host "Dashboard: ${DashboardPath}"

if (${OpenDashboard}.IsPresent -and (Test-Path ${DashboardPath})) {
    Invoke-Item ${DashboardPath}
}

Disconnect-MgGraph | Out-Null
