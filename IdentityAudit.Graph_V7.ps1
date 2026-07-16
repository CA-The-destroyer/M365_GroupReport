<#
.SYNOPSIS
Identity Audit Graph V7 compatibility entrypoint.

.DESCRIPTION
V7 now forwards to V8, which patches the graph analytics Sort-Object parser issue.
#>

[CmdletBinding()]
param(
    [string] ${TenantId},
    [string] ${ClientId},
    [string] ${CertificateThumbprint},
    [string] ${OutputRoot} = ".\IdentityAudit-Evidence",
    [string] ${CacheRoot} = ".\IdentityAudit-Cache",
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

${ScriptRootPath} = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace(${ScriptRootPath})) {
    ${ScriptRootPath} = Split-Path -Parent $MyInvocation.MyCommand.Path
}

${TargetScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V8.ps1"
if (-not (Test-Path -Path ${TargetScriptPath})) {
    throw "V8 script not found: ${TargetScriptPath}"
}

Write-Host "[IdentityAudit] V7 forwards to IdentityAudit.Graph_V8.ps1" -ForegroundColor Yellow
& ${TargetScriptPath} @PSBoundParameters
