<#
.SYNOPSIS
Identity Audit Graph V9 compatibility entrypoint.

.DESCRIPTION
V9 now forwards to V10. V10 is self-contained and does not depend on V6, V7, V8, or runtime patching.
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

${TargetScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V10.ps1"
if (-not (Test-Path -Path ${TargetScriptPath})) {
    throw "V10 script not found: ${TargetScriptPath}"
}

Write-Host "[IdentityAudit] V9 forwards to IdentityAudit.Graph_V10.ps1" -ForegroundColor Yellow
& ${TargetScriptPath} @PSBoundParameters
