<#
.SYNOPSIS
Identity Audit Graph V5 compatibility entrypoint.

.DESCRIPTION
V5 now forwards to V6, which adds cache-aware execution while preserving the V5 endpoint and dashboard fixes.
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
    [switch] ${RefreshOwners}
)

${ScriptRootPath} = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace(${ScriptRootPath})) {
    ${ScriptRootPath} = Split-Path -Parent $MyInvocation.MyCommand.Path
}

${TargetScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V6.ps1"
if (-not (Test-Path -Path ${TargetScriptPath})) {
    throw "V6 script not found: ${TargetScriptPath}"
}

Write-Host "[IdentityAudit] V5 forwards to IdentityAudit.Graph_V6.ps1" -ForegroundColor Yellow
& ${TargetScriptPath} @PSBoundParameters
