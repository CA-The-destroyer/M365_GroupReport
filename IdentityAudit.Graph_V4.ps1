<#
.SYNOPSIS
Identity Audit Graph V4 compatibility entrypoint.

.DESCRIPTION
V4 now forwards to V5, which includes the endpoint interpolation fix and dashboard alias-collision fix.
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

${ScriptRootPath} = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace(${ScriptRootPath})) {
    ${ScriptRootPath} = Split-Path -Parent $MyInvocation.MyCommand.Path
}

${TargetScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V5.ps1"
if (-not (Test-Path -Path ${TargetScriptPath})) {
    throw "V5 script not found: ${TargetScriptPath}"
}

Write-Host "[IdentityAudit] V4 forwards to IdentityAudit.Graph_V5.ps1" -ForegroundColor Yellow
& ${TargetScriptPath} @PSBoundParameters
