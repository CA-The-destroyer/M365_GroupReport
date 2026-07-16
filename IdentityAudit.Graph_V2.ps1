<#
.SYNOPSIS
Identity Audit Graph V2 compatibility entrypoint.

.DESCRIPTION
V2 now forwards to the self-contained V3 script to avoid the older core-script inline-if issue.
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

${TargetScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V3.ps1"
if (-not (Test-Path -Path ${TargetScriptPath})) {
    throw "V3 script not found: ${TargetScriptPath}"
}

Write-Host "[IdentityAudit] V2 forwards to IdentityAudit.Graph_V3.ps1" -ForegroundColor Yellow
& ${TargetScriptPath} @PSBoundParameters
