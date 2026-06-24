<#
.SYNOPSIS
Identity Audit Graph V8 entrypoint.

.DESCRIPTION
V8 patches the V7 graph analytics parser issue caused by invalid Sort-Object syntax in Windows PowerShell, then executes a corrected runtime copy.
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

${SourceScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V7.ps1"
if (-not (Test-Path -Path ${SourceScriptPath})) {
    throw "V7 source script not found: ${SourceScriptPath}"
}

${PatchedScriptPath} = Join-Path ${env:TEMP} "IdentityAudit.Graph_V8.runtime.ps1"
${ScriptText} = Get-Content -Path ${SourceScriptPath} -Raw

# Fix invalid Sort-Object syntax: Sort-Object RiskScore -Descending,GroupName
${ScriptText} = ${ScriptText}.Replace('$riskRows=@($riskRows|Sort-Object RiskScore -Descending,GroupName)', '$riskRows=@($riskRows|Sort-Object @{Expression=''RiskScore'';Descending=$true},@{Expression=''GroupName'';Ascending=$true})')

${ScriptText} | Out-File -FilePath ${PatchedScriptPath} -Encoding utf8 -Force

Write-Host "[IdentityAudit] V8 patched V7 graph analytics Sort-Object parser issue." -ForegroundColor Yellow
& ${PatchedScriptPath} @PSBoundParameters
