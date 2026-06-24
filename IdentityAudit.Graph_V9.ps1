<#
.SYNOPSIS
Identity Audit Graph V9 entrypoint.

.DESCRIPTION
V9 patches the V7 graph analytics runtime copy so it keeps the original repository script root instead of resolving dependencies from the temporary runtime folder.
It also includes the V8 Sort-Object parser fix.
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

${PatchedScriptPath} = Join-Path ${env:TEMP} "IdentityAudit.Graph_V9.runtime.ps1"
${ScriptText} = Get-Content -Path ${SourceScriptPath} -Raw

# Fix invalid Sort-Object syntax from V7.
${ScriptText} = ${ScriptText}.Replace('$riskRows=@($riskRows|Sort-Object RiskScore -Descending,GroupName)', '$riskRows=@($riskRows|Sort-Object @{Expression=''RiskScore'';Descending=$true},@{Expression=''GroupName'';Ascending=$true})')

# Preserve original repo script root so the runtime copy can find IdentityAudit.Graph_V6.ps1.
${EscapedScriptRootPath} = ${ScriptRootPath}.Replace("'", "''")
${OriginalRootLine} = '$root=$PSScriptRoot;if([string]::IsNullOrWhiteSpace($root)){$root=Split-Path -Parent $MyInvocation.MyCommand.Path}'
${ReplacementRootLine} = "`$root='${EscapedScriptRootPath}'"
${ScriptText} = ${ScriptText}.Replace(${OriginalRootLine}, ${ReplacementRootLine})

${ScriptText} | Out-File -FilePath ${PatchedScriptPath} -Encoding utf8 -Force

Write-Host "[IdentityAudit] V9 patched graph analytics parser issue and preserved repo dependency path." -ForegroundColor Yellow
& ${PatchedScriptPath} @PSBoundParameters
