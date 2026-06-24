<#
.SYNOPSIS
Identity Audit Graph V4 entrypoint.

.DESCRIPTION
V4 fixes the V3 PowerShell URI interpolation issue where `$ep?` could be parsed as one variable name, causing the `/members?` or `/transitiveMembers?` segment to disappear.
It patches V3 at runtime into a temporary copy and executes the patched copy.
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

${SourceScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph_V3.ps1"
if (-not (Test-Path -Path ${SourceScriptPath})) {
    throw "V3 source script not found: ${SourceScriptPath}"
}

${PatchedScriptPath} = Join-Path ${env:TEMP} "IdentityAudit.Graph_V4.runtime.ps1"
${ScriptText} = Get-Content -Path ${SourceScriptPath} -Raw

# Fix variables followed by '?' in double-quoted Graph URIs.
# Without braces, PowerShell may parse `$ep?` or `$id?` as a single variable token.
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$gid/$ep?`$select=id,displayName,userPrincipalName,mail&`$top=999', '/v1.0/groups/${gid}/${ep}?`$select=id,displayName,userPrincipalName,mail&`$top=999')
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$id?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility', '/v1.0/groups/${id}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility')
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$id?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime', '/v1.0/groups/${id}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime')

${ScriptText} | Out-File -FilePath ${PatchedScriptPath} -Encoding utf8 -Force

Write-Host "[IdentityAudit] V4 patched endpoint interpolation and is running the corrected script." -ForegroundColor Yellow
& ${PatchedScriptPath} @PSBoundParameters
