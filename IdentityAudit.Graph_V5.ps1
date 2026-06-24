<#
.SYNOPSIS
Identity Audit Graph V5 entrypoint.

.DESCRIPTION
V5 fixes two PowerShell parsing/runtime issues before execution:
- Braces Graph URI variables that are followed by '?'.
- Renames compact helper functions H and T to avoid built-in alias collisions such as h = Get-History.
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

${PatchedScriptPath} = Join-Path ${env:TEMP} "IdentityAudit.Graph_V5.runtime.ps1"
${ScriptText} = Get-Content -Path ${SourceScriptPath} -Raw

# Fix variables followed by '?' in double-quoted Graph URIs.
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$gid/$ep?`$select=id,displayName,userPrincipalName,mail&`$top=999', '/v1.0/groups/${gid}/${ep}?`$select=id,displayName,userPrincipalName,mail&`$top=999')
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$id?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility', '/v1.0/groups/${id}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,membershipRule,membershipRuleProcessingState,isAssignableToRole,createdDateTime,visibility')
${ScriptText} = ${ScriptText}.Replace('/v1.0/groups/$id?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime', '/v1.0/groups/${id}?`$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime')

# Avoid PowerShell alias collisions: h is commonly Get-History.
${ScriptText} = ${ScriptText}.Replace('function H($v)', 'function HtmlSafe($v)')
${ScriptText} = ${ScriptText}.Replace('function T([string]$title,$rows,[string[]]$cols)', 'function HtmlTableSafe([string]$title,$rows,[string[]]$cols)')
${ScriptText} = ${ScriptText}.Replace('$(H $title)', '$(HtmlSafe $title)')
${ScriptText} = ${ScriptText}.Replace('$(H $c)', '$(HtmlSafe $c)')
${ScriptText} = ${ScriptText}.Replace('$(H $v)', '$(HtmlSafe $v)')
${ScriptText} = ${ScriptText}.Replace('$(H $TenantId)', '$(HtmlSafe $TenantId)')
${ScriptText} = ${ScriptText}.Replace('T ''Top group density by membership percentage''', 'HtmlTableSafe ''Top group density by membership percentage''')
${ScriptText} = ${ScriptText}.Replace('T ''Top groups by enabled user coverage percentage''', 'HtmlTableSafe ''Top groups by enabled user coverage percentage''')
${ScriptText} = ${ScriptText}.Replace('T ''Department / group hotspots''', 'HtmlTableSafe ''Department / group hotspots''')
${ScriptText} = ${ScriptText}.Replace('T ''Cross-department groups''', 'HtmlTableSafe ''Cross-department groups''')
${ScriptText} = ${ScriptText}.Replace('T ''Ownerless groups''', 'HtmlTableSafe ''Ownerless groups''')
${ScriptText} = ${ScriptText}.Replace('T ''Dynamic groups''', 'HtmlTableSafe ''Dynamic groups''')

${ScriptText} | Out-File -FilePath ${PatchedScriptPath} -Encoding utf8 -Force

Write-Host "[IdentityAudit] V5 patched endpoint interpolation and dashboard alias collisions." -ForegroundColor Yellow
& ${PatchedScriptPath} @PSBoundParameters
