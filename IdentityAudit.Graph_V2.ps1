<#
.SYNOPSIS
Identity Audit Graph V2 entrypoint.

.DESCRIPTION
Versioned launcher for the Identity Audit Graph script.
Adds Windows PowerShell compatibility for ConvertFrom-Json -Depth before invoking the core implementation.
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

# Windows PowerShell 5.1 does not support ConvertFrom-Json -Depth.
# The core script was written for newer PowerShell behavior, so V2 safely ignores -Depth
# while preserving pipeline behavior.
function ConvertFrom-Json {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        [string] ${InputObject},
        [int] ${Depth}
    )

    begin {
        ${Buffer} = New-Object System.Text.StringBuilder
    }

    process {
        if ($null -ne ${InputObject}) {
            [void] ${Buffer}.AppendLine(${InputObject})
        }
    }

    end {
        ${JsonText} = ${Buffer}.ToString()
        if ([string]::IsNullOrWhiteSpace(${JsonText})) { return $null }
        return Microsoft.PowerShell.Utility\ConvertFrom-Json -InputObject ${JsonText}
    }
}

${ScriptRootPath} = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace(${ScriptRootPath})) {
    ${ScriptRootPath} = Split-Path -Parent $MyInvocation.MyCommand.Path
}

${CoreScriptPath} = Join-Path ${ScriptRootPath} "IdentityAudit.Graph.ps1"
if (-not (Test-Path -Path ${CoreScriptPath})) {
    throw "Core script not found: ${CoreScriptPath}"
}

& ${CoreScriptPath} @PSBoundParameters
