# Identity Audit v8

This branch adds a Microsoft Graph-based identity audit script while leaving the original `M365GroupReport.ps1` unchanged.

## Versioned script name

Use the current versioned entrypoint:

```powershell
.\IdentityAudit.Graph_V8.ps1
```

`IdentityAudit.Graph_V8.ps1` patches the V7 graph analytics parser issue caused by invalid `Sort-Object` syntax in Windows PowerShell, then runs the graph analysis path.

`IdentityAudit.Graph_V7.ps1` remains available as a compatibility launcher and forwards to V8.

## What V8 adds

V8 creates graph-analysis outputs for identity governance review:

- Identity nodes
- Relationship edges
- Group nesting edges
- Circular group nesting detection
- Privileged path candidates
- Nested group chokepoints
- Risk-scored groups
- Graph JSON for future interactive visualization
- Separate graph dashboard

## Graph output files

Each V8 run adds these files to the normal timestamped output folder:

- `IdentityAudit-Nodes.csv`
- `IdentityAudit-Edges.csv`
- `IdentityAudit-PrivilegedPaths.csv`
- `IdentityAudit-CircularNesting.csv`
- `IdentityAudit-NestingStats.csv`
- `IdentityAudit-RiskScores.csv`
- `IdentityAudit-Graph.json`
- `IdentityAudit-GraphDashboard.html`

## Cache behavior

Default cache folder:

```powershell
.\IdentityAudit-Cache
```

Default cache age:

```powershell
168 hours
```

That is 7 days.

V8 passes cache controls through to V6. Cache is reused by default when cache files exist and are still fresh. Missing or expired cache files are refreshed automatically unless `-UseCacheOnly` is used.

## Common runs

First full run, refresh everything and open the graph dashboard:

```powershell
.\IdentityAudit.Graph_V8.ps1 -InstallModules -RefreshAll -OpenDashboard
```

Normal run, reuse fresh cache and refresh only missing/expired cache:

```powershell
.\IdentityAudit.Graph_V8.ps1 -OpenDashboard
```

Force cache-only graph analysis without connecting to Graph:

```powershell
.\IdentityAudit.Graph_V8.ps1 -UseCacheOnly -OpenDashboard
```

Refresh memberships only, then rebuild graph analytics:

```powershell
.\IdentityAudit.Graph_V8.ps1 -RefreshMemberships -OpenDashboard
```

Tune path depth and high-value target matching:

```powershell
.\IdentityAudit.Graph_V8.ps1 `
  -MaxPathDepth 8 `
  -HighValueGroupPattern "(?i)(admin|privileged|break.?glass|global administrator|application administrator|security administrator|tier.?0|domain)" `
  -OpenDashboard
```

## App-only certificate run

```powershell
.\IdentityAudit.Graph_V8.ps1 `
  -TenantId "<tenant-id>" `
  -ClientId "<app-id>" `
  -CertificateThumbprint "<thumbprint>" `
  -OutputRoot "C:\AuditEvidence\IdentityAudit" `
  -CacheRoot "C:\AuditEvidence\IdentityAuditCache" `
  -RefreshAll `
  -IncludeTransitiveMembership `
  -OpenDashboard
```

## Useful filters

```powershell
.\IdentityAudit.Graph_V8.ps1 -SecurityOnly
.\IdentityAudit.Graph_V8.ps1 -Microsoft365Only
.\IdentityAudit.Graph_V8.ps1 -MailEnabledSecurityOnly
.\IdentityAudit.Graph_V8.ps1 -DistributionListOnly
.\IdentityAudit.Graph_V8.ps1 -MinGroupMembersCount 50
.\IdentityAudit.Graph_V8.ps1 -HighDensityPctThreshold 2.5
.\IdentityAudit.Graph_V8.ps1 -GroupIdsFile .\GroupIds.txt
```

## Standard outputs

Each run still creates the standard evidence files:

- `IdentityAudit-Users.csv`
- `IdentityAudit-Groups.csv`
- `IdentityAudit-GroupMembers.csv`
- `IdentityAudit-GroupOwners.csv`
- `IdentityAudit-DepartmentGroupMatrix.csv`
- `IdentityAudit-GroupDensity.csv`
- `IdentityAudit-Exceptions.csv`
- `IdentityAudit-Dashboard.html`
- `IdentityAudit-Manifest.md`

## Graph dashboard sections

- Top risk-scored groups
- Privileged path candidates
- Circular group nesting
- Nested group chokepoints

## Notes

- Read-only collection.
- Graph path findings are review candidates, not automatic violations.
- Department association comes from the Entra user `department` attribute.
- Transitive membership is available through `-IncludeTransitiveMembership`.
- Cross-department groups are review flags, not automatic findings of inappropriate access.
