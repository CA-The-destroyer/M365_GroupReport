# Identity Audit v10

This branch adds a Microsoft Graph-based identity audit script while leaving the original `M365GroupReport.ps1` unchanged.

## Versioned script name

Use the current versioned entrypoint:

```powershell
.\IdentityAudit.Graph_V10.ps1
```

`IdentityAudit.Graph_V10.ps1` is self-contained. It does not depend on V6, V7, V8, V9, or a temporary runtime patch chain.

`IdentityAudit.Graph_V9.ps1` remains available as a compatibility launcher and forwards to V10.

## What V10 includes

V10 includes collection, caching, standard evidence output, standard dashboard output, and graph-analysis output in one script:

- Microsoft Graph collection
- Cache-aware execution
- Identity users/groups/members/owners CSV evidence
- Department/group matrix
- Group density metrics
- Exception review
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

Each V10 run adds these files to the normal timestamped output folder:

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

V10 reuses cache by default when cache files exist and are still fresh. Missing or expired cache files are refreshed automatically unless `-UseCacheOnly` is used.

## Common runs

First full run, refresh everything and open the graph dashboard:

```powershell
.\IdentityAudit.Graph_V10.ps1 -InstallModules -RefreshAll -OpenDashboard
```

Normal run, reuse fresh cache and refresh only missing/expired cache:

```powershell
.\IdentityAudit.Graph_V10.ps1 -OpenDashboard
```

Force cache-only graph analysis without connecting to Graph:

```powershell
.\IdentityAudit.Graph_V10.ps1 -UseCacheOnly -OpenDashboard
```

Refresh memberships only, then rebuild graph analytics:

```powershell
.\IdentityAudit.Graph_V10.ps1 -RefreshMemberships -OpenDashboard
```

Tune path depth and high-value target matching:

```powershell
.\IdentityAudit.Graph_V10.ps1 `
  -MaxPathDepth 8 `
  -HighValueGroupPattern "(?i)(admin|privileged|break.?glass|global administrator|application administrator|security administrator|tier.?0|domain)" `
  -OpenDashboard
```

## App-only certificate run

```powershell
.\IdentityAudit.Graph_V10.ps1 `
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
.\IdentityAudit.Graph_V10.ps1 -SecurityOnly
.\IdentityAudit.Graph_V10.ps1 -Microsoft365Only
.\IdentityAudit.Graph_V10.ps1 -MailEnabledSecurityOnly
.\IdentityAudit.Graph_V10.ps1 -DistributionListOnly
.\IdentityAudit.Graph_V10.ps1 -MinGroupMembersCount 50
.\IdentityAudit.Graph_V10.ps1 -HighDensityPctThreshold 2.5
.\IdentityAudit.Graph_V10.ps1 -GroupIdsFile .\GroupIds.txt
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
