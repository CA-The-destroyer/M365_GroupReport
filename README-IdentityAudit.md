# Identity Audit v6

This branch adds a Microsoft Graph-based identity audit script while leaving the original `M365GroupReport.ps1` unchanged.

## Versioned script name

Use the current versioned entrypoint:

```powershell
.\IdentityAudit.Graph_V6.ps1
```

`IdentityAudit.Graph_V6.ps1` adds cache-aware execution while preserving the previous endpoint and dashboard fixes.

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

V6 uses cache by default when cache files exist and are still fresh. Missing or expired cache files are refreshed automatically unless `-UseCacheOnly` is used.

## Cache files

- `users.json`
- `groups.json`
- `members.direct.json`
- `members.transitive.json` when `-IncludeTransitiveMembership` is used
- `owners.json`

## Common runs

First full run, refresh everything and open dashboard:

```powershell
.\IdentityAudit.Graph_V6.ps1 -InstallModules -RefreshAll -OpenDashboard
```

Normal run, reuse fresh cache and refresh only missing/expired cache:

```powershell
.\IdentityAudit.Graph_V6.ps1 -OpenDashboard
```

Force cache-only report generation without connecting to Graph:

```powershell
.\IdentityAudit.Graph_V6.ps1 -UseCacheOnly -OpenDashboard
```

Refresh memberships only:

```powershell
.\IdentityAudit.Graph_V6.ps1 -RefreshMemberships -OpenDashboard
```

Refresh users and groups, then regenerate dependent memberships and owners:

```powershell
.\IdentityAudit.Graph_V6.ps1 -RefreshUsers -RefreshGroups -OpenDashboard
```

Use a custom cache folder and cache age:

```powershell
.\IdentityAudit.Graph_V6.ps1 -CacheRoot "C:\AuditEvidence\IdentityAuditCache" -CacheMaxAgeHours 24 -OpenDashboard
```

## App-only certificate run

```powershell
.\IdentityAudit.Graph_V6.ps1 `
  -TenantId "<tenant-id>" `
  -ClientId "<app-id>" `
  -CertificateThumbprint "<thumbprint>" `
  -OutputRoot "C:\AuditEvidence\IdentityAudit" `
  -CacheRoot "C:\AuditEvidence\IdentityAuditCache" `
  -RefreshAll `
  -IncludeTransitiveMembership
```

## Useful filters

```powershell
.\IdentityAudit.Graph_V6.ps1 -SecurityOnly
.\IdentityAudit.Graph_V6.ps1 -Microsoft365Only
.\IdentityAudit.Graph_V6.ps1 -MailEnabledSecurityOnly
.\IdentityAudit.Graph_V6.ps1 -DistributionListOnly
.\IdentityAudit.Graph_V6.ps1 -MinGroupMembersCount 50
.\IdentityAudit.Graph_V6.ps1 -HighDensityPctThreshold 2.5
.\IdentityAudit.Graph_V6.ps1 -GroupIdsFile .\GroupIds.txt
```

## Outputs

Each run creates a timestamped folder under `.\IdentityAudit-Evidence\` containing:

- `IdentityAudit-Users.csv`
- `IdentityAudit-Groups.csv`
- `IdentityAudit-GroupMembers.csv`
- `IdentityAudit-GroupOwners.csv`
- `IdentityAudit-DepartmentGroupMatrix.csv`
- `IdentityAudit-GroupDensity.csv`
- `IdentityAudit-Exceptions.csv`
- `IdentityAudit-Dashboard.html`
- `IdentityAudit-Manifest.md`

## Group density

`MembershipDensityPct` is calculated as:

```text
Group membership rows / all membership rows observed in the run * 100
```

`UserCoveragePct` is calculated as:

```text
Enabled user members in the group / all enabled users observed in the run * 100
```

## Dashboard sections

- Summary cards
- Top group density by membership percentage
- Top groups by enabled user coverage percentage
- Department/group hotspots
- Cross-department groups
- Ownerless groups
- Dynamic groups

## Notes

- Read-only collection.
- Department association comes from the Entra user `department` attribute.
- Transitive membership is available through `-IncludeTransitiveMembership`.
- Cross-department groups are review flags, not automatic findings of inappropriate access.
