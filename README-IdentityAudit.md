# Identity Audit v1

This branch adds a Microsoft Graph-based identity audit script while leaving the original `M365GroupReport.ps1` unchanged.

## Script

```powershell
.\IdentityAudit.Graph.ps1
```

## Purpose

Exports Microsoft Entra ID users, groups, group members, group owners, department mappings, group density metrics, review flags, and a local HTML dashboard.

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

## Interactive run

```powershell
.\IdentityAudit.Graph.ps1 -InstallModules -OpenDashboard
```

## App-only certificate run

```powershell
.\IdentityAudit.Graph.ps1 `
  -TenantId "<tenant-id>" `
  -ClientId "<app-id>" `
  -CertificateThumbprint "<thumbprint>" `
  -OutputRoot "C:\AuditEvidence\IdentityAudit" `
  -IncludeTransitiveMembership
```

## Useful filters

```powershell
.\IdentityAudit.Graph.ps1 -SecurityOnly
.\IdentityAudit.Graph.ps1 -Microsoft365Only
.\IdentityAudit.Graph.ps1 -MailEnabledSecurityOnly
.\IdentityAudit.Graph.ps1 -DistributionListOnly
.\IdentityAudit.Graph.ps1 -MinGroupMembersCount 50
.\IdentityAudit.Graph.ps1 -HighDensityPctThreshold 2.5
.\IdentityAudit.Graph.ps1 -GroupIdsFile .\GroupIds.txt
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
