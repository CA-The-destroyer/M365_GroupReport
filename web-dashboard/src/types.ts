export type AnyRow = Record<string, unknown>;

export interface Summary {
  generatedAt?: string;
  sourceFolder?: string;
  userCount?: number;
  groupCount?: number;
  membershipRowCount?: number;
  ownerRowCount?: number;
  nodeCount?: number;
  edgeCount?: number;
  privilegedPathCount?: number;
  cycleCount?: number;
  ownerlessGroupCount?: number;
  roleAssignableGroupCount?: number;
  highRiskGroupCount?: number;
  groupCountsByCategory?: AnyRow[];
  memberCountsByType?: AnyRow[];
  riskBuckets?: AnyRow[];
}

export interface AuditData {
  schemaVersion?: string;
  summary?: Summary;
  manifest?: string;
  users?: AnyRow[];
  groups?: AnyRow[];
  members?: AnyRow[];
  owners?: AnyRow[];
  departmentMatrix?: AnyRow[];
  exceptions?: AnyRow[];
  riskScores?: AnyRow[];
  nodes?: AnyRow[];
  edges?: AnyRow[];
  paths?: AnyRow[];
  cycles?: AnyRow[];
  nestingStats?: AnyRow[];
  groupUserDetail?: AnyRow[];
  userGroupAssociations?: AnyRow[];
  groupMembershipSummary?: AnyRow[];
  ownerGroupAssociations?: AnyRow[];
}

export type PageKey = 'summary' | 'groups' | 'users' | 'owners' | 'findings' | 'graph' | 'exports' | 'manifest';
