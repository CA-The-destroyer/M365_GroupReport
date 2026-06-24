import { useEffect, useMemo, useRef, useState } from 'react';
import cytoscape from 'cytoscape';
import type { Core } from 'cytoscape';
import type { AnyRow, AuditData, PageKey } from './types';
import { DataTable } from './components/DataTable';
import { StatCard } from './components/StatCard';
import { downloadCsv, intValue, sortByNumberDesc, text } from './utils';

const defaultDataUrl = '/data/IdentityAudit-AppData.json';
const highValuePattern = /(admin|privileged|break.?glass|global administrator|role|security administrator|application administrator|owner|tier.?0)/i;

const groupColumns = ['GroupName', 'GroupCategory', 'RiskScore', 'RiskDrivers', 'MemberRows', 'UserMembers', 'GroupMembers', 'ServicePrincipalMembers', 'DeviceMembers', 'Departments', 'Owners', 'IsAssignableToRole', 'IsDynamicGroup'];
const groupUserColumns = ['GroupName', 'GroupCategory', 'RiskScore', 'RiskDrivers', 'Owners', 'MemberDisplayName', 'MemberUPN', 'MemberType', 'MemberDepartment', 'MemberJobTitle', 'MemberCompanyName', 'MemberAccountEnabled', 'MembershipMode', 'IsAssignableToRole', 'IsDynamicGroup'];
const userGroupColumns = ['MemberDisplayName', 'MemberUPN', 'MemberType', 'MemberDepartment', 'MemberJobTitle', 'MemberCompanyName', 'MemberAccountEnabled', 'GroupCount', 'HighestRiskScore', 'HighRiskGroupCount', 'RoleAssignableGroupCount', 'OwnedGroupCount', 'Groups', 'OwnedGroups'];
const ownerColumns = ['OwnerDisplayName', 'OwnerUPN', 'OwnerType', 'OwnedGroupCount', 'OwnedGroups'];
const riskColumns = ['GroupName', 'RiskScore', 'RiskDrivers', 'MemberCount', 'OwnerCount', 'DepartmentCount', 'MembershipDensityPct', 'PrivilegedPathCount', 'IsAssignableToRole', 'IsDynamicGroup'];
const pathColumns = ['StartLabel', 'StartType', 'EntryGroup', 'TargetGroup', 'HopCount', 'Path', 'Risk'];

type GraphEdge = AnyRow & { __edgeId: string; __sourceId: string; __targetId: string; __edgeType: string };
type PathResult = { nodeIds: string[]; edgeIds: string[]; labels: string[] } | null;

function rows(data: AuditData | null, key: keyof AuditData): AnyRow[] {
  const value = data?.[key];
  return Array.isArray(value) ? value as AnyRow[] : [];
}

function PageButton({ page, current, label, onClick }: { page: PageKey; current: PageKey; label: string; onClick: (page: PageKey) => void }) {
  return <button className={page === current ? 'navButton active' : 'navButton'} onClick={() => onClick(page)}>{label}</button>;
}

function UploadData({ onData }: { onData: (data: AuditData) => void }) {
  async function onFile(file?: File) {
    if (!file) return;
    const textContent = await file.text();
    onData(JSON.parse(textContent));
  }
  return (
    <div className="uploadBox">
      <h2>Load identity audit data</h2>
      <p className="muted">No local dashboard JSON was found at <code>{defaultDataUrl}</code>. Generate it with <code>IdentityAudit.BuildWebData.ps1</code>, or upload <code>IdentityAudit-AppData.json</code>.</p>
      <input type="file" accept="application/json,.json" onChange={(e) => onFile(e.target.files?.[0])} />
    </div>
  );
}

function SummaryPage({ data }: { data: AuditData }) {
  const summary = data.summary ?? {};
  const topRisk = sortByNumberDesc(rows(data, 'riskScores'), 'RiskScore').slice(0, 10);
  const topUsers = sortByNumberDesc(rows(data, 'userGroupAssociations'), 'GroupCount').slice(0, 10);
  return (
    <>
      <div className="statGrid">
        <StatCard label="Groups" value={summary.groupCount ?? rows(data, 'groups').length} />
        <StatCard label="Users" value={summary.userCount ?? rows(data, 'users').length} />
        <StatCard label="Membership rows" value={summary.membershipRowCount ?? rows(data, 'members').length} />
        <StatCard label="Ownerless groups" value={summary.ownerlessGroupCount ?? 0} />
        <StatCard label="High-risk groups" value={summary.highRiskGroupCount ?? 0} detail="RiskScore >= 50" />
        <StatCard label="Privileged paths" value={summary.privilegedPathCount ?? rows(data, 'paths').length} />
        <StatCard label="Circular nesting" value={summary.cycleCount ?? rows(data, 'cycles').length} />
        <StatCard label="Graph edges" value={summary.edgeCount ?? rows(data, 'edges').length} />
      </div>
      <DataTable title="Top risk-scored groups" description="Highest risk groups by V10 score." rows={topRisk} columns={riskColumns} exportName="top-risk-groups" />
      <DataTable title="Top users/members by group count" description="Members with the most group associations." rows={topUsers} columns={userGroupColumns} exportName="top-user-group-associations" />
    </>
  );
}

function isHighValueNode(node: AnyRow): boolean {
  return text(node.Type) === 'Group' && (intValue(node.RiskScore) >= 50 || highValuePattern.test(text(node.Label)) || highValuePattern.test(text(node.RiskDrivers)));
}

function buildDirectedAdjacency(edges: GraphEdge[]): Map<string, { next: string; edgeId: string }[]> {
  const map = new Map<string, { next: string; edgeId: string }[]>();
  edges.forEach((edge) => {
    if (!edge.__sourceId || !edge.__targetId) return;
    const list = map.get(edge.__sourceId) ?? [];
    list.push({ next: edge.__targetId, edgeId: edge.__edgeId });
    map.set(edge.__sourceId, list);
  });
  return map;
}

function buildUndirectedNeighbors(edges: GraphEdge[]): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  edges.forEach((edge) => {
    if (!edge.__sourceId || !edge.__targetId) return;
    if (!map.has(edge.__sourceId)) map.set(edge.__sourceId, new Set());
    if (!map.has(edge.__targetId)) map.set(edge.__targetId, new Set());
    map.get(edge.__sourceId)?.add(edge.__targetId);
    map.get(edge.__targetId)?.add(edge.__sourceId);
  });
  return map;
}

function shortestPath(startId: string, targetIds: Set<string>, edges: GraphEdge[]): PathResult {
  if (!startId || !targetIds.size) return null;
  const adjacency = buildDirectedAdjacency(edges);
  const queue: { nodeId: string; nodeIds: string[]; edgeIds: string[] }[] = [{ nodeId: startId, nodeIds: [startId], edgeIds: [] }];
  const seen = new Set<string>([startId]);
  while (queue.length) {
    const item = queue.shift();
    if (!item) break;
    if (targetIds.has(item.nodeId) && item.nodeId !== startId) {
      return { nodeIds: item.nodeIds, edgeIds: item.edgeIds, labels: [] };
    }
    if (item.nodeIds.length > 10) continue;
    (adjacency.get(item.nodeId) ?? []).forEach((next) => {
      if (seen.has(next.next)) return;
      seen.add(next.next);
      queue.push({ nodeId: next.next, nodeIds: [...item.nodeIds, next.next], edgeIds: [...item.edgeIds, next.edgeId] });
    });
  }
  return null;
}

function GraphExplorer({ data }: { data: AuditData }) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const cyRef = useRef<Core | null>(null);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<AnyRow | null>(null);
  const [nodeType, setNodeType] = useState('All');
  const [edgeType, setEdgeType] = useState('All');
  const [minRisk, setMinRisk] = useState(0);
  const [highValueOnly, setHighValueOnly] = useState(false);
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());
  const [pathResult, setPathResult] = useState<PathResult>(null);
  const rawNodes = rows(data, 'nodes');
  const rawEdges = rows(data, 'edges');

  const allNodesById = useMemo(() => new Map(rawNodes.map((node) => [text(node.Id), node])), [rawNodes]);
  const graphEdges: GraphEdge[] = useMemo(() => rawEdges.map((edge, index) => ({ ...edge, __edgeId: `e${index}`, __sourceId: text(edge.SourceId), __targetId: text(edge.TargetId), __edgeType: text(edge.EdgeType) })), [rawEdges]);
  const neighbors = useMemo(() => buildUndirectedNeighbors(graphEdges), [graphEdges]);
  const targetIds = useMemo(() => new Set(rawNodes.filter(isHighValueNode).map((node) => text(node.Id))), [rawNodes]);
  const pathNodeIds = useMemo(() => new Set(pathResult?.nodeIds ?? []), [pathResult]);
  const pathEdgeIds = useMemo(() => new Set(pathResult?.edgeIds ?? []), [pathResult]);

  const visibleNodes = useMemo(() => {
    const q = query.trim().toLowerCase();
    const seeded = new Set<string>();
    expandedIds.forEach((id) => {
      seeded.add(id);
      (neighbors.get(id) ?? new Set()).forEach((neighbor) => seeded.add(neighbor));
    });
    pathNodeIds.forEach((id) => seeded.add(id));

    const filtered = rawNodes.filter((node) => {
      const id = text(node.Id);
      const label = text(node.Label);
      const type = text(node.Type);
      const risk = intValue(node.RiskScore);
      const matchesQuery = !q || id.toLowerCase().includes(q) || label.toLowerCase().includes(q) || text(node.RiskDrivers).toLowerCase().includes(q);
      const matchesType = nodeType === 'All' || type === nodeType;
      const matchesRisk = minRisk <= 0 || risk >= minRisk || seeded.has(id);
      const matchesHighValue = !highValueOnly || isHighValueNode(node) || seeded.has(id);
      return matchesQuery && matchesType && matchesRisk && matchesHighValue;
    }).sort((a, b) => intValue(b.RiskScore) - intValue(a.RiskScore));

    const selectedSlice = q || highValueOnly || minRisk > 0 || nodeType !== 'All' || expandedIds.size > 0 || pathNodeIds.size > 0 ? filtered.slice(0, 500) : filtered.slice(0, 300);
    const byId = new Map(selectedSlice.map((node) => [text(node.Id), node]));
    seeded.forEach((id) => {
      const node = allNodesById.get(id);
      if (node) byId.set(id, node);
    });
    return Array.from(byId.values()).slice(0, 650);
  }, [rawNodes, query, nodeType, minRisk, highValueOnly, expandedIds, pathNodeIds, neighbors, allNodesById]);

  const visibleEdges = useMemo(() => {
    const allowed = new Set(visibleNodes.map((node) => text(node.Id)));
    return graphEdges.filter((edge) => {
      const matchesType = edgeType === 'All' || edge.__edgeType === edgeType;
      const visible = allowed.has(edge.__sourceId) && allowed.has(edge.__targetId);
      const pathEdge = pathEdgeIds.has(edge.__edgeId);
      return (visible && matchesType) || pathEdge;
    }).slice(0, 1500);
  }, [visibleNodes, graphEdges, edgeType, pathEdgeIds]);

  useEffect(() => {
    if (!containerRef.current) return;
    cyRef.current?.destroy();
    const cy = cytoscape({
      container: containerRef.current,
      elements: [
        ...visibleNodes.map((node) => ({ data: { id: text(node.Id), label: text(node.Label), type: text(node.Type), risk: intValue(node.RiskScore), drivers: text(node.RiskDrivers) }, classes: pathNodeIds.has(text(node.Id)) ? 'pathNode' : '' })),
        ...visibleEdges.map((edge) => ({ data: { id: edge.__edgeId, source: edge.__sourceId, target: edge.__targetId, label: edge.__edgeType }, classes: pathEdgeIds.has(edge.__edgeId) ? 'pathEdge' : '' }))
      ],
      style: [
        { selector: 'node', style: { label: 'data(label)', 'font-size': 8, 'text-wrap': 'wrap', 'text-max-width': 90, 'background-color': '#4f6bed', color: '#172033', width: 24, height: 24 } },
        { selector: 'node[type="Group"]', style: { 'background-color': '#7c3aed', shape: 'round-rectangle' } },
        { selector: 'node[type="User"]', style: { 'background-color': '#059669' } },
        { selector: 'node[type="ServicePrincipal"]', style: { 'background-color': '#d97706', shape: 'diamond' } },
        { selector: 'node[risk >= 50]', style: { 'border-width': 4, 'border-color': '#dc2626' } },
        { selector: 'node.pathNode', style: { 'border-width': 6, 'border-color': '#f59e0b', 'background-color': '#f97316' } },
        { selector: 'edge', style: { width: 1, 'line-color': '#a7b1c2', 'target-arrow-shape': 'triangle', 'target-arrow-color': '#a7b1c2', 'curve-style': 'bezier', label: 'data(label)', 'font-size': 6 } },
        { selector: 'edge.pathEdge', style: { width: 5, 'line-color': '#f59e0b', 'target-arrow-color': '#f59e0b' } },
        { selector: ':selected', style: { 'border-width': 5, 'border-color': '#111827', 'line-color': '#111827', 'target-arrow-color': '#111827' } }
      ],
      layout: { name: 'cose', animate: false, fit: true, padding: 30 }
    });
    cy.on('tap', 'node', (event) => setSelected(event.target.data()));
    cyRef.current = cy;
    return () => cy.destroy();
  }, [visibleNodes, visibleEdges, pathNodeIds, pathEdgeIds]);

  function searchNode() {
    const cy = cyRef.current;
    if (!cy || !query.trim()) return;
    const q = query.trim().toLowerCase();
    const match = cy.nodes().find((node) => text(node.data('label')).toLowerCase().includes(q) || text(node.id()).toLowerCase().includes(q));
    if (match) {
      cy.elements().unselect();
      match.select();
      cy.animate({ center: { eles: match }, zoom: 1.6 }, { duration: 350 });
      setSelected(match.data());
    }
  }

  function expandSelected() {
    const selectedId = text(selected?.id);
    if (!selectedId) return;
    setExpandedIds((current) => {
      const next = new Set(current);
      next.add(selectedId);
      (neighbors.get(selectedId) ?? new Set()).forEach((neighbor) => next.add(neighbor));
      return next;
    });
  }

  function showPathToTarget() {
    const selectedId = text(selected?.id);
    if (!selectedId) return;
    const result = shortestPath(selectedId, targetIds, graphEdges);
    if (!result) {
      setPathResult({ nodeIds: [selectedId], edgeIds: [], labels: ['No directed path to a high-risk/high-value group was found within 10 hops.'] });
      return;
    }
    const labels = result.nodeIds.map((id) => text(allNodesById.get(id)?.Label) || id);
    setPathResult({ ...result, labels });
    setExpandedIds((current) => {
      const next = new Set(current);
      result.nodeIds.forEach((id) => next.add(id));
      return next;
    });
  }

  function resetGraph() {
    setSelected(null);
    setExpandedIds(new Set());
    setPathResult(null);
    setQuery('');
    setNodeType('All');
    setEdgeType('All');
    setMinRisk(0);
    setHighValueOnly(false);
  }

  return (
    <section className="panel graphPanel">
      <div className="sectionHead">
        <div>
          <h2>Graph explorer</h2>
          <p className="muted">Cytoscape graph view with filters, neighbor expansion, high-value path search, and filtered graph exports.</p>
        </div>
        <span className="pill">{visibleNodes.length} nodes / {visibleEdges.length} edges</span>
      </div>
      <div className="graphFilters">
        <label>Node type<select value={nodeType} onChange={(e) => setNodeType(e.target.value)}><option>All</option><option>Group</option><option>User</option><option>ServicePrincipal</option><option>Device</option><option>DirectoryObject</option></select></label>
        <label>Edge type<select value={edgeType} onChange={(e) => setEdgeType(e.target.value)}><option>All</option><option>MemberOf</option><option>OwnsGroup</option></select></label>
        <label>Min risk<input type="number" value={minRisk} onChange={(e) => setMinRisk(Number(e.target.value))} /></label>
        <label className="checkLabel"><input type="checkbox" checked={highValueOnly} onChange={(e) => setHighValueOnly(e.target.checked)} /> High-value groups only</label>
      </div>
      <div className="graphToolbar">
        <input className="search" value={query} onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') searchNode(); }} placeholder="Search node label, UPN, group name, or ID..." />
        <button onClick={searchNode}>Find node</button>
        <button onClick={expandSelected} disabled={!selected}>Expand neighbors</button>
        <button onClick={showPathToTarget} disabled={!selected}>Path to high-value group</button>
        <button onClick={() => cyRef.current?.layout({ name: 'cose', animate: true, fit: true, padding: 30 }).run()}>Re-layout</button>
        <button onClick={() => downloadCsv('visible-graph-nodes.csv', visibleNodes)}>Export visible nodes</button>
        <button onClick={() => downloadCsv('visible-graph-edges.csv', visibleEdges)}>Export visible edges</button>
        <button onClick={resetGraph}>Reset</button>
      </div>
      {pathResult?.labels?.length ? <div className="pathBox"><b>Path result:</b> {pathResult.labels.join(' -> ')}</div> : null}
      <div className="graphLayout">
        <div ref={containerRef} className="graphCanvas" />
        <aside className="nodeDetails">
          <h3>Selected node</h3>
          {selected ? Object.entries(selected).map(([key, value]) => <p key={key}><b>{key}:</b> {text(value)}</p>) : <p className="muted">Select a node or search for one.</p>}
          {selected ? <p className="muted small">Tip: use Expand neighbors to add adjacent nodes, then Path to high-value group to highlight the shortest directed path.</p> : null}
        </aside>
      </div>
    </section>
  );
}

function ExportPage({ data }: { data: AuditData }) {
  const [minRisk, setMinRisk] = useState(50);
  const [department, setDepartment] = useState('');
  const detailRows = rows(data, 'groupUserDetail');
  const filteredRisk = rows(data, 'riskScores').filter((row) => intValue(row.RiskScore) >= minRisk);
  const filteredDept = detailRows.filter((row) => !department.trim() || text(row.MemberDepartment).toLowerCase().includes(department.toLowerCase()));
  return (
    <section className="panel">
      <h2>Filtered exports</h2>
      <p className="muted">Browser-side exports from the loaded JSON. These do not change evidence; they create working extracts for review.</p>
      <div className="filterGrid">
        <label>Minimum risk score<input type="number" value={minRisk} onChange={(e) => setMinRisk(Number(e.target.value))} /></label>
        <label>Department contains<input value={department} onChange={(e) => setDepartment(e.target.value)} placeholder="Finance, IT, HR..." /></label>
      </div>
      <div className="buttonRow">
        <button onClick={() => downloadCsv(`groups-risk-${minRisk}-plus.csv`, filteredRisk)}>Export groups with risk &gt;= {minRisk}</button>
        <button onClick={() => downloadCsv('department-membership-filter.csv', filteredDept)}>Export department membership filter</button>
        <button onClick={() => downloadCsv('all-group-user-detail.csv', detailRows)}>Export all group-user detail</button>
        <button onClick={() => downloadCsv('all-user-group-associations.csv', rows(data, 'userGroupAssociations'))}>Export all user-group associations</button>
      </div>
      <DataTable title="Current department/member filter preview" rows={filteredDept.slice(0, 250)} columns={groupUserColumns} exportName="department-filter-preview" />
    </section>
  );
}

export default function App() {
  const [data, setData] = useState<AuditData | null>(null);
  const [page, setPage] = useState<PageKey>('summary');
  const [loadError, setLoadError] = useState<string | null>(null);

  useEffect(() => {
    fetch(defaultDataUrl)
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        return response.json();
      })
      .then(setData)
      .catch((error) => setLoadError(String(error)));
  }, []);

  const currentTitle = useMemo(() => ({
    summary: 'Executive summary', groups: 'Groups', users: 'Users and members', owners: 'Owners', findings: 'Findings', graph: 'Graph explorer', exports: 'Filtered exports', manifest: 'Manifest'
  }[page]), [page]);

  if (!data) {
    return <main className="app"><h1>Identity Audit Web Dashboard</h1>{loadError ? <UploadData onData={setData} /> : <p>Loading dashboard data...</p>}</main>;
  }

  return (
    <main className="app">
      <header className="hero">
        <div>
          <h1>Identity Audit Web Dashboard</h1>
          <p className="muted">Source: {data.summary?.sourceFolder ?? 'IdentityAudit-AppData.json'}</p>
        </div>
        <span className="pill">Generated {data.summary?.generatedAt ?? 'unknown'}</span>
      </header>
      <nav className="navBar">
        <PageButton page="summary" current={page} label="Summary" onClick={setPage} />
        <PageButton page="groups" current={page} label="Groups" onClick={setPage} />
        <PageButton page="users" current={page} label="Users" onClick={setPage} />
        <PageButton page="owners" current={page} label="Owners" onClick={setPage} />
        <PageButton page="findings" current={page} label="Findings" onClick={setPage} />
        <PageButton page="graph" current={page} label="Graph" onClick={setPage} />
        <PageButton page="exports" current={page} label="Exports" onClick={setPage} />
        <PageButton page="manifest" current={page} label="Manifest" onClick={setPage} />
      </nav>
      <h2 className="pageTitle">{currentTitle}</h2>
      {page === 'summary' && <SummaryPage data={data} />}
      {page === 'groups' && <><DataTable title="Group membership summaries" rows={rows(data, 'groupMembershipSummary')} columns={groupColumns} exportName="group-membership-summary" /><DataTable title="Group to users/members" rows={rows(data, 'groupUserDetail')} columns={groupUserColumns} exportName="group-user-detail" /></>}
      {page === 'users' && <DataTable title="User/member to group associations" rows={rows(data, 'userGroupAssociations')} columns={userGroupColumns} exportName="user-group-associations" />}
      {page === 'owners' && <DataTable title="Owner to group associations" rows={rows(data, 'ownerGroupAssociations')} columns={ownerColumns} exportName="owner-group-associations" />}
      {page === 'findings' && <><DataTable title="Risk-scored groups" rows={rows(data, 'riskScores')} columns={riskColumns} exportName="risk-scored-groups" /><DataTable title="Privileged paths" rows={rows(data, 'paths')} columns={pathColumns} exportName="privileged-paths" /><DataTable title="Circular nesting" rows={rows(data, 'cycles')} columns={['Length', 'Cycle']} exportName="circular-nesting" /></>}
      {page === 'graph' && <GraphExplorer data={data} />}
      {page === 'exports' && <ExportPage data={data} />}
      {page === 'manifest' && <section className="panel"><h2>Manifest</h2><pre>{data.manifest || 'No manifest loaded.'}</pre></section>}
    </main>
  );
}
