import { useMemo, useState } from 'react';
import type { AnyRow } from '../types';
import { downloadCsv, matches, text } from '../utils';

interface DataTableProps {
  title: string;
  description?: string;
  rows: AnyRow[];
  columns: string[];
  exportName: string;
  maxRows?: number;
}

export function DataTable({ title, description, rows, columns, exportName, maxRows = 500 }: DataTableProps) {
  const [query, setQuery] = useState('');
  const filtered = useMemo(() => rows.filter((row) => matches(row, query)), [rows, query]);
  const visible = filtered.slice(0, maxRows);

  return (
    <section className="panel">
      <div className="sectionHead">
        <div>
          <h2>{title}</h2>
          {description ? <p className="muted">{description}</p> : null}
        </div>
        <div className="actions">
          <span className="pill">{filtered.length.toLocaleString()} rows</span>
          <button onClick={() => downloadCsv(`${exportName}.csv`, filtered)}>Export filtered CSV</button>
        </div>
      </div>
      <input className="search" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search this table..." />
      <div className="tableWrap">
        <table>
          <thead>
            <tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr>
          </thead>
          <tbody>
            {visible.map((row, index) => (
              <tr key={index}>{columns.map((column) => <td key={column}>{text(row[column])}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
      {filtered.length > maxRows ? <p className="muted small">Showing first {maxRows.toLocaleString()} rows. Export includes all filtered rows.</p> : null}
    </section>
  );
}
