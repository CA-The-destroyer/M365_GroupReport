import type { AnyRow } from './types';

export function text(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value);
}

export function intValue(value: unknown): number {
  const parsed = Number.parseInt(text(value), 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function matches(row: AnyRow, query: string): boolean {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return true;
  return Object.values(row).some((value) => text(value).toLowerCase().includes(normalized));
}

export function toCsv(rows: AnyRow[]): string {
  if (!rows.length) return '';
  const columns = Array.from(rows.reduce((set, row) => {
    Object.keys(row).forEach((key) => set.add(key));
    return set;
  }, new Set<string>()));
  const escape = (value: unknown) => {
    const raw = text(value);
    if (/[",\n\r]/.test(raw)) return `"${raw.replace(/"/g, '""')}"`;
    return raw;
  };
  return [columns.join(','), ...rows.map((row) => columns.map((col) => escape(row[col])).join(','))].join('\n');
}

export function downloadCsv(filename: string, rows: AnyRow[]): void {
  const blob = new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function sortByNumberDesc(rows: AnyRow[], key: string): AnyRow[] {
  return [...rows].sort((a, b) => intValue(b[key]) - intValue(a[key]));
}
