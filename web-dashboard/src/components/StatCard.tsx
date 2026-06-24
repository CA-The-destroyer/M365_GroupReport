interface StatCardProps {
  label: string;
  value: unknown;
  detail?: string;
}

export function StatCard({ label, value, detail }: StatCardProps) {
  return (
    <div className="statCard">
      <div className="statLabel">{label}</div>
      <div className="statValue">{String(value ?? 0)}</div>
      {detail ? <div className="statDetail">{detail}</div> : null}
    </div>
  );
}
