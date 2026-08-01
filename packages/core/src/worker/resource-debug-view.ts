import type { OoxmlResourceMetrics } from '../types/resource-metrics.js';

/** Deterministic color-free view used by browser CSS and Node ANSI emitters. */
export function formatOoxmlResourceDebugReport(
  report: OoxmlResourceMetrics,
  ansi = false,
): string {
  const width = 68;
  const subject = report.scope === 'session' ? 'SESSION' : 'LOAD';
  const status = report.status === 'ok'
    ? report.scope === 'session' ? 'COMPLETE' : 'READY'
    : 'FAILED';
  const title = ` OOXML ${subject}  ${report.format.toUpperCase()}  ${status} `;
  const lines: string[] = [topBorder(title, width)];
  lines.push(row(
    `mode ${report.mode.padEnd(8)}  elapsed ${formatDuration(report.elapsedMs)}`,
    width,
  ));
  if (report.sourceBytes !== undefined) {
    lines.push(row(`source ${formatBytes(report.sourceBytes)}`, width));
  }
  lines.push(separator(' admission ', width));
  lines.push(meterRow(
    'largest entry',
    report.usage?.largestInflatedEntryBytes,
    report.policy.maxArchiveEntryBytes,
    width,
  ));
  lines.push(meterRow(
    'total inflated',
    report.usage?.distinctInflatedBytes,
    report.policy.maxTotalInflatedBytes,
    width,
  ));
  if (report.usage) {
    lines.push(row(
      `entries ${formatInteger(report.usage.archiveEntryCount)}  declared ${formatBytes(report.usage.declaredInflatedBytes)}`,
      width,
    ));
  } else {
    lines.push(row('usage unavailable for this report', width));
  }
  if (report.checkpoints.length > 0) {
    lines.push(separator(' checkpoints ', width));
    for (const point of report.checkpoints) {
      const total = point.usage ? formatBytes(point.usage.distinctInflatedBytes) : '—';
      lines.push(row(
        `${formatDuration(point.elapsedMs).padStart(8)}  ${point.name.padEnd(25)} ${total.padStart(12)}`,
        width,
      ));
    }
  }
  if (report.outcome && Object.keys(report.outcome).length > 0) {
    lines.push(separator(' result ', width));
    lines.push(row(
      Object.entries(report.outcome)
        .map(([key, value]) => `${key} ${formatInteger(value)}`)
        .join('  '),
      width,
    ));
  }
  if (report.error) {
    lines.push(separator(' failure ', width));
    const summary = [
      report.error.code,
      report.error.stage,
      report.error.resource,
      report.error.metric,
    ].filter((value): value is string => !!value).join(' · ');
    lines.push(row(summary || 'unclassified error', width));
  }
  lines.push(`└${'─'.repeat(width)}┘`);
  const text = lines.join('\n');
  if (!ansi) return text;
  const color = report.status === 'ok' ? '\u001b[38;5;42m' : '\u001b[38;5;196m';
  return `${color}${text}\u001b[0m`;
}

export function emitOoxmlResourceDebugReport(report: OoxmlResourceMetrics): void {
  const processLike = (globalThis as { process?: { stdout?: { isTTY?: boolean } } }).process;
  if (typeof window === 'undefined') {
    console.log(formatOoxmlResourceDebugReport(report, processLike?.stdout?.isTTY === true));
    return;
  }
  const color = report.status === 'ok' ? '#22c55e' : '#ef4444';
  console.log(
    `%c${formatOoxmlResourceDebugReport(report)}`,
    `color:${color};background:#0b1020;padding:8px 10px;border-radius:6px;font:12px/1.35 ui-monospace,SFMono-Regular,Menlo,monospace`,
  );
}

function formatBytes(value: number): string {
  if (!Number.isFinite(value) || value < 0) return '—';
  const units = ['B', 'KiB', 'MiB', 'GiB', 'TiB'];
  let amount = value;
  let unit = 0;
  while (amount >= 1024 && unit < units.length - 1) {
    amount /= 1024;
    unit += 1;
  }
  const precision = unit === 0 || amount >= 100 ? 0 : amount >= 10 ? 1 : 2;
  return `${amount.toFixed(precision)} ${units[unit]}`;
}

function formatDuration(value: number): string {
  return value < 1_000 ? `${Math.round(value)} ms` : `${(value / 1_000).toFixed(2)} s`;
}

function formatInteger(value: number): string {
  return Number.isSafeInteger(value) ? value.toLocaleString('en-US') : '—';
}

function topBorder(title: string, width: number): string {
  const clipped = title.slice(0, width);
  return `┌${clipped}${'─'.repeat(width - clipped.length)}┐`;
}

function separator(label: string, width: number): string {
  const clipped = label.slice(0, width);
  return `├${clipped}${'─'.repeat(width - clipped.length)}┤`;
}

function row(content: string, width: number): string {
  const clipped = content.slice(0, width - 2);
  return `│ ${clipped}${' '.repeat(width - clipped.length - 2)} │`;
}

function meterRow(
  label: string,
  observed: number | undefined,
  limit: number | null,
  width: number,
): string {
  const cells = 16;
  const ratio = observed === undefined || limit === null || limit <= 0
    ? 0
    : Math.min(1, observed / limit);
  const filled = Math.round(ratio * cells);
  const meter = `${'█'.repeat(filled)}${'░'.repeat(cells - filled)}`;
  const values = `${observed === undefined ? '—' : formatBytes(observed)} / ${limit === null ? 'off' : formatBytes(limit)}`;
  return row(`${label.padEnd(14)} ${meter}  ${values}`, width);
}
