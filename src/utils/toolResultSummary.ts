/**
 * Produce a short, human-readable summary of a tool result for display in the
 * collapsed tool card header. Returns an empty string when no useful summary
 * can be extracted (so the caller can decide whether to render anything).
 */
export function toolResultSummary(result: unknown): string {
  if (result === undefined || result === null) return '';

  let str: string;
  try {
    str = typeof result === 'string' ? result : JSON.stringify(result);
  } catch {
    return '';
  }
  if (!str) return '';

  // Try to parse as JSON for structured inspection
  let parsed: unknown;
  try {
    parsed = JSON.parse(str);
  } catch {
    // Plain string result — truncate if needed
    return str.length > 100 ? str.slice(0, 100) + '...' : str;
  }

  if (Array.isArray(parsed)) {
    return `${parsed.length} item${parsed.length !== 1 ? 's' : ''}`;
  }

  if (typeof parsed !== 'object' || parsed === null) {
    const s = String(parsed);
    return s.length > 100 ? s.slice(0, 100) + '...' : s;
  }

  const obj = parsed as Record<string, unknown>;

  // Explicit failure
  if (obj.success === false) {
    const msg = obj.message ?? obj.error;
    return typeof msg === 'string' && msg ? `Error: ${msg}` : 'Failed';
  }

  // Prefer a message field (common in tool responses)
  if (typeof obj.message === 'string' && obj.message) {
    const m = obj.message;
    return m.length > 100 ? m.slice(0, 100) + '...' : m;
  }

  // Grid dimensions (e.g. range reads)
  if (typeof obj.rowCount === 'number' && typeof obj.columnCount === 'number') {
    return `${obj.rowCount} x ${obj.columnCount}`;
  }

  // Single numeric count field
  if (typeof obj.count === 'number') {
    return `${obj.count} item${obj.count !== 1 ? 's' : ''}`;
  }

  // A single array field — summarise by its length and field name
  const arrayEntries = Object.entries(obj).filter(([, v]) => Array.isArray(v));
  if (arrayEntries.length === 1) {
    const [key, val] = arrayEntries[0];
    const arr = val as unknown[];
    return `${arr.length} ${key}`;
  }

  return '';
}
