/**
 * Integration tests for toolResultSummary().
 *
 * Validates the short, human-readable summary generation for tool results
 * displayed in collapsed tool card headers.
 */
import { describe, it, expect } from 'vitest';
import { toolResultSummary } from '@/utils/toolResultSummary';

describe('Integration: toolResultSummary', () => {
  // --- null / undefined / empty ---
  it('returns empty string for null', () => {
    expect(toolResultSummary(null)).toBe('');
  });

  it('returns empty string for undefined', () => {
    expect(toolResultSummary(undefined)).toBe('');
  });

  // --- plain strings ---
  it('returns short string as-is', () => {
    expect(toolResultSummary('Done')).toBe('Done');
  });

  it('truncates long strings to 100 chars + ellipsis', () => {
    const long = 'x'.repeat(150);
    const result = toolResultSummary(long);
    expect(result).toHaveLength(103); // 100 + '...'
    expect(result.endsWith('...')).toBe(true);
  });

  it('returns exactly 100 char string without truncation', () => {
    const exact = 'a'.repeat(100);
    expect(toolResultSummary(exact)).toBe(exact);
  });

  // --- JSON array results (serialised) ---
  it('summarises single-element array', () => {
    expect(toolResultSummary(JSON.stringify([1]))).toBe('1 item');
  });

  it('summarises multi-element array with plural', () => {
    expect(toolResultSummary(JSON.stringify([1, 2, 3]))).toBe('3 items');
  });

  it('summarises empty array as "0 items"', () => {
    expect(toolResultSummary(JSON.stringify([]))).toBe('0 items');
  });

  // --- success === false ---
  it('shows error message when success is false with message', () => {
    const result = toolResultSummary(JSON.stringify({ success: false, message: 'Not found' }));
    expect(result).toBe('Error: Not found');
  });

  it('shows error from error field when success is false', () => {
    const result = toolResultSummary(JSON.stringify({ success: false, error: 'Timeout' }));
    expect(result).toBe('Error: Timeout');
  });

  it('shows "Failed" when success is false with no message', () => {
    expect(toolResultSummary(JSON.stringify({ success: false }))).toBe('Failed');
  });

  // --- message field ---
  it('returns message field from object', () => {
    const result = toolResultSummary(JSON.stringify({ message: 'Range updated' }));
    expect(result).toBe('Range updated');
  });

  it('truncates long message field', () => {
    const msg = 'm'.repeat(150);
    const result = toolResultSummary(JSON.stringify({ message: msg }));
    expect(result).toHaveLength(103);
    expect(result.endsWith('...')).toBe(true);
  });

  // --- grid dimensions ---
  it('returns "rows x columns" for grid data', () => {
    const result = toolResultSummary(JSON.stringify({ rowCount: 10, columnCount: 5 }));
    expect(result).toBe('10 x 5');
  });

  // --- count field ---
  it('returns "N items" for count field', () => {
    expect(toolResultSummary(JSON.stringify({ count: 7 }))).toBe('7 items');
  });

  it('returns "1 item" for singular count', () => {
    expect(toolResultSummary(JSON.stringify({ count: 1 }))).toBe('1 item');
  });

  // --- single array field ---
  it('summarises single array field with its name', () => {
    const result = toolResultSummary(JSON.stringify({ sheets: ['Sheet1', 'Sheet2'] }));
    expect(result).toBe('2 sheets');
  });

  it('summarises single array field with length 0', () => {
    const result = toolResultSummary(JSON.stringify({ charts: [] }));
    expect(result).toBe('0 charts');
  });

  // --- multiple array fields → no summary ---
  it('returns empty string for object with multiple array fields', () => {
    const result = toolResultSummary(JSON.stringify({ rows: [1, 2], columns: ['A', 'B'] }));
    expect(result).toBe('');
  });

  // --- empty object ---
  it('returns empty string for empty object', () => {
    expect(toolResultSummary(JSON.stringify({}))).toBe('');
  });

  // --- non-string non-null primitives ---
  it('handles number input', () => {
    // number becomes JSON.stringify → "42" → parse → primitive → String(42)
    expect(toolResultSummary(42)).toBe('42');
  });

  it('handles boolean input', () => {
    expect(toolResultSummary(true)).toBe('true');
  });

  // --- priority: success:false takes precedence over message ---
  it('success:false with message takes precedence over message field', () => {
    const result = toolResultSummary(
      JSON.stringify({ success: false, message: 'Something broke' })
    );
    expect(result).toBe('Error: Something broke');
  });

  // --- Bug regression: circular references should not crash ---
  it('returns empty string for objects with circular references', () => {
    const obj: Record<string, unknown> = { name: 'test' };
    obj.self = obj; // circular reference
    // Before fix: JSON.stringify threw TypeError, crashing the caller
    expect(toolResultSummary(obj)).toBe('');
  });
});
