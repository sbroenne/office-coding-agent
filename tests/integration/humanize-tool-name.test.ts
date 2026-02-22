import { describe, it, expect } from 'vitest';
import { humanizeToolName } from '@/utils/humanizeToolName';

describe('humanizeToolName', () => {
  it.each([
    ['get_workbook_info', 'Get workbook info'],
    ['set_range_values', 'Set range values'],
    ['create_table', 'Create table'],
    ['add_conditional_format', 'Add conditional format'],
    ['list_named_ranges', 'List named ranges'],
    ['delete_sheet', 'Delete sheet'],
  ])('converts %s → %s', (input, expected) => {
    expect(humanizeToolName(input)).toBe(expected);
  });

  it('returns the original for empty string', () => {
    expect(humanizeToolName('')).toBe('');
  });

  it('handles single word', () => {
    expect(humanizeToolName('read')).toBe('Read');
  });

  it('handles already-humanized input', () => {
    expect(humanizeToolName('Get info')).toBe('Get info');
  });
});
