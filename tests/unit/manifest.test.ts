import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { generateManifest } from '@/tools/codegen/manifest';
import type { ToolConfig } from '@/tools/codegen/types';

describe('generateManifest', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2025-01-15T12:00:00Z'));
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  const sampleConfig: ToolConfig = {
    name: 'get_range',
    description: 'Get values from a range',
    params: {
      address: {
        type: 'string',
        required: true,
        description: 'The cell range address',
      },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional sheet name',
        default: 'Sheet1',
      },
    },
    execute: vi.fn(),
  };

  const enumConfig: ToolConfig = {
    name: 'set_format',
    description: 'Set cell format',
    params: {
      format: {
        type: 'string',
        required: true,
        description: 'Format type',
        enum: ['bold', 'italic', 'underline'],
      },
    },
    execute: vi.fn(),
  };

  it('generates a manifest from a single config array', () => {
    const manifest = generateManifest([sampleConfig]);

    expect(manifest.version).toBe('1.0.0');
    expect(manifest.generatedAt).toBe('2025-01-15T12:00:00.000Z');
    expect(manifest.tools).toHaveLength(1);
    expect(manifest.tools[0].name).toBe('get_range');
    expect(manifest.tools[0].description).toBe('Get values from a range');
  });

  it('converts params correctly with required/optional', () => {
    const manifest = generateManifest([sampleConfig]);
    const params = manifest.tools[0].params;

    expect(params.address).toEqual({
      type: 'string',
      required: true,
      description: 'The cell range address',
    });
    expect(params.sheetName).toEqual({
      type: 'string',
      required: false,
      description: 'Optional sheet name',
      default: 'Sheet1',
    });
  });

  it('includes enum values when present', () => {
    const manifest = generateManifest([enumConfig]);
    const params = manifest.tools[0].params;

    expect(params.format.enum).toEqual(['bold', 'italic', 'underline']);
  });

  it('combines multiple config arrays', () => {
    const manifest = generateManifest([sampleConfig], [enumConfig]);

    expect(manifest.tools).toHaveLength(2);
    expect(manifest.tools[0].name).toBe('get_range');
    expect(manifest.tools[1].name).toBe('set_format');
  });

  it('handles empty config arrays', () => {
    const manifest = generateManifest([], []);

    expect(manifest.tools).toHaveLength(0);
    expect(manifest.version).toBe('1.0.0');
  });
});
