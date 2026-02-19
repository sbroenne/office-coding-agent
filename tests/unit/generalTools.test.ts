import { describe, it, expect, vi } from 'vitest';
import Ajv from 'ajv';
import { webFetchTool } from '@/tools/general';

/**
 * General tools tests — validate schema and execution behaviour of
 * web_fetch without making real network calls.
 */

const ajv = new Ajv({ allErrors: true });

/** Validate data against a JSON Schema */
function validate(schema: unknown, data: unknown): { success: boolean } {
  const fn = ajv.compile(schema as object);
  return { success: !!fn(data) };
}

/** Retrieve a tool's handler function, throwing if it's absent. */
function getHandler(t: unknown) {
  const tool = t as { handler?: ((args: unknown, invocation: unknown) => Promise<unknown>) | undefined };
  const { handler } = tool;
  if (!handler) throw new Error('Tool has no handler function');
  return handler;
}

// ─── web_fetch ────────────────────────────────────────────────────────────────

describe('webFetchTool', () => {
  describe('parameters', () => {
    it('accepts a valid URL', () => {
      const result = validate(webFetchTool.parameters, { url: 'https://example.com' });
      expect(result.success).toBe(true);
    });

    it('accepts url with optional maxLength', () => {
      const result = validate(webFetchTool.parameters, {
        url: 'https://example.com/api',
        maxLength: 5000,
      });
      expect(result.success).toBe(true);
    });

    it('rejects missing url', () => {
      const result = validate(webFetchTool.parameters, { maxLength: 100 });
      expect(result.success).toBe(false);
    });

    it('rejects non-string url', () => {
      const result = validate(webFetchTool.parameters, { url: 42 });
      expect(result.success).toBe(false);
    });

    it('rejects non-numeric maxLength', () => {
      const result = validate(webFetchTool.parameters, {
        url: 'https://example.com',
        maxLength: 'a lot',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('handler', () => {
    it('returns text from a successful fetch', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve('Hello, world!'),
      });
      vi.stubGlobal('fetch', mockFetch);

      const result = await getHandler(webFetchTool)({ url: 'https://example.com' }, {});
      expect(result).toBe('Hello, world!');
      expect(mockFetch).toHaveBeenCalledWith(
        `/api/fetch?url=${encodeURIComponent('https://example.com')}`
      );

      vi.unstubAllGlobals();
    });

    it('truncates responses longer than maxLength', async () => {
      const longText = 'x'.repeat(20_000);
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(longText),
      });
      vi.stubGlobal('fetch', mockFetch);

      const result = (await getHandler(webFetchTool)(
        { url: 'https://example.com', maxLength: 100 },
        {}
      )) as string;

      expect(result).toHaveLength(100 + '… [truncated]'.length);
      expect(result).toContain('… [truncated]');

      vi.unstubAllGlobals();
    });

    it('throws on HTTP error responses', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 404,
        statusText: 'Not Found',
      });
      vi.stubGlobal('fetch', mockFetch);

      await expect(
        getHandler(webFetchTool)({ url: 'https://example.com/missing' }, {})
      ).rejects.toThrow('HTTP 404: Not Found');

      vi.unstubAllGlobals();
    });
  });
});

