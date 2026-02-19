import { describe, it, expect, vi, beforeEach } from 'vitest';
import { webFetchTool, createRunSubagentTool, getGeneralTools } from '@/tools/general';
import type { LanguageModel, ToolSet } from 'ai';

/**
 * General tools tests — validate schemas and execution behaviour of
 * web_fetch and run_subagent without making real network/model calls.
 */

// Mock generateText so run_subagent.execute doesn't need a real model
vi.mock('ai', async () => {
  const actual = await vi.importActual<typeof import('ai')>('ai');
  return {
    ...actual,
    generateText: vi.fn().mockResolvedValue({ text: 'subagent result' }),
  };
});

interface ZodLike {
  safeParse: (data: unknown) => { success: boolean; error?: unknown };
}

function asZod(schema: unknown): ZodLike {
  return schema as ZodLike;
}

// ─── web_fetch ────────────────────────────────────────────────────────────────

describe('webFetchTool', () => {
  describe('inputSchema', () => {
    it('accepts a valid URL', () => {
      const result = asZod(webFetchTool.inputSchema).safeParse({ url: 'https://example.com' });
      expect(result.success).toBe(true);
    });

    it('accepts url with optional maxLength', () => {
      const result = asZod(webFetchTool.inputSchema).safeParse({
        url: 'https://example.com/api',
        maxLength: 5000,
      });
      expect(result.success).toBe(true);
    });

    it('rejects missing url', () => {
      const result = asZod(webFetchTool.inputSchema).safeParse({ maxLength: 100 });
      expect(result.success).toBe(false);
    });

    it('rejects non-string url', () => {
      const result = asZod(webFetchTool.inputSchema).safeParse({ url: 42 });
      expect(result.success).toBe(false);
    });

    it('rejects non-numeric maxLength', () => {
      const result = asZod(webFetchTool.inputSchema).safeParse({
        url: 'https://example.com',
        maxLength: 'a lot',
      });
      expect(result.success).toBe(false);
    });
  });

  describe('execute', () => {
    it('returns text from a successful fetch', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve('Hello, world!'),
      });
      vi.stubGlobal('fetch', mockFetch);

      const result = await webFetchTool.execute!({ url: 'https://example.com' }, {} as never);
      expect(result).toBe('Hello, world!');
      expect(mockFetch).toHaveBeenCalledWith('https://example.com');

      vi.unstubAllGlobals();
    });

    it('truncates responses longer than maxLength', async () => {
      const longText = 'x'.repeat(20_000);
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(longText),
      });
      vi.stubGlobal('fetch', mockFetch);

      const result = (await webFetchTool.execute!(
        { url: 'https://example.com', maxLength: 100 },
        {} as never
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
        webFetchTool.execute!({ url: 'https://example.com/missing' }, {} as never)
      ).rejects.toThrow('HTTP 404: Not Found');

      vi.unstubAllGlobals();
    });
  });
});

// ─── run_subagent ─────────────────────────────────────────────────────────────

describe('createRunSubagentTool', () => {
  const fakeModel = {} as LanguageModel;
  const fakeHostTools: ToolSet = {};

  describe('inputSchema', () => {
    it('accepts a task string', () => {
      const tool = createRunSubagentTool(fakeModel, fakeHostTools);
      const result = asZod(tool.inputSchema).safeParse({ task: 'Summarise the data' });
      expect(result.success).toBe(true);
    });

    it('accepts task with optional systemPrompt', () => {
      const tool = createRunSubagentTool(fakeModel, fakeHostTools);
      const result = asZod(tool.inputSchema).safeParse({
        task: 'Analyse trends',
        systemPrompt: 'You are a data analyst.',
      });
      expect(result.success).toBe(true);
    });

    it('rejects missing task', () => {
      const tool = createRunSubagentTool(fakeModel, fakeHostTools);
      const result = asZod(tool.inputSchema).safeParse({ systemPrompt: 'foo' });
      expect(result.success).toBe(false);
    });

    it('rejects non-string task', () => {
      const tool = createRunSubagentTool(fakeModel, fakeHostTools);
      const result = asZod(tool.inputSchema).safeParse({ task: 123 });
      expect(result.success).toBe(false);
    });
  });
});

// ─── getGeneralTools ──────────────────────────────────────────────────────────

describe('getGeneralTools', () => {
  let generateText: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    const ai = await import('ai');
    generateText = ai.generateText as ReturnType<typeof vi.fn>;
    generateText.mockClear();
  });

  it('returns web_fetch and run_subagent tools', () => {
    const tools = getGeneralTools({} as LanguageModel, {});
    expect(Object.keys(tools)).toContain('web_fetch');
    expect(Object.keys(tools)).toContain('run_subagent');
  });

  it('does not include run_subagent inside the subagent tools (no recursion)', async () => {
    const tools = getGeneralTools({} as LanguageModel, {});
    await tools.run_subagent.execute!({ task: 'test task' }, {} as never);

    const callOpts = generateText.mock.calls[0]?.[0] as { tools?: Record<string, unknown> };
    expect(Object.keys(callOpts.tools ?? {})).not.toContain('run_subagent');
  });

  it('provides web_fetch to the subagent', async () => {
    const tools = getGeneralTools({} as LanguageModel, {});
    await tools.run_subagent.execute!({ task: 'test task' }, {} as never);

    const callOpts = generateText.mock.calls[0]?.[0] as { tools?: Record<string, unknown> };
    expect(Object.keys(callOpts.tools ?? {})).toContain('web_fetch');
  });

  it('passes host tools through to the subagent', async () => {
    const hostTools = { some_excel_tool: webFetchTool } as ToolSet;
    const tools = getGeneralTools({} as LanguageModel, hostTools);
    await tools.run_subagent.execute!({ task: 'test task' }, {} as never);

    const callOpts = generateText.mock.calls[0]?.[0] as { tools?: Record<string, unknown> };
    expect(Object.keys(callOpts.tools ?? {})).toContain('some_excel_tool');
  });
});
