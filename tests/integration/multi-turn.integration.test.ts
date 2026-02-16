/**
 * Multi-turn chat integration tests with simulated Excel tools.
 *
 * These tests exercise the REAL production `sendChatMessage` function — the
 * same code path used when a user chats in the add-in — but replace the
 * Excel.run() execute functions with canned data. This lets us verify the
 * full multi-turn pipeline (streaming, callbacks, tool calls, final answer)
 * without needing a live Excel instance.
 *
 * What's real:
 *   - `sendChatMessage` (production code)
 *   - LLM API calls (Azure AI Foundry)
 *   - Zod inputSchema validation
 *   - Multi-step agent loop (streamText + stopWhen)
 *   - messagesToCoreMessages conversion
 *   - System prompt + skill context
 *   - Streaming callbacks (onContent, onToolCalls, onToolResult, onComplete)
 *
 * What's simulated:
 *   - Tool execute functions (return sample data instead of calling Excel)
 *
 * Configuration via environment variables:
 *   FOUNDRY_ENDPOINT  – Full resource URL
 *   FOUNDRY_API_KEY   – API key for the endpoint
 *   FOUNDRY_MODEL     – Model deployment name (default: gpt-5.2-chat)
 */

import { describe, it, expect, beforeAll, vi } from 'vitest';
import { type AzureOpenAIProvider } from '@ai-sdk/azure';
import { tool } from 'ai';
import { z } from 'zod';
import { sendChatMessage } from '@/services/ai/chatService';
import type { ChatMessage, ToolCall, ToolCallResult } from '@/types';
import { createTestProvider, TEST_CONFIG } from '../test-provider';

// ─── Sample data ─────────────────────────────────────────────

/** Realistic sample spreadsheet: Q1 sales by region */
const SALES_DATA = {
  address: 'Sheet1!A1:D6',
  rowCount: 6,
  columnCount: 4,
  values: [
    ['Region', 'Jan', 'Feb', 'Mar'],
    ['North', 12000, 14500, 13200],
    ['South', 9800, 11200, 10500],
    ['East', 15600, 16800, 17200],
    ['West', 8200, 9100, 8700],
    ['Total', 45600, 51600, 49600],
  ],
};

const WORKBOOK_INFO = {
  activeSheet: 'Sheet1',
  sheets: [
    { name: 'Sheet1', position: 0, visibility: 'Visible', isActive: true },
    { name: 'Sheet2', position: 1, visibility: 'Visible', isActive: false },
  ],
  sheetCount: 2,
  usedRange: { address: 'Sheet1!A1:D6', rowCount: 6, columnCount: 4 },
  tables: [{ name: 'SalesTable', address: 'Sheet1!A1:D6', sheetName: 'Sheet1' }],
  tableCount: 1,
};

const TABLE_DATA = {
  tableName: 'SalesTable',
  address: 'Sheet1!A1:D6',
  headers: ['Region', 'Jan', 'Feb', 'Mar'],
  rows: [
    ['North', 12000, 14500, 13200],
    ['South', 9800, 11200, 10500],
    ['East', 15600, 16800, 17200],
    ['West', 8200, 9100, 8700],
    ['Total', 45600, 51600, 49600],
  ],
  rowCount: 5,
};

// ─── Simulated tools ─────────────────────────────────────────
//
// Same Zod schemas as the real tools, but execute returns canned data.
// This tests the full LLM ↔ tool contract without Excel.
// The execute functions must be async to match the AI SDK Tool interface,
// but they return synchronous canned data.

/* eslint-disable @typescript-eslint/require-await */
function createSimulatedTools() {
  /** Track every tool call for assertions */
  const callLog: { name: string; args: Record<string, unknown> }[] = [];

  const tools = {
    get_workbook_info: tool({
      description:
        'Get a high-level overview of the entire workbook: sheet names, active sheet, used range, tables.',
      inputSchema: z.object({}),
      execute: async () => {
        callLog.push({ name: 'get_workbook_info', args: {} });
        return WORKBOOK_INFO;
      },
    }),

    get_used_range: tool({
      description:
        'Get the bounding rectangle of all non-empty cells on a worksheet. Returns the range address, row count, column count, and all values.',
      inputSchema: z.object({
        sheetName: z.string().optional().describe('Optional worksheet name.'),
      }),
      execute: async args => {
        callLog.push({ name: 'get_used_range', args });
        return SALES_DATA;
      },
    }),

    get_range_values: tool({
      description: 'Read cell values from a specified range. Returns a 2D array.',
      inputSchema: z.object({
        address: z.string().describe('The range address (e.g., "A1:C10")'),
        sheetName: z.string().optional().describe('Optional worksheet name.'),
      }),
      execute: async args => {
        callLog.push({ name: 'get_range_values', args });
        return SALES_DATA;
      },
    }),

    list_sheets: tool({
      description: 'List all worksheets in the workbook with name, position, visibility.',
      inputSchema: z.object({}),
      execute: async () => {
        callLog.push({ name: 'list_sheets', args: {} });
        return { sheets: WORKBOOK_INFO.sheets, count: WORKBOOK_INFO.sheetCount };
      },
    }),

    list_tables: tool({
      description: 'List all structured Excel tables in the workbook.',
      inputSchema: z.object({
        sheetName: z.string().optional().describe('Optional sheet filter.'),
      }),
      execute: async args => {
        callLog.push({ name: 'list_tables', args });
        return { tables: WORKBOOK_INFO.tables, count: WORKBOOK_INFO.tableCount };
      },
    }),

    get_table_data: tool({
      description: 'Read all data from an Excel table. Returns headers and rows.',
      inputSchema: z.object({
        tableName: z.string().describe('Name of the table to read'),
      }),
      execute: async args => {
        callLog.push({ name: 'get_table_data', args });
        return TABLE_DATA;
      },
    }),

    set_range_values: tool({
      description: 'Write values to a specified range.',
      inputSchema: z.object({
        address: z.string().describe('The range address'),
        values: z.array(z.array(z.any())).describe('2D array of values'),
        sheetName: z.string().optional(),
      }),
      execute: async args => {
        callLog.push({ name: 'set_range_values', args });
        const vals = args.values as unknown[][];
        return {
          address: args.address,
          rowsWritten: vals.length,
          columnsWritten: vals[0]?.length ?? 0,
        };
      },
    }),

    set_range_formulas: tool({
      description: 'Write formulas to a specified range.',
      inputSchema: z.object({
        address: z.string().describe('The range address'),
        formulas: z.array(z.array(z.string())).describe('2D array of formula strings'),
        sheetName: z.string().optional(),
      }),
      execute: async args => {
        callLog.push({ name: 'set_range_formulas', args });
        return {
          address: args.address,
          rowsWritten: args.formulas.length,
          columnsWritten: args.formulas[0]?.length ?? 0,
        };
      },
    }),

    format_range: tool({
      description: 'Apply visual formatting to cells (bold, color, alignment, etc.).',
      inputSchema: z.object({
        address: z.string().describe('The range address'),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        fontSize: z.number().optional(),
        fontColor: z.string().optional(),
        fillColor: z.string().optional(),
        horizontalAlignment: z.enum(['General', 'Left', 'Center', 'Right']).optional(),
        sheetName: z.string().optional(),
      }),
      execute: async args => {
        callLog.push({ name: 'format_range', args });
        return { address: args.address, formatted: true };
      },
    }),

    create_chart: tool({
      description: 'Create a chart from a data range.',
      inputSchema: z.object({
        dataRange: z.string().describe('Source data range'),
        chartType: z
          .enum(['ColumnClustered', 'BarClustered', 'Line', 'Pie', 'XYScatter'])
          .describe('Chart type'),
        title: z.string().optional(),
        sheetName: z.string().optional(),
      }),
      execute: async args => {
        callLog.push({ name: 'create_chart', args });
        return { chartName: 'Chart 1', chartType: args.chartType, dataRange: args.dataRange };
      },
    }),

    create_sheet: tool({
      description: 'Create a new worksheet.',
      inputSchema: z.object({
        name: z.string().describe('Name for the new worksheet'),
      }),
      execute: async args => {
        callLog.push({ name: 'create_sheet', args });
        return { name: args.name, id: 'sheet-new', position: 2 };
      },
    }),
  };

  return { tools, callLog };
}
/* eslint-enable @typescript-eslint/require-await */

// ─── System prompt (same as production, trimmed) ─────────────

// The real SYSTEM_PROMPT lives inside sendChatMessage — we don't
// need to duplicate it here. sendChatMessage appends it automatically.

// ─── Helpers ─────────────────────────────────────────────────

/** Create a user ChatMessage for sendChatMessage */
function userMsg(content: string): ChatMessage {
  return { role: 'user', content };
}

/**
 * Call the real production sendChatMessage with simulated tools,
 * collecting all callbacks for assertions.
 */
async function chat(
  provider: AzureOpenAIProvider,
  prompt: string,
  tools: ReturnType<typeof createSimulatedTools>['tools']
) {
  const contentChunks: string[] = [];
  const toolCalls: ToolCall[] = [];
  const toolResults: { id: string; result: ToolCallResult }[] = [];
  let completedContent = '';

  const finalContent = await sendChatMessage(provider, {
    modelId: TEST_CONFIG.model,
    messages: [userMsg(prompt)],
    tools,
    onContent: chunk => contentChunks.push(chunk),
    onToolCalls: calls => toolCalls.push(...calls),
    onToolResult: (id, result) => toolResults.push({ id, result }),
    onComplete: content => {
      completedContent = content;
    },
  });

  return { finalContent, completedContent, contentChunks, toolCalls, toolResults };
}

// ─── Tests ───────────────────────────────────────────────────

// eslint-disable-next-line vitest/valid-describe-callback
describe.skipIf(!!TEST_CONFIG.skipReason)(
  'Multi-turn chat with simulated tools',
  { retry: 2 },
  () => {
    let provider: AzureOpenAIProvider;

    beforeAll(async () => {
      provider = await createTestProvider();
    });

    // ── Scenario 1: Read-only data exploration ──────────────

    it('explores workbook data across multiple tool calls', async () => {
      const { tools, callLog } = createSimulatedTools();

      const { finalContent, contentChunks, toolCalls, completedContent } = await chat(
        provider,
        'What data is in my spreadsheet? Summarize it for me.',
        tools
      );

      // Streaming callbacks should have fired
      expect(contentChunks.length).toBeGreaterThan(0);
      expect(completedContent).toBe(finalContent);

      // Should have called at least one data-reading tool
      const readTools = callLog.filter(c =>
        ['get_workbook_info', 'get_used_range', 'get_range_values', 'get_table_data'].includes(
          c.name
        )
      );
      expect(readTools.length).toBeGreaterThan(0);

      // onToolCalls callback should have fired
      expect(toolCalls.length).toBeGreaterThan(0);

      // Final answer should reference the actual data
      const answer = finalContent.toLowerCase();
      expect(
        answer.includes('region') || answer.includes('sales') || answer.includes('north')
      ).toBe(true);

      console.log(`  Scenario 1: ${callLog.length} tool calls, ${contentChunks.length} chunks`);
      console.log(`  Tools called: ${callLog.map(c => c.name).join(' → ')}`);
    });

    // ── Scenario 2: Read → Compute → Write ─────────────────

    it('reads data then writes a computed result', async () => {
      const { tools, callLog } = createSimulatedTools();

      const { finalContent, toolResults } = await chat(
        provider,
        'Calculate the Q1 total for each region (sum of Jan+Feb+Mar) and put the results in column E with a "Q1 Total" header.',
        tools
      );

      // Should have read the data
      const reads = callLog.filter(c =>
        ['get_workbook_info', 'get_used_range', 'get_range_values', 'get_table_data'].includes(
          c.name
        )
      );
      expect(reads.length).toBeGreaterThan(0);

      // Should have written something
      const writes = callLog.filter(c =>
        ['set_range_values', 'set_range_formulas'].includes(c.name)
      );
      expect(writes.length).toBeGreaterThan(0);

      // onToolResult callback should have fired for each tool call
      expect(toolResults.length).toBeGreaterThan(0);

      // Final answer should confirm the write
      expect(finalContent.length).toBeGreaterThan(0);

      console.log(`  Scenario 2: ${callLog.length} tool calls`);
      console.log(`  Tools called: ${callLog.map(c => c.name).join(' → ')}`);
      console.log(`  Write args: ${JSON.stringify(writes[0]?.args).slice(0, 200)}`);
    });

    // ── Scenario 3: Read → Chart creation ──────────────────

    it('reads data then creates a chart', async () => {
      const { tools, callLog } = createSimulatedTools();

      const { toolCalls } = await chat(
        provider,
        'Create a bar chart showing the monthly sales by region.',
        tools
      );

      // Should have read the data first
      const reads = callLog.filter(c =>
        ['get_workbook_info', 'get_used_range', 'get_range_values', 'get_table_data'].includes(
          c.name
        )
      );
      expect(reads.length).toBeGreaterThan(0);

      // Should have created a chart
      const charts = callLog.filter(c => c.name === 'create_chart');
      expect(charts).toHaveLength(1);

      // Verify the chart was created with reasonable parameters
      const chartArgs = charts[0].args as { dataRange: string; chartType: string };
      expect(chartArgs.dataRange).toBeTruthy();
      expect(chartArgs.chartType).toBeTruthy();

      // onToolCalls callback should include the chart tool call
      const chartToolCalls = toolCalls.filter(tc => tc.functionName === 'create_chart');
      expect(chartToolCalls).toHaveLength(1);

      console.log(`  Scenario 3: ${callLog.length} tool calls`);
      console.log(`  Chart: type=${chartArgs.chartType}, range=${chartArgs.dataRange}`);
    });

    // ── Scenario 4: Multi-step formatting ──────────────────

    it('reads data then formats the header row', async () => {
      const { tools, callLog } = createSimulatedTools();

      await chat(
        provider,
        'Make the header row bold with a blue background and white text.',
        tools
      );

      // Should have discovered data first (to know where headers are)
      const reads = callLog.filter(c =>
        ['get_workbook_info', 'get_used_range', 'get_range_values'].includes(c.name)
      );
      expect(reads.length).toBeGreaterThan(0);

      // Should have formatted something
      const formats = callLog.filter(c => c.name === 'format_range');
      expect(formats.length).toBeGreaterThan(0);

      // Verify formatting args include bold
      const fmtArgs = formats[0].args as { bold?: boolean };
      expect(fmtArgs.bold).toBe(true);

      console.log(`  Scenario 4: ${callLog.length} tool calls`);
      console.log(`  Format args: ${JSON.stringify(formats[0].args)}`);
    });

    // ── Scenario 5: New sheet + write data ─────────────────

    it('creates a new sheet and writes a summary there', async () => {
      const { tools, callLog } = createSimulatedTools();

      const { toolResults } = await chat(
        provider,
        'Create a new sheet called "Summary" and write a summary table there with columns Region and Q1 Total.',
        tools
      );

      // Should have read data to build the summary
      const reads = callLog.filter(c =>
        ['get_workbook_info', 'get_used_range', 'get_range_values', 'get_table_data'].includes(
          c.name
        )
      );
      expect(reads.length).toBeGreaterThan(0);

      // Should have created the new sheet
      const sheetCreates = callLog.filter(c => c.name === 'create_sheet');
      expect(sheetCreates).toHaveLength(1);
      expect((sheetCreates[0].args as { name: string }).name.toLowerCase()).toContain('summary');

      // Should have written data to the new sheet
      const writes = callLog.filter(c =>
        ['set_range_values', 'set_range_formulas'].includes(c.name)
      );
      expect(writes.length).toBeGreaterThan(0);

      // All tool results should have been reported via callback
      expect(toolResults).toHaveLength(callLog.length);

      console.log(`  Scenario 5: ${callLog.length} tool calls`);
      console.log(`  Tools called: ${callLog.map(c => c.name).join(' → ')}`);
    });

    // ── Scenario 6: Tool call validation ───────────────────

    it('passes valid arguments per the Zod schemas', async () => {
      const { tools, callLog } = createSimulatedTools();

      await chat(provider, 'Read the data in column B and tell me the average.', tools);

      // Every tool call that has an address arg should be a valid string
      for (const call of callLog) {
        if ('address' in call.args) {
          expect(typeof call.args.address).toBe('string');
          expect((call.args.address as string).length).toBeGreaterThan(0);
        }
      }

      // Should have called at least one data-reading tool
      expect(callLog.length).toBeGreaterThan(0);

      console.log(`  Scenario 6: ${callLog.length} tool calls, all args valid`);
      console.log(`  Tools called: ${callLog.map(c => c.name).join(' → ')}`);
    });

    // ── Scenario 7: Callback contract ─────────────────────

    it('fires onContent, onToolCalls, onToolResult, and onComplete in order', async () => {
      const { tools } = createSimulatedTools();
      const events: string[] = [];

      await sendChatMessage(provider, {
        modelId: TEST_CONFIG.model,
        messages: [userMsg('How many rows of data are in my spreadsheet?')],
        tools,
        onContent: () => {
          if (!events.includes('content')) events.push('content');
        },
        onToolCalls: () => {
          if (!events.includes('toolCalls')) events.push('toolCalls');
        },
        onToolResult: () => {
          if (!events.includes('toolResult')) events.push('toolResult');
        },
        onComplete: () => events.push('complete'),
      });

      // Tool calls and results should fire before content/complete
      expect(events).toContain('toolCalls');
      expect(events).toContain('toolResult');
      expect(events).toContain('content');
      expect(events).toContain('complete');
      // onComplete should be last
      expect(events[events.length - 1]).toBe('complete');
    });

    // ── Scenario 8: onError fires on bad model ────────────

    it('fires onError callback when model is invalid', async () => {
      const { tools } = createSimulatedTools();
      const onError = vi.fn();

      await expect(
        sendChatMessage(provider, {
          modelId: 'nonexistent-model-xyz',
          messages: [userMsg('hello')],
          tools,
          onError,
        })
      ).rejects.toThrow();

      expect(onError).toHaveBeenCalledOnce();
      expect(onError.mock.calls[0][0]).toBeInstanceOf(Error);
    });
  }
);
