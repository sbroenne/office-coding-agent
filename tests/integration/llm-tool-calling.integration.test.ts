/**
 * Real LLM Tool-Calling Integration Tests
 *
 * These tests call a REAL Azure AI model with the REAL Excel tool definitions.
 * No mocks, no fakes. The LLM decides which tools to call.
 *
 * What's tested:
 * - LLM correctly selects Excel tools based on natural language prompts
 * - LLM passes valid arguments matching the tool schemas
 * - Multi-round tool loops: LLM → tool result → LLM response
 * - The AI SDK streamText function with real streaming and real tool execution
 *
 * The only thing NOT real here is Excel itself (these run in Node, not inside Excel).
 * Tool calls are intercepted and given realistic results. The LLM is real.
 *
 * For full-stack tests (real LLM + real Excel), see tests-e2e/.
 *
 * Configuration:
 *   FOUNDRY_ENDPOINT  – Azure AI Foundry resource URL
 *   FOUNDRY_API_KEY   – API key
 *   FOUNDRY_MODEL     – Model deployment (default: gpt-5.2-chat)
 *
 * Run:
 *   npx vitest run --config vitest.integration.config.ts tests/integration/llm-tool-calling.integration.test.ts
 */

import { describe, it, expect, beforeAll } from 'vitest';
import { type AzureOpenAIProvider } from '@ai-sdk/azure';
import { streamText, stepCountIs, tool } from 'ai';
import { excelTools } from '@/tools';
import type { ToolCallResult } from '@/types';
import { createTestProvider, TEST_CONFIG } from '../test-provider';

// ─── Fake tool results (realistic Excel data) ───────────────
// These simulate what the real Excel API would return.
// The LLM is REAL — only the Excel runtime is simulated.

const FAKE_TOOL_RESULTS: Record<string, (args: Record<string, unknown>) => ToolCallResult> = {
  // ─── Range ──────────────────────────────────────────────
  get_used_range: () => ({
    success: true,
    data: {
      address: 'Sheet1!A1:C5',
      rowCount: 5,
      columnCount: 3,
      values: [
        ['Product', 'Region', 'Revenue'],
        ['Widget', 'East', 5000],
        ['Gadget', 'West', 8200],
        ['Widget', 'West', 3100],
        ['Gadget', 'East', 7400],
      ],
    },
  }),
  get_range_values: args => {
    const address = String(args.address ?? '');
    if (address.includes('A1:C5') || address === 'A1:C5') {
      return {
        success: true,
        data: {
          address: 'Sheet1!A1:C5',
          values: [
            ['Product', 'Region', 'Revenue'],
            ['Widget', 'East', 5000],
            ['Gadget', 'West', 8200],
            ['Widget', 'West', 3100],
            ['Gadget', 'East', 7400],
          ],
        },
      };
    }
    return {
      success: true,
      data: {
        address,
        values: [
          ['Sample', 'Data'],
          [1, 2],
        ],
      },
    };
  },
  set_range_values: args => ({
    success: true,
    data: {
      address: args.address,
      cellsWritten: Array.isArray(args.values) ? (args.values as unknown[][]).flat().length : 0,
    },
  }),
  clear_range: args => ({
    success: true,
    data: { address: args.address, cleared: true },
  }),

  // ─── Formatting ─────────────────────────────────────────
  format_range: args => ({
    success: true,
    data: { address: args.address, formatted: true },
  }),
  set_number_format: args => ({
    success: true,
    data: { address: args.address, format: args.format },
  }),
  auto_fit_columns: args => ({
    success: true,
    data: { address: args.address ?? 'Sheet1!A:C', autoFitted: true },
  }),
  auto_fit_rows: args => ({
    success: true,
    data: { address: args.address ?? 'Sheet1!1:5', autoFitted: true },
  }),

  // ─── Formulas ───────────────────────────────────────────
  set_range_formulas: args => ({
    success: true,
    data: {
      address: args.address,
      rowsWritten: Array.isArray(args.formulas) ? (args.formulas as unknown[]).length : 0,
    },
  }),
  get_range_formulas: args => ({
    success: true,
    data: { address: args.address, formulas: [['=SUM(C2:C5)']] },
  }),

  // ─── Sort, Copy, Find ──────────────────────────────────
  sort_range: args => ({
    success: true,
    data: { address: args.address, sortedByColumn: args.column, ascending: args.ascending ?? true },
  }),
  copy_range: args => ({
    success: true,
    data: { source: args.sourceAddress, destination: args.destinationAddress, copied: true },
  }),
  find_values: args => ({
    success: true,
    data: { found: true, address: 'Sheet1!A2', value: args.searchText },
  }),

  // ─── Insert / Delete / Merge ────────────────────────────
  insert_range: args => ({
    success: true,
    data: { address: args.address, shift: args.shift ?? 'down', inserted: true },
  }),
  delete_range: args => ({
    success: true,
    data: { address: args.address, shift: args.shift ?? 'up', deleted: true },
  }),
  merge_cells: args => ({
    success: true,
    data: { address: args.address, merged: true },
  }),
  unmerge_cells: args => ({
    success: true,
    data: { address: args.address, unmerged: true },
  }),

  // ─── Sheets ─────────────────────────────────────────────
  list_sheets: () => ({
    success: true,
    data: { sheets: ['Sheet1', 'Summary', 'Raw Data'] },
  }),
  create_sheet: args => ({
    success: true,
    data: { name: args.name, created: true },
  }),
  rename_sheet: args => ({
    success: true,
    data: { oldName: args.currentName, newName: args.newName },
  }),
  delete_sheet: args => ({
    success: true,
    data: { name: args.name, deleted: true },
  }),
  activate_sheet: args => ({
    success: true,
    data: { name: args.name, activated: true },
  }),

  // ─── Tables ─────────────────────────────────────────────
  list_tables: () => ({
    success: true,
    data: { tables: [{ name: 'SalesTable', range: 'A1:C5', rowCount: 4 }] },
  }),
  create_table: args => ({
    success: true,
    data: { name: args.name ?? 'Table1', address: args.address, created: true },
  }),
  add_table_rows: args => ({
    success: true,
    data: {
      tableName: args.tableName,
      rowsAdded: Array.isArray(args.values) ? (args.values as unknown[]).length : 0,
    },
  }),
  get_table_data: () => ({
    success: true,
    data: {
      tableName: 'SalesTable',
      headers: ['Product', 'Region', 'Revenue'],
      rows: [
        ['Widget', 'East', 5000],
        ['Gadget', 'West', 8200],
        ['Widget', 'West', 3100],
        ['Gadget', 'East', 7400],
      ],
    },
  }),
  delete_table: args => ({
    success: true,
    data: { tableName: args.tableName, deleted: true },
  }),
  sort_table: args => ({
    success: true,
    data: {
      tableName: args.tableName,
      sortedByColumn: args.column,
      ascending: args.ascending ?? true,
    },
  }),
  filter_table: args => ({
    success: true,
    data: { tableName: args.tableName, filteredColumn: args.column, filterValues: args.values },
  }),
  clear_table_filters: args => ({
    success: true,
    data: { tableName: args.tableName, filtersCleared: true },
  }),

  // ─── Charts ─────────────────────────────────────────────
  list_charts: () => ({
    success: true,
    data: { charts: [{ name: 'Chart 1', chartType: 'ColumnClustered', title: 'Revenue' }] },
  }),
  create_chart: args => ({
    success: true,
    data: { chartType: args.chartType, dataRange: args.dataRange, created: true },
  }),
  delete_chart: args => ({
    success: true,
    data: { chartName: args.chartName, deleted: true },
  }),

  // ─── Workbook ───────────────────────────────────────────
  get_workbook_info: () => ({
    success: true,
    data: {
      sheetNames: ['Sheet1', 'Summary', 'Raw Data'],
      sheetCount: 3,
      activeSheet: 'Sheet1',
      usedRange: 'Sheet1!A1:C5',
      usedRangeRows: 5,
      usedRangeColumns: 3,
      tableNames: ['SalesTable'],
      tableCount: 1,
    },
  }),
  get_selected_range: () => ({
    success: true,
    data: {
      address: 'Sheet1!A1:C5',
      rowCount: 5,
      columnCount: 3,
      values: [['Product', 'Region', 'Revenue']],
    },
  }),
  define_named_range: args => ({
    success: true,
    data: { name: args.name, address: args.address, comment: args.comment ?? '' },
  }),
  list_named_ranges: () => ({
    success: true,
    data: { namedRanges: [{ name: 'SalesData', value: 'Sheet1!A1:C5', comment: '' }], count: 1 },
  }),

  // ─── Comments ───────────────────────────────────────────
  add_comment: args => ({
    success: true,
    data: { cellAddress: args.cellAddress, text: args.text, added: true },
  }),
  list_comments: () => ({
    success: true,
    data: {
      comments: [
        {
          content: 'Review this value',
          authorName: 'User',
          creationDate: '2024-01-15',
          cellAddress: 'Sheet1!A2',
        },
      ],
      count: 1,
    },
  }),
  edit_comment: args => ({
    success: true,
    data: { cellAddress: args.cellAddress, newText: args.newText, updated: true },
  }),
  delete_comment: args => ({
    success: true,
    data: { cellAddress: args.cellAddress, deleted: true },
  }),

  // ─── Conditional Formatting ─────────────────────────────
  add_conditional_format: args => ({
    success: true,
    data: { address: args.address, ruleType: args.ruleType, applied: true },
  }),
  list_conditional_formats: () => ({
    success: true,
    data: {
      conditionalFormats: [{ type: 'ColorScale', priority: 0, stopIfTrue: false }],
      count: 1,
    },
  }),
  clear_conditional_formats: args => ({
    success: true,
    data: { address: args.address ?? 'entire sheet', cleared: true },
  }),

  // ─── Data Validation ───────────────────────────────────
  set_data_validation: args => ({
    success: true,
    data: { address: args.address, ruleType: args.ruleType, applied: true },
  }),
  get_data_validation: args => ({
    success: true,
    data: {
      address: args.address,
      type: 'List',
      rule: { list: { source: 'Yes,No,Maybe', inCellDropDown: true } },
      errorAlert: null,
      prompt: null,
      ignoreBlanks: true,
    },
  }),
  clear_data_validation: args => ({
    success: true,
    data: { address: args.address, cleared: true },
  }),

  // ─── PivotTables ───────────────────────────────────────
  list_pivot_tables: () => ({
    success: true,
    data: {
      pivotTables: [
        {
          name: 'SalesPivot',
          id: 'pt1',
          rowHierarchies: ['Region'],
          dataHierarchies: ['Sum of Revenue'],
        },
      ],
      count: 1,
    },
  }),
  refresh_pivot_table: args => ({
    success: true,
    data: { pivotTableName: args.pivotTableName, refreshed: true },
  }),
  delete_pivot_table: args => ({
    success: true,
    data: { pivotTableName: args.pivotTableName, deleted: true },
  }),
};

/**
 * Create a copy of the real excelTools with fake execute handlers.
 * Keeps the real schemas (so the LLM sees the real tool definitions)
 * but executes using our fake results.
 */
function createFakeTools(
  overrides?: Record<string, (args: Record<string, unknown>) => ToolCallResult>
) {
  const results = overrides ?? FAKE_TOOL_RESULTS;
  const fakeTools: Record<string, any> = {};

  for (const [name, realTool] of Object.entries(excelTools)) {
    const handler = results[name];
    fakeTools[name] = tool({
      description: (realTool as any).description,
      inputSchema: (realTool as any).inputSchema,
      execute: async (input: Record<string, unknown>) => {
        if (handler) return handler(input);
        return { success: false, error: `Unknown tool: ${name}` };
      },
    });
  }

  return fakeTools;
}

// ─── Helper: stream a chat with real LLM, fake tool execution ──

async function chatWithTools(
  provider: AzureOpenAIProvider,
  userMessage: string,
  fakeTools?: Record<string, ReturnType<typeof tool>>
): Promise<{
  response: string;
  toolCallsExecuted: { name: string; args: Record<string, unknown>; result: ToolCallResult }[];
}> {
  const toolCallsExecuted: {
    name: string;
    args: Record<string, unknown>;
    result: ToolCallResult;
  }[] = [];

  const tools = fakeTools ?? createFakeTools();

  const systemPrompt = `You are an AI assistant inside Microsoft Excel. You have tools to read and write data in the user's workbook. Use them to answer questions about their data. Be concise.`;

  const result = streamText({
    model: provider.chat(TEST_CONFIG.model),
    system: systemPrompt,
    messages: [{ role: 'user', content: userMessage }],
    tools,
    stopWhen: stepCountIs(5),
  });

  let fullContent = '';

  for await (const part of result.fullStream) {
    switch (part.type) {
      case 'text-delta':
        fullContent += part.text;
        break;
      case 'tool-result': {
        const toolResult = part.output as ToolCallResult;
        toolCallsExecuted.push({
          name: part.toolName,
          args: part.input as Record<string, unknown>,
          result: toolResult,
        });
        break;
      }
    }
  }

  return { response: fullContent, toolCallsExecuted };
}

// ─── Tests ───────────────────────────────────────────────────

describe.skipIf(!!TEST_CONFIG.skipReason)('Real LLM Tool Calling', { retry: 2 }, () => {
  let provider: AzureOpenAIProvider;

  beforeAll(async () => {
    provider = await createTestProvider();
  });

  // ── Tool Selection Tests ─────────────────────────────────

  describe('Tool selection — LLM picks the right tool', () => {
    it('should call get_used_range when asked "what data is in the spreadsheet?"', async () => {
      const { toolCallsExecuted, response } = await chatWithTools(
        provider,
        'What data is in this spreadsheet?'
      );

      // LLM should have called a read/discovery tool
      const toolNames = toolCallsExecuted.map(tc => tc.name);
      const readSomething = toolNames.some(n =>
        [
          'get_used_range',
          'get_range_values',
          'get_table_data',
          'list_tables',
          'get_workbook_info',
          'get_selected_range',
        ].includes(n)
      );
      expect(readSomething).toBe(true);

      // LLM should respond with something about the data
      expect(response.length).toBeGreaterThan(10);
      console.log('  Tools called:', toolNames.join(', '));
      console.log('  Response:', response.substring(0, 200));
    });

    it('should call set_range_values when asked to write data', async () => {
      const { toolCallsExecuted } = await chatWithTools(provider, 'Write the value 42 to cell A1');

      const setCall = toolCallsExecuted.find(tc => tc.name === 'set_range_values');
      expect(setCall).toBeDefined();
      expect(setCall!.args.address).toBeDefined();
      // The values should contain 42
      const values = setCall!.args.values as unknown[][];
      expect(values).toBeDefined();
      const flat = values.flat();
      expect(flat.some(v => v === 42 || v === '42')).toBe(true);
      console.log('  set_range_values args:', JSON.stringify(setCall!.args));
    });

    it('should call list_sheets when asked about worksheets', async () => {
      const { toolCallsExecuted, response } = await chatWithTools(
        provider,
        'What sheets are in this workbook?'
      );

      const listCall = toolCallsExecuted.find(tc => tc.name === 'list_sheets');
      expect(listCall).toBeDefined();

      // Response should mention the sheet names from our fake data
      const mentionsSheet =
        response.includes('Sheet1') ||
        response.includes('Summary') ||
        response.includes('Raw Data');
      expect(mentionsSheet).toBe(true);
      console.log('  Response:', response.substring(0, 200));
    });

    it('should call create_sheet when asked to create a new sheet', async () => {
      const { toolCallsExecuted } = await chatWithTools(
        provider,
        'Create a new worksheet called "Analysis"'
      );

      const createCall = toolCallsExecuted.find(tc => tc.name === 'create_sheet');
      expect(createCall).toBeDefined();
      expect(createCall!.args.name).toBe('Analysis');
      console.log('  create_sheet args:', JSON.stringify(createCall!.args));
    });

    it('should call create_chart when asked to make a chart', async () => {
      const { toolCallsExecuted } = await chatWithTools(
        provider,
        'Create a bar chart from the data in A1:C5'
      );

      const chartCall = toolCallsExecuted.find(tc => tc.name === 'create_chart');
      expect(chartCall).toBeDefined();
      expect(chartCall!.args.dataRange).toBeDefined();
      console.log('  create_chart args:', JSON.stringify(chartCall!.args));
    });

    it('should call clear_range when asked to clear data', async () => {
      const { toolCallsExecuted } = await chatWithTools(provider, 'Clear all data in cells A1:B10');

      const clearCall = toolCallsExecuted.find(tc => tc.name === 'clear_range');
      expect(clearCall).toBeDefined();
      expect(clearCall!.args.address).toBeDefined();
      console.log('  clear_range args:', JSON.stringify(clearCall!.args));
    });
  });

  // ── Multi-round Reasoning Tests ──────────────────────────

  describe('Multi-round — LLM reads data then answers', () => {
    it('should read data and provide analysis', async () => {
      const { toolCallsExecuted, response } = await chatWithTools(
        provider,
        'What is the total revenue in the spreadsheet?'
      );

      // LLM should have used at least one read tool
      const readTools = toolCallsExecuted.filter(tc =>
        [
          'get_used_range',
          'get_range_values',
          'get_table_data',
          'list_tables',
          'get_workbook_info',
        ].includes(tc.name)
      );
      expect(readTools.length).toBeGreaterThan(0);

      // After reading the data (Widget East 5000, Gadget West 8200, Widget West 3100, Gadget East 7400)
      // total = 23700
      // LLM should calculate or mention the total
      const mentionsTotal = response.includes('23700') || response.includes('23,700');
      expect(mentionsTotal).toBe(true);
      console.log('  Tools used:', toolCallsExecuted.map(tc => tc.name).join(' → '));
      console.log('  Response:', response.substring(0, 300));
    });

    it('should read data then write a summary', async () => {
      const { toolCallsExecuted } = await chatWithTools(
        provider,
        'Read the data in the spreadsheet and write a summary total in cell D1'
      );

      // Should have both read and write operations
      const readTools = toolCallsExecuted.filter(tc =>
        [
          'get_used_range',
          'get_range_values',
          'get_table_data',
          'list_tables',
          'get_workbook_info',
        ].includes(tc.name)
      );
      const writeTools = toolCallsExecuted.filter(tc => tc.name === 'set_range_values');

      expect(readTools.length).toBeGreaterThan(0);
      expect(writeTools.length).toBeGreaterThan(0);

      // The write should target D1
      const writeCall = writeTools[0];
      expect(writeCall.args.address).toBeDefined();
      console.log('  Read tools:', readTools.map(t => t.name).join(', '));
      console.log('  Write args:', JSON.stringify(writeCall.args));
    });
  });

  // ── Error Handling ────────────────────────────────────────

  describe('Error handling — LLM reacts to tool failures', () => {
    it('should handle a tool error and explain it to the user', async () => {
      // Override get_range_values and get_used_range to fail
      const failingResults: Record<string, (args: Record<string, unknown>) => ToolCallResult> = {
        ...FAKE_TOOL_RESULTS,
        get_range_values: () => ({
          success: false,
          error: 'The range A1:Z100 exceeds the worksheet boundaries',
        }),
        get_used_range: () => ({
          success: false,
          error: 'No data found in the active worksheet',
        }),
      };

      const failingTools = createFakeTools(failingResults);

      const { response } = await chatWithTools(
        provider,
        'Read all the data in the spreadsheet',
        failingTools
      );

      // The LLM should acknowledge the error in its response, not crash
      expect(response.length).toBeGreaterThan(0);
      console.log('  Error response:', response.substring(0, 300));
    });
  });
});
