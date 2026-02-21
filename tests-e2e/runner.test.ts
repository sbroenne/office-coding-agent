/**
 * Excel AI E2E Test Runner
 *
 * Orchestrates E2E testing by:
 * 1. Starting a custom test server to receive results from Excel (via POST body)
 * 2. Building and serving the test add-in via Vite (port 3001)
 * 3. Sideloading into Excel Desktop
 * 4. Collecting and validating test results
 *
 * Tests are auto-generated from the tool manifest — every tool gets an it() block.
 * test-taskpane.ts (in Excel) calls the actual tool execute() against real Excel.run(),
 * reports results to this runner, and we validate them here.
 *
 * NOTE: We use a custom Express server instead of office-addin-test-server because
 * the latter puts all results in the URL query string (req.query.data), which hits
 * Node.js's ~8KB URL length limit when sending 83+ test results. Our server reads
 * from the POST body instead.
 */

import * as assert from 'assert';
import { AppType, startDebugging, stopDebugging } from 'office-addin-debugging';
import { toOfficeApp } from 'office-addin-manifest';
import { closeDesktopApplication } from './src/node-helpers';
import * as path from 'path';
import * as https from 'https';
import express from 'express';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { e2eContext, TestResult } from './test-context';

/* global process, describe, before, it, after, console */

const host = 'excel';
const manifestPath = path.resolve(`${process.cwd()}/tests-e2e/test-manifest.xml`);
const port = 4201;

async function pingServer(serverPort: number): Promise<{ status: number }> {
  return await new Promise((resolve, reject) => {
    const request = https.get(
      `https://localhost:${serverPort}/ping`,
      { rejectUnauthorized: false },
      response => {
        resolve({ status: response.statusCode ?? 0 });
      }
    );
    request.on('error', reject);
    request.end();
  });
}

/**
 * Custom test server that reads POST body (not URL query params).
 */
class CustomTestServer {
  private app = express();
  private server: https.Server | null = null;
  private resolveResults: ((results: TestResult[]) => void) | null = null;
  private resultsPromise: Promise<TestResult[]>;
  private serverPort: number;

  constructor(serverPort: number) {
    this.serverPort = serverPort;
    this.resultsPromise = new Promise(resolve => {
      this.resolveResults = resolve;
    });
  }

  async start(): Promise<void> {
    const options = await getHttpsServerOptions();
    // CORS headers inline (avoid @types/cors dependency)
    this.app.use((_req: express.Request, res: express.Response, next: express.NextFunction) => {
      res.header('Access-Control-Allow-Origin', '*');
      res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
      res.header('Access-Control-Allow-Headers', 'Content-Type');
      if (_req.method === 'OPTIONS') {
        res.sendStatus(200);
        return;
      }
      next();
    });
    this.app.use(express.json({ limit: '5mb' }));

    this.app.get('/ping', (_req: express.Request, res: express.Response) => {
      res.send(process.platform === 'win32' ? 'Windows' : process.platform);
    });

    this.app.get('/heartbeat', (req: express.Request, res: express.Response) => {
      const msg = req.query.msg || '(no message)';
      console.log(`[HEARTBEAT] ${msg}`);
      res.send('ok');
    });

    this.app.post('/results', (req: express.Request, res: express.Response) => {
      res.send('200');
      const data = req.body as TestResult[];
      if (data && Array.isArray(data)) {
        console.log(`Received ${data.length} results via POST body`);
        this.resolveResults?.(data);
      } else {
        console.error('Invalid results format received');
      }
    });

    this.server = https.createServer(options, this.app);
    await new Promise<void>((resolve, reject) => {
      this.server!.listen(this.serverPort, () => {
        resolve();
      });
      this.server!.on('error', reject);
    });
  }

  async getResults(): Promise<TestResult[]> {
    return this.resultsPromise;
  }

  async stop(): Promise<void> {
    if (this.server) {
      this.server.close();
    }
  }
}

const testServer = new CustomTestServer(port);

// ─── Tool name groups (matches test-taskpane.ts suites) ─────────

const rangeTools = [
  'get_range_values',
  'set_range_values',
  'get_used_range',
  'get_used_range:maxRows',
  'clear_range',
  'format_range',
  'set_number_format',
  'auto_fit_columns',
  'auto_fit_rows',
  'set_range_formulas',
  'get_range_formulas',
  'sort_range',
  'copy_range',
  'find_values',
  'insert_range',
  'delete_range',
  'merge_cells',
  'unmerge_cells',
  'replace_values',
  'remove_duplicates',
  'set_hyperlink',
  'toggle_row_column_visibility',
  'group_rows_columns',
  'ungroup_rows_columns',
  'set_cell_borders',
  // Variants
  'format_range:underline',
  'format_range:align_left',
  'format_range:align_right',
  'sort_range:descending',
  'sort_range:no_headers',
  'find_values:not_found',
  'find_values:match_case',
  'find_values:match_entire_cell',
  'insert_range:shift_right',
  'delete_range:shift_left',
  'merge_cells:across',
  'replace_values:no_address',
  'auto_fit_columns:no_address',
  'auto_fit_rows:no_address',
  'set_hyperlink:tooltip',
  'set_hyperlink:remove',
  'toggle_visibility:rows',
  'set_cell_borders:edge_all',
  'set_cell_borders:medium',
  'set_cell_borders:dashed',
  'set_cell_borders:double',
  'get_used_range:no_truncation',
  'format_range:align_general',
  'format_range:align_justify',
  'set_cell_borders:thick',
  'set_cell_borders:dotted',
  'set_cell_borders:dashdot',
  'auto_fill_range',
  'flash_fill_range',
  'get_special_cells',
  'get_range_precedents',
  'get_range_dependents',
  'recalculate_range',
  'get_tables_for_range',
];

const tableTools = [
  'create_table',
  'list_tables',
  'add_table_rows',
  'get_table_data',
  'sort_table',
  'filter_table',
  'clear_table_filters',
  'add_table_column',
  'delete_table_column',
  'convert_table_to_range',
  'delete_table',
  // Variants
  'create_table:no_headers',
  'sort_table:descending',
  'add_table_column:auto_name',
  'list_tables:workbook_wide',
  'resize_table',
  'set_table_style',
  'set_table_header_totals_visibility',
  'reapply_table_filters',
];

const chartTools = [
  'create_chart',
  'list_charts',
  'set_chart_title',
  'set_chart_type',
  'set_chart_data_source',
  'delete_chart',
  // Variants
  'create_chart:line',
  'create_chart:bar_clustered',
  'create_chart:no_title',
  'set_chart_type:bar_stacked',
  'set_chart_type:doughnut',
  'set_chart_type:scatter',
  'create_chart:column_stacked',
  'create_chart:line_markers',
  'set_chart_position',
  'set_chart_legend_visibility',
  'set_chart_axis_title',
  'set_chart_axis_visibility',
  'set_chart_series_filtered',
];

const sheetTools = [
  'list_sheets',
  'create_sheet',
  'activate_sheet',
  'rename_sheet',
  'copy_sheet',
  'move_sheet',
  'freeze_panes',
  'protect_sheet',
  'unprotect_sheet',
  'set_sheet_visibility',
  'set_page_layout',
  'delete_sheet',
  // Variants
  'protect_sheet:password',
  'unprotect_sheet:password',
  'set_visibility:very_hidden',
  'set_visibility:tab_color_only',
  'set_visibility:clear_tab_color',
  'set_page_layout:portrait_a4',
  'copy_sheet:auto_name',
  'set_page_layout:letter',
  'set_page_layout:legal',
  'set_page_layout:tabloid',
  'set_sheet_gridlines',
  'set_sheet_headings',
  'recalculate_sheet',
];

const workbookTools = [
  'get_workbook_info',
  'get_selected_range',
  'define_named_range',
  'list_named_ranges',
  'recalculate_workbook',
  // Variants
  'define_named_range:no_comment',
  'recalculate:recalculate',
  'recalculate:default',
  'save_workbook',
  'get_workbook_properties',
  'set_workbook_properties',
  'get_workbook_protection',
  'protect_workbook',
  'unprotect_workbook',
  'refresh_data_connections',
  'list_queries',
  'get_query',
  'get_query_count',
];

const commentTools = [
  'add_comment',
  'list_comments',
  'edit_comment',
  'delete_comment',
  // Variants
  'add_comment:no_sheet',
  'list_comments:no_sheet',
  'edit_comment:no_sheet',
  'delete_comment:no_sheet',
];

const conditionalFormatTools = [
  'add_color_scale',
  'add_data_bar',
  'add_cell_value_format',
  'add_top_bottom_format',
  'add_contains_text_format',
  'add_custom_format',
  'list_conditional_formats',
  'clear_conditional_formats',
  // Variants
  'add_color_scale:3_color',
  'add_color_scale:defaults',
  'add_data_bar:default',
  'add_cell_value_format:between',
  'add_cell_value_format:less_than',
  'add_cell_value_format:equal_to',
  'add_top_bottom:bottom_items',
  'add_top_bottom:top_percent',
  'add_top_bottom:defaults_fontcolor',
  'add_contains_text:defaults',
  'add_contains_text:fill_color',
  'add_custom_format:font_color',
  'clear_cf:whole_sheet',
  'add_cell_value_format:not_equal',
  'add_cell_value_format:gte',
  'add_cell_value_format:lte',
  'add_cell_value_format:not_between',
  'add_top_bottom:bottom_percent',
  'clear_cf:final_cleanup',
];

const dataValidationTools = [
  'set_list_validation',
  'set_number_validation',
  'set_date_validation',
  'set_text_length_validation',
  'set_custom_validation',
  'get_data_validation',
  'clear_data_validation',
  // Variants
  'set_list_validation:alerts',
  'set_number_validation:decimal_eq',
  'set_number_validation:greater_than',
  'set_number_validation:lte',
  'set_date_validation:between',
  'set_text_length:between',
  'set_text_length:gte',
  'set_custom_validation:alerts',
  'set_number_validation:not_equal',
  'set_number_validation:not_between',
  'set_date_validation:less_than',
  'set_date_validation:not_between',
  'set_text_length:equal_to',
  'set_text_length:not_between',
];

const pivotTableTools = [
  'create_pivot_table',
  'list_pivot_tables',
  'refresh_pivot_table',
  'get_pivot_table_source_info',
  'set_pivot_table_options',
  'add_pivot_field',
  'set_pivot_layout',
  'get_pivot_field_filters',
  'apply_pivot_label_filter',
  'sort_pivot_field_labels',
  'apply_pivot_manual_filter',
  'sort_pivot_field_values',
  'set_pivot_field_show_all_items',
  'clear_pivot_field_filters',
  'remove_pivot_field',
  'delete_pivot_table',
  // Variants
  'create_pivot_table:multi_fields',
  'set_pivot_table_options:disable_flags',
  'add_pivot_field:filter',
  'set_pivot_layout:outline_off',
  'apply_pivot_label_filter:between',
  'sort_pivot_field_labels:ascending',
  'apply_pivot_manual_filter:multi',
  'sort_pivot_field_values:ascending',
  'set_pivot_field_show_all_items:true',
  'clear_pivot_field_filters:all',
  'remove_pivot_field:filter',
  // Note: add_pivot_field:data and remove_pivot_field:data skipped (all fields assigned)
  'delete_pivot_table:variant',
];

const settingsTests = [
  'officeruntime_storage_available',
  'officeruntime_storage_roundtrip',
  'officeruntime_storage_missing_key',
  'officeruntime_storage_remove',
];

const aiTests = [
  'ai_roundtrip_read',
  'ai_roundtrip_response',
  'ai_roundtrip_write',
  'ai_roundtrip_verify',
];

// ─── Helper: generate it() blocks from a name list ──────────────

function assertToolResult(name: string): void {
  const result = e2eContext.getResult(name);
  assert.ok(result, `No result received for "${name}" — test may not have run in Excel`);
  assert.strictEqual(
    result.Type,
    'pass',
    `${name}: ${(result.Metadata?.error as string) || 'Test failed'}`
  );
}

// ─── Test Suite ─────────────────────────────────────────────────

describe('Excel AI E2E Tests', function () {
  this.timeout(0); // Tests involve sideloading Excel — no timeout

  before(`Setup: start test server and sideload ${host}`, async () => {
    console.log('Setting up test environment...');

    // Start custom test server (reads POST body, no URL length limits)
    await testServer.start();
    const serverResponse = await pingServer(port);
    assert.strictEqual(
      (serverResponse as { status: number }).status,
      200,
      'Test server should respond'
    );
    console.log(`Test server started on port ${port}`);

    // Build test add-in and sideload into Excel
    const devServerCmd = 'npx vite --config ./tests-e2e/vite.config.ts';
    const options = {
      appType: AppType.Desktop,
      app: toOfficeApp(host),
      devServerCommandLine: devServerCmd,
      devServerPort: 3001,
      enableDebugging: false,
    };

    console.log('Starting dev server and sideloading add-in...');
    await startDebugging(manifestPath, options);
    console.log('Add-in sideloaded');

    // Wait for results from the add-in running inside Excel
    console.log('Waiting for test results from add-in...');
    const results = await testServer.getResults();
    e2eContext.setResults(results);
    console.log(`Received ${results.length} test results`);

    const userAgent = results.find(v => v.Name === 'UserAgent');
    if (userAgent) {
      console.log(`User Agent: ${userAgent.Value}`);
    }
  });

  after('Teardown: stop server, close Excel, unregister add-in', async () => {
    console.log('Tearing down...');

    await testServer.stop();

    console.log('Closing Excel...');
    try {
      await closeDesktopApplication();
    } catch (error) {
      console.log(`Note: Excel may already be closed: ${error}`);
    }

    console.log('Stopping debugging...');
    await stopDebugging(manifestPath);
    console.log('Teardown complete');
  });

  // ─── Result Collection ────────────────────────────────────────

  describe('Result Collection', () => {
    it('should receive test results from Excel', () => {
      const results = e2eContext.getResults();
      assert.ok(results.length >= 2, `Expected at least 2 results, got ${results.length}`);
    });
  });

  // ─── Tool Tests — one it() per tool ───────────────────────────

  describe('Range Tools (59)', () => {
    for (const name of rangeTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Table Tools (19)', () => {
    for (const name of tableTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Chart Tools (19)', () => {
    for (const name of chartTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Sheet Tools (25)', () => {
    for (const name of sheetTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Workbook Tools (18)', () => {
    for (const name of workbookTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Comment Tools (8)', () => {
    for (const name of commentTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Conditional Format Tools (27)', () => {
    for (const name of conditionalFormatTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Data Validation Tools (21)', () => {
    for (const name of dataValidationTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  describe('Pivot Table Tools (28)', () => {
    for (const name of pivotTableTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  // ─── Settings Persistence (OfficeRuntime.storage) ─────────────

  describe('Settings Persistence', () => {
    for (const name of settingsTests) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  // ─── AI Round-Trip (Real LLM + Real Excel) ────────────────────

  describe('AI Round-Trip', () => {
    for (const name of aiTests) {
      it(name, function () {
        const result = e2eContext.getResult(name);
        if (!result) {
          const skipped = e2eContext.getResult('ai_roundtrip_skipped');
          if (skipped) this.skip();
          else assert.fail(`No result for "${name}" — AI test may not have run`);
          return;
        }
        assert.strictEqual(
          result.Type,
          'pass',
          `${name}: ${(result.Metadata?.error as string) || 'Test failed'}`
        );
      });
    }
  });

  // ─── Summary ──────────────────────────────────────────────────

  describe('Summary', () => {
    it('all in-Excel tests should pass', () => {
      const failed = e2eContext.getFailedTests();
      if (failed.length > 0) {
        const names = failed.map(f => `${f.Name}: ${f.Metadata?.error || 'unknown'}`).join('\n  ');
        assert.fail(`${failed.length} test(s) failed in Excel:\n  ${names}`);
      }
    });
  });
});
