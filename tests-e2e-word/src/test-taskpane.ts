/**
 * E2E Test Taskpane for Word — calls actual tool handlers against real Word.
 *
 * Unlike mock-based tests, these run the SAME code path that runs in production.
 * Each tool's handler() is called directly; the handler internally calls Word.run().
 *
 * Organisation:
 * 1. Setup: seed the document with known content (headings + paragraphs)
 * 2. Tool tests: one function per tool, verifying result strings/types
 * 3. Cleanup: close the document
 * 4. Send results to test server
 */
/* eslint-disable no-console */

import { sleep, addTestResult, TestResult, closeDocument } from './test-helpers';
import { wordConfigs } from '@/tools/word/index';
import type { WordToolConfig } from '@/tools/codegen';

/* global Office, document, Word, navigator, console, window */

// ─── Heartbeat ────────────────────────────────────────────────────

function heartbeat(msg: string): void {
  try {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', `https://localhost:4203/heartbeat?msg=${encodeURIComponent(msg)}`, true);
    xhr.send();
  } catch {
    /* ignore */
  }
}
heartbeat('word_script_loaded');

// ─── Constants ────────────────────────────────────────────────────

const port = 4203;
const testValues: TestResult[] = [];

// ─── Helpers ──────────────────────────────────────────────────────

function safeString(val: unknown): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'string') return val;
  return JSON.stringify(val);
}

// ─── Error handlers ───────────────────────────────────────────────

window.onerror = (message, source, lineno, _colno, _error) => {
  const msgStr = typeof message === 'string' ? message : '[Event]';
  console.error(`[WORD-E2E] Uncaught: ${msgStr} at ${source ?? ''}:${lineno}`);
  addTestResult(testValues, 'uncaught_error', null, 'fail', {
    error: msgStr,
    source: String(source ?? ''),
    line: lineno,
  });
  finishAndSend().catch(_err => {
    /* ignore finishAndSend error */
  });
  return false;
};

window.onunhandledrejection = (event: PromiseRejectionEvent) => {
  console.error(`[WORD-E2E] Unhandled rejection: ${String(event.reason)}`);
  addTestResult(testValues, 'unhandled_rejection', null, 'fail', { error: String(event.reason) });
  finishAndSend().catch(_err => {
    /* ignore finishAndSend error */
  });
};

// ─── Logging ──────────────────────────────────────────────────────

function log(msg: string): void {
  const el = document.getElementById('test-log');
  if (el) {
    const p = document.createElement('p');
    p.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
  }
  console.log(msg);
}

function setStatus(text: string, type: 'running' | 'success' | 'error'): void {
  const statusDiv = document.getElementById('status');
  const statusText = document.getElementById('status-text');
  if (statusDiv && statusText) {
    statusDiv.className = `status status-${type}`;
    statusText.textContent = text;
  }
}

// ─── Results ──────────────────────────────────────────────────────

let resultsSent = false;

async function sendTestResults(data: TestResult[]): Promise<void> {
  const url = `https://localhost:${port}/results`;
  await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
}

async function pingTestServer(): Promise<{ status: number }> {
  try {
    const resp = await fetch(`https://localhost:${port}/ping`);
    return { status: resp.status };
  } catch {
    return { status: 0 };
  }
}

async function finishAndSend(): Promise<void> {
  if (resultsSent) return;
  resultsSent = true;
  const passCount = testValues.filter(r => r.Type === 'pass').length;
  const failCount = testValues.filter(r => r.Type === 'fail').length;
  log(`Sending ${testValues.length} results (${passCount}P/${failCount}F)...`);
  setStatus(`Done! ${passCount} passed, ${failCount} failed`, failCount > 0 ? 'error' : 'success');
  await sendTestResults(testValues);
}

function pass(name: string): void {
  log(`  ✓ ${name}`);
  addTestResult(testValues, name, true, 'pass');
}

function fail(name: string, error: string): void {
  log(`  ✗ ${name}: ${error}`);
  addTestResult(testValues, name, null, 'fail', { error: error.substring(0, 200) });
}

// ─── Tool helpers ─────────────────────────────────────────────────

async function callTool(
  configs: readonly WordToolConfig[],
  name: string,
  args: Record<string, unknown> = {}
): Promise<unknown> {
  const config = configs.find(c => c.name === name);
  if (!config) throw new Error(`Tool config not found: ${name}`);
  let result: unknown;
  await Word.run(async context => {
    result = await config.execute(context, args);
  });
  return result;
}

async function runTool(
  configs: readonly WordToolConfig[],
  name: string,
  args: Record<string, unknown> = {},
  verify?: (result: unknown) => string | null,
  testName?: string
): Promise<unknown> {
  const label = testName ?? name;
  try {
    const result = await callTool(configs, name, args);
    // Detect tool failure result
    if (
      result &&
      typeof result === 'object' &&
      (result as Record<string, unknown>).resultType === 'failure'
    ) {
      const errMsg =
        ((result as Record<string, unknown>).error as string) ?? 'Tool returned failure';
      fail(label, errMsg);
      return result;
    }
    if (verify) {
      const err = verify(result);
      if (err) fail(label, err);
      else pass(label);
    } else {
      pass(label);
    }
    return result;
  } catch (error) {
    fail(label, String(error));
    return null;
  }
}

// ─── Helpers ──────────────────────────────────────────────────────

async function selectText(searchText: string): Promise<void> {
  await Word.run(async context => {
    const body = context.document.body;
    const results = body.search(searchText, { matchCase: false });
    results.load('items');
    await context.sync();
    if (results.items.length > 0) {
      results.items[0].select();
      await context.sync();
    }
  });
  await sleep(200);
}

async function moveCursorToEnd(): Promise<void> {
  await Word.run(async context => {
    const body = context.document.body;
    // Insert a blank paragraph so there's always a valid non-degenerate
    // insertion point (inserting after the absolute document end is invalid).
    const para = body.insertParagraph('', Word.InsertLocation.end);
    para.select();
    await context.sync();
  });
  await sleep(200);
}

// ─── Setup ────────────────────────────────────────────────────────

async function setup(): Promise<void> {
  log('── Setup ──');

  await Word.run(async context => {
    const body = context.document.body;
    body.clear();
    body.insertHtml(
      '<h1>Introduction</h1>' +
        '<p>Word E2E Test Content for automated tests. Find me for replacement.</p>' +
        '<h2>Details</h2>' +
        '<p>This section has detail text with additional information.</p>' +
        '<p>More content here for testing purposes.</p>',
      Word.InsertLocation.start
    );
    await context.sync();
  });

  await sleep(500);
  log('  Setup complete (document seeded with test content)');
}

// ─── Word Tool Tests ───────────────────────────────────────────────

async function testWordTools(): Promise<void> {
  log('── Word Tools ──');

  // 1. get_document_overview
  await runTool(wordConfigs, 'get_document_overview', {}, r => {
    const s = safeString(r);
    return s.includes('Document Overview') || s.includes('Paragraphs')
      ? null
      : `Expected "Document Overview" or "Paragraphs" in result, got: ${s.substring(0, 100)}`;
  });

  // 2. get_document_content
  await runTool(wordConfigs, 'get_document_content', {}, r => {
    const s = safeString(r);
    return s.length > 20 ? null : `Expected non-trivial HTML content, got: ${s.substring(0, 100)}`;
  });

  // 3. get_document_section
  await runTool(wordConfigs, 'get_document_section', { headingText: 'Introduction' }, r => {
    const s = safeString(r);
    return s.length > 0 ? null : 'Expected section content, got empty result';
  });

  // 4. get_selection — select known text first
  await selectText('Word E2E Test Content');
  await runTool(wordConfigs, 'get_selection', {}, r => {
    const s = safeString(r);
    return s.length > 0 ? null : 'Expected OOXML content from selection';
  });

  // 5. get_selection_text
  await selectText('Word E2E Test Content');
  await runTool(wordConfigs, 'get_selection_text', {}, r => {
    const s = safeString(r);
    return s.includes('Word E2E') || s.includes('automated')
      ? null
      : `Expected selection text to include "Word E2E", got: ${s.substring(0, 100)}`;
  });

  // 6. insert_content_at_selection
  await moveCursorToEnd();
  await runTool(
    wordConfigs,
    'insert_content_at_selection',
    { html: '<p>Inserted via E2E insert_content_at_selection test</p>', location: 'After' },
    r => {
      const s = safeString(r);
      return s.includes('inserted') || s.includes('Content')
        ? null
        : `Expected success message from insert_content_at_selection, got: ${s.substring(0, 100)}`;
    }
  );

  // 7. find_and_replace
  await runTool(
    wordConfigs,
    'find_and_replace',
    { find: 'Find me for replacement', replace: 'Replaced by E2E test' },
    r => {
      const s = safeString(r);
      return s.includes('Replaced') || s.includes('occurrence')
        ? null
        : `Expected replacement confirmation, got: ${s.substring(0, 100)}`;
    }
  );

  // 8. insert_table — grid style
  await moveCursorToEnd();
  await runTool(
    wordConfigs,
    'insert_table',
    {
      rows: 3,
      columns: 2,
      data: [
        ['Header A', 'Header B'],
        ['Row 1 Col 1', 'Row 1 Col 2'],
        ['Row 2 Col 1', 'Row 2 Col 2'],
      ],
      style: 'grid',
      hasHeaderRow: true,
    },
    r => {
      const s = safeString(r);
      return s.includes('3') && s.includes('2')
        ? null
        : `Expected table size in result, got: ${s.substring(0, 100)}`;
    }
  );

  // 9. insert_table:striped
  await moveCursorToEnd();
  await runTool(
    wordConfigs,
    'insert_table',
    {
      rows: 4,
      columns: 3,
      style: 'striped',
      hasHeaderRow: true,
    },
    r => {
      const s = safeString(r);
      return s.includes('4') && s.includes('3')
        ? null
        : `Expected table size in result, got: ${s.substring(0, 100)}`;
    },
    'insert_table:striped'
  );

  // 10. apply_style_to_selection — bold + color
  await selectText('Inserted via E2E');
  await runTool(
    wordConfigs,
    'apply_style_to_selection',
    { bold: true, fontColor: '#FF0000' },
    r => {
      const s = safeString(r);
      return s.includes('bold') || s.includes('Applied')
        ? null
        : `Expected style applied message, got: ${s.substring(0, 100)}`;
    }
  );

  // 11. apply_style_to_selection:font
  await selectText('detail text');
  await runTool(
    wordConfigs,
    'apply_style_to_selection',
    { fontSize: 18, fontName: 'Arial', italic: true },
    r => {
      const s = safeString(r);
      return s.includes('fontSize') || s.includes('Applied') || s.includes('fontName')
        ? null
        : `Expected font style message, got: ${s.substring(0, 100)}`;
    },
    'apply_style_to_selection:font'
  );

  // 12. set_document_content — last, clears document
  await runTool(
    wordConfigs,
    'set_document_content',
    { html: '<h1>E2E Final Content</h1><p>Set by set_document_content test.</p>' },
    r => {
      const s = safeString(r);
      return s.includes('replaced') || s.includes('successfully')
        ? null
        : `Expected "replaced successfully", got: ${s.substring(0, 100)}`;
    }
  );
}

// ─── LaunchEvent handler (for auto-start via manifest) ─────────────

// Register the launch event function BEFORE Office.onReady.
// This runs synchronously when the script loads in the shared runtime.
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('onWordLaunch', (event: Office.AddinCommands.Event) => {
    event.completed();
  });
}

// ─── Main ─────────────────────────────────────────────────────────

if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') {
  const diagnostic = `Office.js runtime unavailable (href=${window.location.href})`;
  console.error(`[WORD-E2E] ${diagnostic}`);
  heartbeat('office_runtime_missing');
  addTestResult(testValues, 'office_runtime_missing', null, 'fail', { error: diagnostic });
  finishAndSend().catch(_err => {
    /* ignore finishAndSend error */
  });
} else {
  void Office.onReady(async () => {
    heartbeat('word_onready_fired');
    console.log('[WORD-E2E] Office.onReady fired');

    const safetyTimer = setTimeout(() => {
      console.error('[WORD-E2E] Safety timeout (120s) — forcing result send');
      fail('safety_timeout', 'Tests did not complete within 120 seconds');
      void finishAndSend();
    }, 120000);

    try {
      await (
        Office as Record<string, unknown> & {
          addin: { showAsTaskpane: () => Promise<void> };
        }
      ).addin.showAsTaskpane();
    } catch {
      /* already visible or not supported */
    }

    const sideloadMsg = document.getElementById('sideload-msg');
    const appBody = document.getElementById('app-body');
    if (sideloadMsg) sideloadMsg.style.display = 'none';
    if (appBody) appBody.style.display = 'block';

    addTestResult(testValues, 'UserAgent', navigator.userAgent, 'info');
    log('Word Add-in loaded. Connecting to test server...');
    setStatus('Connecting...', 'running');

    try {
      const response = await pingTestServer();
      if (response.status !== 200) {
        setStatus('Test server unreachable', 'error');
        fail('test_server_connection', `Server returned status ${response.status}`);
        await finishAndSend();
        return;
      }

      log(`Test server connected on port ${port}`);
      heartbeat('word_tests_starting');
      setStatus('Running Word tests...', 'running');

      await setup();
      await testWordTools();

      clearTimeout(safetyTimer);
      await finishAndSend();
      log('Closing document...');
      await closeDocument();
    } catch (error) {
      clearTimeout(safetyTimer);
      log(`Fatal: ${String(error)}`);
      setStatus(`Error: ${String(error)}`, 'error');
      fail('fatal_error', String(error));
      try {
        await finishAndSend();
      } catch {
        /* ignore */
      }
    }
  });
}
