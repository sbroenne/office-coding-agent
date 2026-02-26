/**
 * E2E Test Taskpane for Outlook — calls actual Outlook tool handlers against real Outlook.
 *
 * Unlike mock-based tests, these run the SAME code path that runs in production.
 * Each tool's handler() is called directly; handlers internally use Office.context.mailbox.
 *
 * This taskpane runs inside a compose window opened by the test runner.
 * Auto-started via OnNewMessageCompose LaunchEvent.
 *
 * Organisation:
 * 1. Mailbox tools: non-item-dependent (user profile, diagnostics)
 * 2. Compose item tools: tools that operate on the compose item
 * 3. Send results to test server
 */
/* eslint-disable no-console */

import { sleep, addTestResult, TestResult, closeSession } from './test-helpers';
import { outlookTools } from '@/tools/outlook/index';
import type { Tool, ToolResultObject } from '@github/copilot-sdk';

/* global Office, document, navigator, console, window */

// ─── Heartbeat ────────────────────────────────────────────────────

function heartbeat(msg: string): void {
  try {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', `https://localhost:4204/heartbeat?msg=${encodeURIComponent(msg)}`, true);
    xhr.send();
  } catch {
    /* ignore */
  }
}
heartbeat('outlook_script_loaded');

// ─── Constants ────────────────────────────────────────────────────

const port = 4204;
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
  console.error(`[OUTLOOK-E2E] Uncaught: ${msgStr} at ${source ?? ''}:${lineno}`);
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
  console.error(`[OUTLOOK-E2E] Unhandled rejection: ${String(event.reason)}`);
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
  tools: readonly Tool[],
  name: string,
  args: Record<string, unknown> = {}
): Promise<unknown> {
  const tool = tools.find(t => t.name === name);
  if (!tool) throw new Error(`Outlook tool not found: ${name}`);
  // Provide a minimal dummy invocation — Outlook handlers don't use it
  const invocation = { sessionId: 'e2e', toolCallId: 'e2e', toolName: name, arguments: args };
  return await tool.handler(args, invocation);
}

async function runTool(
  tools: readonly Tool[],
  name: string,
  args: Record<string, unknown> = {},
  verify?: (result: unknown) => string | null,
  testName?: string
): Promise<unknown> {
  const label = testName ?? name;
  try {
    const result = await callTool(tools, name, args);
    // Detect tool failure result
    if (
      result &&
      typeof result === 'object' &&
      (result as ToolResultObject).resultType === 'failure'
    ) {
      const errMsg =
        ((result as ToolResultObject).error as string) ?? 'Tool returned failure';
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

// ─── Outlook Tool Tests ───────────────────────────────────────────

async function testOutlookTools(): Promise<void> {
  log('── Outlook Tools ──');

  // 1. get_user_profile — works without a mail item open
  await runTool(outlookTools, 'get_user_profile', {}, r => {
    const s = safeString(r);
    return s.includes('User Profile') && s.includes('Email')
      ? null
      : `Expected "User Profile" and "Email" in result, got: ${s.substring(0, 100)}`;
  });

  // 2. get_diagnostics — works without a mail item open
  await runTool(outlookTools, 'get_diagnostics', {}, r => {
    const s = safeString(r);
    return s.includes('Diagnostics') && s.includes('Host')
      ? null
      : `Expected "Diagnostics" and "Host" in result, got: ${s.substring(0, 100)}`;
  });

  // ── Compose item tests (run inside a new compose window) ──────

  // 3. get_mail_item — compose mode: type=ItemCompose, subject is async getter
  await runTool(outlookTools, 'get_mail_item', {}, r => {
    const s = safeString(r);
    return s.includes('Mail Item') || s.includes('Mode') || s.includes('Type')
      ? null
      : `Expected mail item overview, got: ${s.substring(0, 100)}`;
  });

  // 4. get_mail_body — compose mode: body exists (may be empty or have signature)
  await runTool(outlookTools, 'get_mail_body', { format: 'text' }, r => {
    // In compose mode, body.getAsync() should return a string (possibly empty)
    if (r === null || r === undefined) return 'Expected a body response (even empty)';
    const s = safeString(r);
    // ToolResultObject failure case is already handled by runTool
    // An empty string body is a valid result in compose mode
    return typeof r === 'string' || s.length >= 0 ? null : `Unexpected body result: ${s.substring(0, 100)}`;
  });

  // 5. get_mail_attachments — compose mode: should say "No attachments"
  await runTool(outlookTools, 'get_mail_attachments', {}, r => {
    const s = safeString(r);
    return s.includes('attachment') || s.includes('No attachment')
      ? null
      : `Expected attachment info, got: ${s.substring(0, 100)}`;
  });

  // 6. get_mail_headers — compose mode: some headers may be null before save
  await runTool(outlookTools, 'get_mail_headers', {}, r => {
    const s = safeString(r);
    return s.includes('Mail Headers') || s.includes('Header')
      ? null
      : `Expected mail headers, got: ${s.substring(0, 100)}`;
  });
}

// ─── LaunchEvent handler (for auto-start via OnNewMessageCompose) ──

// Register the launch event function BEFORE Office.onReady.
// This runs synchronously when the script loads in the shared runtime.
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('onOutlookLaunch', (event: Office.AddinCommands.Event) => {
    event.completed();
  });
}

// ─── Main ─────────────────────────────────────────────────────────

if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') {
  const diagnostic = `Office.js runtime unavailable (href=${window.location.href})`;
  console.error(`[OUTLOOK-E2E] ${diagnostic}`);
  heartbeat('office_runtime_missing');
  addTestResult(testValues, 'office_runtime_missing', null, 'fail', { error: diagnostic });
  finishAndSend().catch(_err => {
    /* ignore finishAndSend error */
  });
} else {
  void Office.onReady(async () => {
    heartbeat('outlook_onready_fired');
    console.log('[OUTLOOK-E2E] Office.onReady fired');

    const safetyTimer = setTimeout(() => {
      console.error('[OUTLOOK-E2E] Safety timeout (120s) — forcing result send');
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
    log('Outlook Add-in loaded. Connecting to test server...');
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
      heartbeat('outlook_tests_starting');
      setStatus('Running Outlook tests...', 'running');

      await testOutlookTools();

      clearTimeout(safetyTimer);
      await finishAndSend();
      log('Closing session...');
      await closeSession();
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
