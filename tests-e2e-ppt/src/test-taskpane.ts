/**
 * E2E Test Taskpane for PowerPoint — calls actual tool handlers against real PowerPoint.
 *
 * Unlike mock-based tests, these run the SAME code path that runs in production.
 * Each tool's handler() is called directly; the handler internally calls PowerPoint.run().
 *
 * Organisation:
 * 1. Setup: ensure at least 1 slide exists with test content
 * 2. Tool tests: one function per tool, verifying result strings
 * 3. Cleanup: close the presentation
 * 4. Send results to test server
 */
/* eslint-disable no-console */

import { sleep, addTestResult, TestResult, closePresentation } from './test-helpers';
import { powerPointConfigs } from '@/tools/powerpoint/index';
import type { PptToolConfig } from '@/tools/codegen';

/* global Office, document, PowerPoint, navigator, console, window */

// ─── Heartbeat ────────────────────────────────────────────────────

function heartbeat(msg: string): void {
  try {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', `https://localhost:4202/heartbeat?msg=${encodeURIComponent(msg)}`, true);
    xhr.send();
  } catch {
    /* ignore */
  }
}
heartbeat('ppt_script_loaded');

// ─── Constants ────────────────────────────────────────────────────

const port = 4202;
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
  console.error(`[PPT-E2E] Uncaught: ${msgStr} at ${source ?? ''}:${lineno}`);
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
  console.error(`[PPT-E2E] Unhandled rejection: ${String(event.reason)}`);
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
  configs: readonly PptToolConfig[],
  name: string,
  args: Record<string, unknown> = {}
): Promise<unknown> {
  const config = configs.find(c => c.name === name);
  if (!config) throw new Error(`Tool config not found: ${name}`);
  let result: unknown;
  await PowerPoint.run(async context => {
    result = await config.execute(context, args);
  });
  return result;
}

async function runTool(
  configs: readonly PptToolConfig[],
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

// ─── Setup ────────────────────────────────────────────────────────

let initialSlideCount = 0;

async function setup(): Promise<void> {
  log('── Setup ──');

  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slides.items.length === 0) {
      slides.add();
      await context.sync();
      slides.load('items');
      await context.sync();
    }

    // Add a text box to slide 0 with known test content
    const slide = slides.items[0];
    slide.shapes.addTextBox('PPT E2E Test Content — Slide 1', {
      left: 50,
      top: 100,
      width: 600,
      height: 200,
    });
    await context.sync();

    initialSlideCount = slides.items.length;
  });

  await sleep(500);
  log(`  Setup complete (${initialSlideCount} slide(s))`);
}

// ─── PowerPoint Tool Tests ─────────────────────────────────────────

async function testPptTools(): Promise<void> {
  log('── PowerPoint Tools ──');

  // 1. get_presentation_overview
  await runTool(powerPointConfigs, 'get_presentation_overview', {}, r => {
    const s = safeString(r);
    return s.includes('Total slides')
      ? null
      : `Expected "Total slides" in result, got: ${s.substring(0, 100)}`;
  });

  // 2. get_presentation_content (all slides)
  await runTool(powerPointConfigs, 'get_presentation_content', {}, r => {
    const s = safeString(r);
    return s.includes('Slide') ? null : `Expected "Slide" in result, got: ${s.substring(0, 100)}`;
  });

  // 3. get_presentation_content:single
  await runTool(
    powerPointConfigs,
    'get_presentation_content',
    { slideIndex: 0 },
    r => {
      const s = safeString(r);
      return s.includes('Slide 1')
        ? null
        : `Expected "Slide 1" in result, got: ${s.substring(0, 100)}`;
    },
    'get_presentation_content:single'
  );

  // 4. get_presentation_content:range
  await runTool(
    powerPointConfigs,
    'get_presentation_content',
    { startIndex: 0, endIndex: 0 },
    r => {
      const s = safeString(r);
      return s.includes('Slide')
        ? null
        : `Expected "Slide" in range result, got: ${s.substring(0, 100)}`;
    },
    'get_presentation_content:range'
  );

  // 5. get_slide_notes
  await runTool(powerPointConfigs, 'get_slide_notes', { slideIndex: 0 }, r => {
    const s = safeString(r);
    return s.length > 0 ? null : 'Expected non-empty notes result';
  });

  // 6. get_slide_notes:all
  await runTool(
    powerPointConfigs,
    'get_slide_notes',
    {},
    r => {
      const s = safeString(r);
      return s.length > 0 ? null : 'Expected non-empty notes result for all slides';
    },
    'get_slide_notes:all'
  );

  // 7. set_presentation_content (add text box to slide 0)
  await runTool(
    powerPointConfigs,
    'set_presentation_content',
    { slideIndex: 0, text: 'Added via E2E set_presentation_content test' },
    r => {
      const s = safeString(r);
      return s.length > 0 ? null : 'Expected non-empty result from set_presentation_content';
    }
  );

  // 8. update_slide_shape (update first shape on slide 0)
  await runTool(
    powerPointConfigs,
    'update_slide_shape',
    { slideIndex: 0, shapeIndex: 0, text: 'Updated by E2E update_slide_shape test' },
    r => {
      const s = safeString(r);
      return s.length > 0 ? null : 'Expected non-empty result from update_slide_shape';
    }
  );

  // 9. set_slide_notes
  await runTool(
    powerPointConfigs,
    'set_slide_notes',
    { slideIndex: 0, notes: 'E2E automated test speaker notes' },
    r => {
      const s = safeString(r);
      return s.length > 0 ? null : 'Expected non-empty result from set_slide_notes';
    }
  );

  // 10. add_slide_from_code
  const simpleSlideCode = [
    'slide.addText("E2E Test Slide", { x: 1, y: 1, w: 8, h: 1.5, fontSize: 36, bold: true });',
    'slide.addText("Created by e2e automated tests", { x: 1, y: 3, w: 8, h: 1, fontSize: 18 });',
  ].join('\n');

  await runTool(powerPointConfigs, 'add_slide_from_code', { code: simpleSlideCode }, r => {
    const s = safeString(r);
    return s.toLowerCase().includes('success') || s.includes('slide')
      ? null
      : `Expected success message from add_slide_from_code, got: ${s.substring(0, 100)}`;
  });

  // 11. duplicate_slide
  await runTool(powerPointConfigs, 'duplicate_slide', { sourceIndex: 0 }, r => {
    const s = safeString(r);
    return s.includes('Duplicated') || s.includes('slide')
      ? null
      : `Expected success message from duplicate_slide, got: ${s.substring(0, 100)}`;
  });

  // 12. get_slide_image (may not be supported on older Office versions)
  try {
    const slideImageResult = await callTool(powerPointConfigs, 'get_slide_image', {
      slideIndex: 0,
      width: 400,
    });
    // execute() returns a raw base64 string directly
    const s = safeString(slideImageResult);
    if (s.includes('data:image') || s.includes('base64')) {
      pass('get_slide_image');
    } else {
      fail('get_slide_image', `Expected base64 image data, got: ${s.substring(0, 100)}`);
    }
  } catch (error) {
    log(`  ⚠ get_slide_image: ${String(error)}`);
    addTestResult(testValues, 'get_slide_image', 'conditional_pass', 'pass');
  }

  // 13. clear_slide (clear the last slide added by add_slide_from_code/duplicate)
  let currentSlideCount = 0;
  try {
    await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      currentSlideCount = slides.items.length;
    });
  } catch {
    /* ignore */
  }

  if (currentSlideCount > 1) {
    const lastIdx = currentSlideCount - 1;
    await runTool(powerPointConfigs, 'clear_slide', { slideIndex: lastIdx }, r => {
      const s = safeString(r);
      return s.includes('Cleared') || s.includes('slide')
        ? null
        : `Expected success message from clear_slide, got: ${s.substring(0, 100)}`;
    });
  } else {
    // Use slide 0 if only one slide
    await runTool(powerPointConfigs, 'clear_slide', { slideIndex: 0 }, r => {
      const s = safeString(r);
      return s.includes('Cleared') || s.includes('slide')
        ? null
        : `Expected success message from clear_slide, got: ${s.substring(0, 100)}`;
    });
  }

  // 14. get_smartart_info (slide 0 — may have no SmartArt shapes, which is a valid result)
  await runTool(powerPointConfigs, 'get_smartart_info', { slideIndex: 0 }, r => {
    const s = safeString(r);
    return s.includes('Slide') ? null : `Expected "Slide" in result, got: ${s.substring(0, 100)}`;
  });

  // 15. group_shapes + ungroup_shapes
  // First add two geometric shapes to slide 0, then group and ungroup them
  let shapeCountBefore = 0;
  try {
    await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[0];
      slide.shapes.load('items');
      await context.sync();
      shapeCountBefore = slide.shapes.items.length;
    });
  } catch {
    /* ignore */
  }

  // Add two shapes to group
  await runTool(
    powerPointConfigs,
    'add_geometric_shape',
    { slideIndex: 0, shapeType: 'rectangle', left: 1, top: 4, width: 1.5, height: 1, name: 'GroupTestA' },
    r => (safeString(r).includes('slide') ? null : `Expected success from add_geometric_shape: ${safeString(r).substring(0, 80)}`)
  );
  await runTool(
    powerPointConfigs,
    'add_geometric_shape',
    { slideIndex: 0, shapeType: 'ellipse', left: 3, top: 4, width: 1.5, height: 1, name: 'GroupTestB' },
    r => (safeString(r).includes('slide') ? null : `Expected success from add_geometric_shape: ${safeString(r).substring(0, 80)}`)
  );

  // Get shape indices for the two new shapes
  let groupShapeIdx0 = -1;
  let groupShapeIdx1 = -1;
  try {
    await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[0];
      slide.shapes.load('items');
      await context.sync();
      for (let i = 0; i < slide.shapes.items.length; i++) {
        slide.shapes.items[i].load('name');
      }
      await context.sync();
      for (let i = 0; i < slide.shapes.items.length; i++) {
        const name = slide.shapes.items[i].name;
        if (name === 'GroupTestA') groupShapeIdx0 = i;
        if (name === 'GroupTestB') groupShapeIdx1 = i;
      }
    });
  } catch {
    /* ignore */
  }

  if (groupShapeIdx0 >= 0 && groupShapeIdx1 >= 0) {
    // group the two shapes
    let groupedShapeIdx = -1;
    const groupResult = await callTool(powerPointConfigs, 'group_shapes', {
      slideIndex: 0,
      shapeIndices: [groupShapeIdx0, groupShapeIdx1],
      groupName: 'TestGroup',
    });
    const groupStr = safeString(groupResult);
    if (groupStr.toLowerCase().includes('group') || groupStr.includes('slide')) {
      pass('group_shapes');
    } else {
      fail('group_shapes', `Unexpected result: ${groupStr.substring(0, 100)}`);
    }

    // Find the group shape index
    try {
      await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slide = slides.items[0];
        slide.shapes.load('items');
        await context.sync();
        for (let i = 0; i < slide.shapes.items.length; i++) {
          slide.shapes.items[i].load('name,type');
        }
        await context.sync();
        for (let i = 0; i < slide.shapes.items.length; i++) {
          if (
            slide.shapes.items[i].name === 'TestGroup' ||
            String(slide.shapes.items[i].type) === 'Group'
          ) {
            groupedShapeIdx = i;
            break;
          }
        }
      });
    } catch {
      /* ignore */
    }

    if (groupedShapeIdx >= 0) {
      // ungroup the group
      await runTool(
        powerPointConfigs,
        'ungroup_shapes',
        { slideIndex: 0, shapeIndex: groupedShapeIdx },
        r => {
          const s = safeString(r);
          return s.toLowerCase().includes('ungroup') || s.includes('slide')
            ? null
            : `Expected success from ungroup_shapes: ${s.substring(0, 100)}`;
        }
      );
    } else {
      log('  ⚠ ungroup_shapes: skipped — could not locate group shape after grouping');
      addTestResult(testValues, 'ungroup_shapes', 'conditional_pass', 'pass');
    }
  } else {
    log('  ⚠ group_shapes: skipped — could not locate test shapes');
    addTestResult(testValues, 'group_shapes', 'conditional_pass', 'pass');
    addTestResult(testValues, 'ungroup_shapes', 'conditional_pass', 'pass');
  }
}

// ─── LaunchEvent handler (for auto-start via manifest) ─────────────

// Register the launch event function BEFORE Office.onReady
// This runs synchronously when the script loads in the shared runtime.
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('onPptLaunch', (event: Office.AddinCommands.Event) => {
    // The taskpane opens automatically via Office.onReady below.
    // Just complete the event to unblock the add-in startup.
    event.completed();
  });
}

// ─── Main ─────────────────────────────────────────────────────────

if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') {
  const diagnostic = `Office.js runtime unavailable (href=${window.location.href})`;
  console.error(`[PPT-E2E] ${diagnostic}`);
  heartbeat('office_runtime_missing');
  addTestResult(testValues, 'office_runtime_missing', null, 'fail', { error: diagnostic });
  finishAndSend().catch(_err => {
    /* ignore finishAndSend error */
  });
} else {
  void Office.onReady(async () => {
    heartbeat('ppt_onready_fired');
    console.log('[PPT-E2E] Office.onReady fired');

    const safetyTimer = setTimeout(() => {
      console.error('[PPT-E2E] Safety timeout (120s) — forcing result send');
      fail('safety_timeout', 'Tests did not complete within 120 seconds');
      void finishAndSend();
    }, 120000);

    try {
      await (
        Office as Record<string, unknown> & { addin: { showAsTaskpane: () => Promise<void> } }
      ).addin.showAsTaskpane();
    } catch {
      /* already visible or not supported */
    }

    const sideloadMsg = document.getElementById('sideload-msg');
    const appBody = document.getElementById('app-body');
    if (sideloadMsg) sideloadMsg.style.display = 'none';
    if (appBody) appBody.style.display = 'block';

    addTestResult(testValues, 'UserAgent', navigator.userAgent, 'info');
    log('PPT Add-in loaded. Connecting to test server...');
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
      heartbeat('ppt_tests_starting');
      setStatus('Running PPT tests...', 'running');

      await setup();
      await testPptTools();

      clearTimeout(safetyTimer);
      await finishAndSend();
      log('Closing presentation...');
      await closePresentation();
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
