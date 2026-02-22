/**
 * Word AI E2E Test Runner
 *
 * Orchestrates E2E testing by:
 * 1. Starting a custom test server to receive results from Word (via POST body)
 * 2. Building and serving the test add-in via Vite (port 3003)
 * 3. Sideloading into Word Desktop
 * 4. Collecting and validating test results
 *
 * The test-taskpane.ts (in Word) calls each tool's handler() directly —
 * handlers internally call Word.run() — and reports results to this runner.
 *
 * Auto-start: The manifest uses LaunchEvent (OnNewDocument) to start the
 * shared runtime automatically when Word opens a new document.
 */
/* eslint-disable vitest/expect-expect, vitest/valid-title */

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

const host = 'word';
const manifestPath = path.resolve(`${process.cwd()}/tests-e2e-word/test-manifest.xml`);
const port = 4203;

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
      const rawMsg = req.query.msg;
      const msg = typeof rawMsg === 'string' ? rawMsg : '(no message)';
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

    const server = https.createServer(options, this.app);
    this.server = server;
    await new Promise<void>((resolve, reject) => {
      server.listen(this.serverPort, () => {
        resolve();
      });
      server.on('error', reject);
    });
  }

  async getResults(): Promise<TestResult[]> {
    return this.resultsPromise;
  }

  stop(): Promise<void> {
    if (this.server) {
      this.server.close();
    }
    return Promise.resolve();
  }
}

const testServer = new CustomTestServer(port);

// ─── Tool name lists (must match test-taskpane.ts) ────────────────

const wordToolNames = [
  'get_document_overview',
  'get_document_content',
  'get_document_section',
  'get_selection',
  'get_selection_text',
  'insert_content_at_selection',
  'find_and_replace',
  'insert_table',
  'insert_table:striped',
  'apply_style_to_selection',
  'apply_style_to_selection:font',
  'set_document_content',
];

// ─── Helper ───────────────────────────────────────────────────────

function assertToolResult(name: string): void {
  const result = e2eContext.getResult(name);
  assert.ok(result, `No result received for "${name}" — test may not have run in Word`);
  assert.strictEqual(
    result.Type,
    'pass',
    `${name}: ${(result.Metadata?.error as string) || 'Test failed'}`
  );
}

// ─── Test Suite ───────────────────────────────────────────────────

describe('Word AI E2E Tests', function () {
  this.timeout(0);

  before(`Setup: start test server and sideload ${host}`, async () => {
    console.log('Setting up Word test environment...');

    await testServer.start();
    const serverResponse = await pingServer(port);
    assert.strictEqual(
      (serverResponse as { status: number }).status,
      200,
      'Test server should respond'
    );
    console.log(`Test server started on port ${port}`);

    const devServerCmd = 'npx vite --config ./tests-e2e-word/vite.config.ts';
    const options = {
      appType: AppType.Desktop,
      app: toOfficeApp(host),
      devServerCommandLine: devServerCmd,
      devServerPort: 3003,
      enableDebugging: false,
    };

    console.log('Starting dev server and sideloading add-in into Word...');
    await startDebugging(manifestPath, options);
    console.log('Add-in sideloaded into Word');

    console.log('Waiting for test results from add-in...');
    const results = await testServer.getResults();
    e2eContext.setResults(results);
    console.log(`Received ${results.length} test results`);

    const userAgent = results.find(v => v.Name === 'UserAgent');
    if (userAgent) {
      console.log(`User Agent: ${String(userAgent.Value)}`);
    }
  });

  after('Teardown: stop server, close Word, unregister add-in', async () => {
    console.log('Tearing down...');
    await testServer.stop();

    console.log('Closing Word...');
    try {
      await closeDesktopApplication();
    } catch (error) {
      console.log(`Note: Word may already be closed: ${String(error)}`);
    }

    console.log('Stopping debugging...');
    await stopDebugging(manifestPath);
    console.log('Teardown complete');
  });

  // ─── Result Collection ─────────────────────────────────────────

  describe('Result Collection', () => {
    it('should receive test results from Word', () => {
      const results = e2eContext.getResults();
      assert.ok(results.length >= 1, `Expected at least 1 result, got ${results.length}`);
    });
  });

  // ─── Tool Tests ────────────────────────────────────────────────

  describe('Word Tools (12)', () => {
    for (const name of wordToolNames) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  // ─── Summary ───────────────────────────────────────────────────

  describe('Summary', () => {
    it('all in-Word tests should pass', () => {
      const failed = e2eContext.getFailedTests();
      if (failed.length > 0) {
        const names = failed
          .map(f => `${f.Name}: ${(f.Metadata?.error as string) || 'unknown'}`)
          .join('\n  ');
        assert.fail(`${failed.length} test(s) failed in Word:\n  ${names}`);
      }
    });
  });
});
