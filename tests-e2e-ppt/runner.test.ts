/**
 * PowerPoint AI E2E Test Runner
 *
 * Orchestrates E2E testing by:
 * 1. Starting a custom test server to receive results from PowerPoint (via POST body)
 * 2. Building and serving the test add-in via Vite (port 3002)
 * 3. Sideloading into PowerPoint Desktop
 * 4. Collecting and validating test results
 *
 * The test-taskpane.ts (in PowerPoint) calls each tool's handler() directly —
 * handlers internally call PowerPoint.run() — reports results to this runner.
 *
 * Auto-start: The manifest uses LaunchEvent (OnNewPresentation) to start the
 * shared runtime automatically when PowerPoint opens a new presentation.
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

const host = 'powerpoint';
const manifestPath = path.resolve(`${process.cwd()}/tests-e2e-ppt/test-manifest.xml`);
const port = 4202;

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

const pptTools = [
  'get_presentation_overview',
  'get_presentation_content',
  'get_presentation_content:single',
  'get_presentation_content:range',
  'get_slide_notes',
  'get_slide_notes:all',
  'set_presentation_content',
  'update_slide_shape',
  'set_slide_notes',
  'add_slide_from_code',
  'duplicate_slide',
  'get_slide_image',
  'clear_slide',
];

// ─── Helper ───────────────────────────────────────────────────────

function assertToolResult(name: string): void {
  const result = e2eContext.getResult(name);
  assert.ok(result, `No result received for "${name}" — test may not have run in PowerPoint`);
  assert.strictEqual(
    result.Type,
    'pass',
    `${name}: ${(result.Metadata?.error as string) || 'Test failed'}`
  );
}

// ─── Test Suite ───────────────────────────────────────────────────

describe('PowerPoint AI E2E Tests', function () {
  this.timeout(0);

  before(`Setup: start test server and sideload ${host}`, async () => {
    console.log('Setting up PowerPoint test environment...');

    await testServer.start();
    const serverResponse = await pingServer(port);
    assert.strictEqual(
      (serverResponse as { status: number }).status,
      200,
      'Test server should respond'
    );
    console.log(`Test server started on port ${port}`);

    const devServerCmd = 'npx vite --config ./tests-e2e-ppt/vite.config.ts';
    const options = {
      appType: AppType.Desktop,
      app: toOfficeApp(host),
      devServerCommandLine: devServerCmd,
      devServerPort: 3002,
      enableDebugging: false,
    };

    console.log('Starting dev server and sideloading add-in into PowerPoint...');
    await startDebugging(manifestPath, options);
    console.log('Add-in sideloaded into PowerPoint');

    console.log('Waiting for test results from add-in...');
    const results = await testServer.getResults();
    e2eContext.setResults(results);
    console.log(`Received ${results.length} test results`);

    const userAgent = results.find(v => v.Name === 'UserAgent');
    if (userAgent) {
      console.log(`User Agent: ${String(userAgent.Value)}`);
    }
  });

  after('Teardown: stop server, close PowerPoint, unregister add-in', async () => {
    console.log('Tearing down...');
    await testServer.stop();

    console.log('Closing PowerPoint...');
    try {
      await closeDesktopApplication();
    } catch (error) {
      console.log(`Note: PowerPoint may already be closed: ${String(error)}`);
    }

    console.log('Stopping debugging...');
    await stopDebugging(manifestPath);
    console.log('Teardown complete');
  });

  // ─── Result Collection ─────────────────────────────────────────

  describe('Result Collection', () => {
    it('should receive test results from PowerPoint', () => {
      const results = e2eContext.getResults();
      assert.ok(results.length >= 1, `Expected at least 1 result, got ${results.length}`);
    });
  });

  // ─── Tool Tests ────────────────────────────────────────────────

  describe('PowerPoint Tools (13)', () => {
    for (const name of pptTools) {
      it(name, () => {
        assertToolResult(name);
      });
    }
  });

  // ─── Summary ───────────────────────────────────────────────────

  describe('Summary', () => {
    it('all in-PowerPoint tests should pass', () => {
      const failed = e2eContext.getFailedTests();
      if (failed.length > 0) {
        const names = failed
          .map(f => `${f.Name}: ${(f.Metadata?.error as string) || 'unknown'}`)
          .join('\n  ');
        assert.fail(`${failed.length} test(s) failed in PowerPoint:\n  ${names}`);
      }
    });
  });
});
