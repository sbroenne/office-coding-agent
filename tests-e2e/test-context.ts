/**
 * Shared test context for E2E tests.
 * Populated by the runner, accessed by test validation files.
 */

export interface TestResult {
  Name: string;
  Value: unknown;
  Type: string;
  Metadata: Record<string, unknown>;
  Timestamp: string;
}

class E2ETestContext {
  private _results: TestResult[] = [];

  setResults(results: TestResult[]): void {
    this._results = results;
  }

  getResults(): TestResult[] {
    return this._results;
  }

  getResult(name: string): TestResult | undefined {
    return this._results.find((r) => r.Name === name);
  }

  getPassedTests(): TestResult[] {
    return this._results.filter((r) => r.Type === 'pass');
  }

  getFailedTests(): TestResult[] {
    return this._results.filter((r) => r.Type === 'fail');
  }

  hasResults(): boolean {
    return this._results.length > 0;
  }
}

export const e2eContext = new E2ETestContext();
