/* global setTimeout */

/**
 * Sleep for a given number of milliseconds.
 */
export function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Test result sent to the test server.
 */
export interface TestResult {
  Name: string;
  Value: unknown;
  Type: string;
  Metadata: Record<string, unknown>;
  Timestamp: string;
}

/**
 * Add a result to the results array.
 */
export function addTestResult(
  testValues: TestResult[],
  name: string,
  value: unknown,
  type: string,
  metadata?: Record<string, unknown>
): void {
  testValues.push({
    Name: name,
    Value: value,
    Type: type,
    Metadata: metadata ?? {},
    Timestamp: new Date().toISOString(),
  });
}

/**
 * Signal close of the current Word document without saving.
 * The test runner closes Word via process kill after collecting results.
 */
export async function closeDocument(): Promise<void> {
  await sleep(3000);
  // no-op: runner kills WINWORD process in the after() hook
}
