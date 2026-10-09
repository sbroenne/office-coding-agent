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
 * Close only the test document without saving; leave other Word windows alone.
 */
export async function closeDocument(): Promise<void> {
  await Word.run(async context => {
    context.document.close(Word.CloseBehavior.skipSave);
    await context.sync();
  });
}
