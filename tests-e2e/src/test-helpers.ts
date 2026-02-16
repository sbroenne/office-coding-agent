/* global Excel, setTimeout */

/**
 * Sleep for a given number of milliseconds.
 */
export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
  metadata?: Record<string, unknown>,
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
 * Close the current Excel workbook without saving.
 */
export async function closeWorkbook(): Promise<void> {
  await sleep(3000);
  try {
    await Excel.run(async (context) => {
      context.workbook.close(Excel.CloseBehavior.skipSave);
      await Promise.resolve();
    });
  } catch (err) {
    await Promise.reject(`Error on closing workbook: ${err}`);
  }
}
