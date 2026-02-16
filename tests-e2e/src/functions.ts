/**
 * Minimal dummy custom function for E2E testing.
 *
 * This exists solely to enable the SharedRuntime auto-start.
 * The CustomFunctions extension point requires at least one function,
 * which triggers Excel to load the shared runtime page (test-taskpane.html)
 * on workbook open — allowing tests to auto-execute.
 */

/* global CustomFunctions */

/**
 * Returns the E2E test add-in version.
 * @customfunction E2E_VERSION
 * @returns The version string.
 */
export function e2eVersion(): string {
  return '1.0.0';
}

// Register with the CustomFunctions runtime
CustomFunctions.associate('E2E_VERSION', e2eVersion);
