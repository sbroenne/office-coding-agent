import type { ToolCallResult } from '@/types';

/**
 * Helper to run an Excel.run() context and return a ToolCallResult.
 */
export async function excelRun(
  fn: (context: Excel.RequestContext) => Promise<unknown>,
): Promise<ToolCallResult> {
  try {
    const data = await Excel.run(fn);
    return { success: true, data };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Get a worksheet by name or default to active sheet.
 */
export function getSheet(
  context: Excel.RequestContext,
  sheetName?: string,
): Excel.Worksheet {
  if (sheetName) {
    return context.workbook.worksheets.getItem(sheetName);
  }
  return context.workbook.worksheets.getActiveWorksheet();
}
