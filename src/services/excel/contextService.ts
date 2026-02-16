/**
 * Gathers the current Excel context (active sheet, selection, data preview)
 * so the AI agent knows what the user is looking at.
 */

export interface ExcelContext {
  /** Name of the active worksheet */
  activeSheet: string;
  /** Address of the selected range (e.g. "A1:D10") */
  selectionAddress: string;
  /** Dimensions: rows × columns */
  selectionRows: number;
  selectionCols: number;
  /** First few rows of data (headers + sample) as string[][] */
  preview: string[][];
}

/** Maximum rows/cols to include in the data preview. */
const MAX_PREVIEW_ROWS = 6;
const MAX_PREVIEW_COLS = 10;

/**
 * Snapshot the current Excel state: active sheet, selected range, and a
 * small data preview. Returns null if Excel is unavailable.
 */
export async function getExcelContext(): Promise<ExcelContext | null> {
  if (typeof Excel === 'undefined') return null;

  try {
    return await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');

      const sel = ctx.workbook.getSelectedRange();
      sel.load(['address', 'rowCount', 'columnCount']);
      await ctx.sync();

      const rows = sel.rowCount;
      const cols = sel.columnCount;

      // Read a small preview window
      const previewRows = Math.min(rows, MAX_PREVIEW_ROWS);
      const previewCols = Math.min(cols, MAX_PREVIEW_COLS);
      const preview = sel.getCell(0, 0).getResizedRange(previewRows - 1, previewCols - 1);
      preview.load('values');
      await ctx.sync();

      // Strip the sheet name prefix from the address (e.g. "Sheet1!A1:D10" → "A1:D10")
      const rawAddress: string = sel.address;
      const address = rawAddress.includes('!') ? rawAddress.split('!')[1] : rawAddress;

      return {
        activeSheet: sheet.name,
        selectionAddress: address,
        selectionRows: rows,
        selectionCols: cols,
        preview: (preview.values as (string | number | boolean)[][]).map(row =>
          row.map(cell => (cell === null || cell === '' ? '' : String(cell)))
        ),
      };
    });
  } catch {
    return null;
  }
}

/**
 * Format the Excel context as a compact text block to prepend to the user
 * message, so the AI agent knows what the user is looking at.
 */
export function formatExcelContext(ctx: ExcelContext): string {
  const lines: string[] = [
    `[Excel context]`,
    `Sheet: ${ctx.activeSheet}`,
    `Selection: ${ctx.selectionAddress} (${ctx.selectionRows}×${ctx.selectionCols})`,
  ];

  // Include data preview if the selection isn't empty
  const hasData = ctx.preview.some(row => row.some(cell => cell !== ''));
  if (hasData) {
    lines.push('Data preview:');
    for (const row of ctx.preview) {
      lines.push(row.join('\t'));
    }
    if (ctx.selectionRows > MAX_PREVIEW_ROWS || ctx.selectionCols > MAX_PREVIEW_COLS) {
      lines.push('(truncated)');
    }
  }

  return lines.join('\n');
}
