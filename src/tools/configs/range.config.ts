/**
 * Range tool configs — 31 tools for reading, writing, formatting, and
 * manipulating cell ranges.
 *
 * Fixes applied (from tool audit):
 *   - get_used_range: description now mentions values are returned
 *   - clear_range: code uses ClearApplyTo.all (matches "clears formatting" description)
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const rangeConfigs: readonly ToolConfig[] = [
  // ─── Read ───────────────────────────────────────────────
  {
    name: 'get_range_values',
    description:
      'Read cell values from a specified range. Returns a 2D array of the displayed values (not formulas). Use get_range_formulas instead if you need to inspect formulas.',
    params: {
      address: {
        type: 'string',
        description: 'The range address (e.g., "A1:C10", "Sheet1!B2:D5")',
      },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
      };
    },
  },

  {
    name: 'set_range_values',
    description:
      'Write values to a specified range, overwriting any existing content. Values must be a 2D array matching the range dimensions. Use set_range_formulas for formulas.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "A1:C3")' },
      values: { type: 'any[][]', description: '2D array of values to write (rows × columns)' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const values = args.values as unknown[][];
      range.values = values;
      range.load('address');
      await context.sync();
      return {
        address: range.address,
        rowsWritten: values.length,
        columnsWritten: values[0]?.length ?? 0,
      };
    },
  },

  {
    name: 'get_used_range',
    description:
      'Get the bounding rectangle of all non-empty cells on a worksheet. By default returns only the address and dimensions (rowCount, columnCount) — no cell values. Set maxRows to include values for the first N rows (e.g., maxRows=5 to preview headers and a few data rows). Use get_range_values to read a specific sub-range when you already know the address.',
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name. Uses active sheet if omitted.',
      },
      maxRows: {
        type: 'number',
        required: false,
        description:
          'How many rows of cell values to include in the response. Omit to get dimensions only (no values). Set to a small number like 5 to preview headers + a few data rows, saving tokens on large sheets.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const usedRange = sheet.getUsedRange();

      const maxRows = args.maxRows as number | undefined;
      const includeValues = maxRows != null;

      if (includeValues) {
        usedRange.load(['address', 'rowCount', 'columnCount', 'values']);
      } else {
        usedRange.load(['address', 'rowCount', 'columnCount']);
      }
      await context.sync();

      const result: Record<string, unknown> = {
        address: usedRange.address,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
      };

      if (includeValues) {
        result.values =
          maxRows < usedRange.rowCount ? usedRange.values.slice(0, maxRows) : usedRange.values;
        if (maxRows < usedRange.rowCount) {
          result.truncated = true;
          result.rowsReturned = maxRows;
        }
      }

      return result;
    },
  },

  // ─── Clear ──────────────────────────────────────────────
  {
    name: 'clear_range',
    description:
      'Clear all content (values, formulas, and formatting) from a specified range. The cells remain but become empty.',
    params: {
      address: { type: 'string', description: 'The range address to clear' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.clear(Excel.ClearApplyTo.all);
      range.load('address');
      await context.sync();
      return { address: range.address, cleared: true };
    },
  },

  // ─── Formatting ─────────────────────────────────────────
  {
    name: 'format_range',
    description:
      'Apply visual formatting to a range of cells. Only set the properties you want to change — omitted properties are left unchanged. Supports bold, italic, underline, font size, font color, fill/background color, horizontal alignment, and text wrapping. For number display formats (currency, dates, percent), use set_number_format instead.',
    params: {
      address: { type: 'string', description: 'The range address to format (e.g., "A1:C1")' },
      bold: { type: 'boolean', required: false, description: 'Make text bold' },
      italic: { type: 'boolean', required: false, description: 'Make text italic' },
      underline: { type: 'boolean', required: false, description: 'Underline text' },
      fontSize: {
        type: 'number',
        required: false,
        description: 'Font size in points (e.g., 12, 14, 18)',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color as hex (e.g., "#FF0000" for red) or named color',
      },
      fillColor: {
        type: 'string',
        required: false,
        description: 'Cell background/fill color as hex (e.g., "#FFFF00" for yellow)',
      },
      horizontalAlignment: {
        type: 'string',
        required: false,
        description: 'Horizontal text alignment',
        enum: ['General', 'Left', 'Center', 'Right', 'Fill', 'Justify', 'Distributed'],
      },
      wrapText: { type: 'boolean', required: false, description: 'Enable text wrapping in cells' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const bold = args.bold as boolean | undefined;
      const italic = args.italic as boolean | undefined;
      const fontSize = args.fontSize as number | undefined;
      const fontColor = args.fontColor as string | undefined;
      const fillColor = args.fillColor as string | undefined;
      const horizontalAlignment = args.horizontalAlignment as string | undefined;
      const underline = args.underline as boolean | undefined;
      const wrapText = args.wrapText as boolean | undefined;

      if (bold !== undefined) range.format.font.bold = bold;
      if (italic !== undefined) range.format.font.italic = italic;
      if (fontSize !== undefined) range.format.font.size = fontSize;
      if (fontColor !== undefined) range.format.font.color = fontColor;
      if (fillColor !== undefined) range.format.fill.color = fillColor;
      if (horizontalAlignment !== undefined) {
        range.format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment;
      }
      if (underline !== undefined) {
        range.format.font.underline = underline
          ? Excel.RangeUnderlineStyle.single
          : Excel.RangeUnderlineStyle.none;
      }
      if (wrapText !== undefined) range.format.wrapText = wrapText;

      range.load('address');
      await context.sync();

      return {
        address: range.address,
        formatted: true,
        appliedFormats: {
          ...(bold !== undefined && { bold }),
          ...(italic !== undefined && { italic }),
          ...(fontSize !== undefined && { fontSize }),
          ...(fontColor !== undefined && { fontColor }),
          ...(fillColor !== undefined && { fillColor }),
          ...(horizontalAlignment !== undefined && { horizontalAlignment }),
          ...(underline !== undefined && { underline }),
          ...(wrapText !== undefined && { wrapText }),
        },
      };
    },
  },

  {
    name: 'set_number_format',
    description:
      'Apply a number format to a range. Common formats: "#,##0.00" (number), "$#,##0.00" (currency), "0.00%" (percent), "yyyy-mm-dd" (date), "0" (integer).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "B2:B10")' },
      format: {
        type: 'string',
        description:
          'Excel number format string. Examples: "#,##0.00", "$#,##0.00", "0.00%", "yyyy-mm-dd", "hh:mm:ss", "0", "@" (text)',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const format = args.format as string;
      // Use 1×1 broadcast — Excel applies the single format to the entire range
      range.numberFormat = [[format]];
      range.load(['address', 'rowCount', 'columnCount']);
      await context.sync();
      return { address: range.address, format, cellCount: range.rowCount * range.columnCount };
    },
  },

  {
    name: 'auto_fit_columns',
    description:
      'Auto-fit column widths to fit their content. If no address is specified, fits all used columns.',
    params: {
      address: {
        type: 'string',
        required: false,
        description: 'Optional range address. Columns in this range will be auto-fitted.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string | undefined;
      const range = address ? sheet.getRange(address) : sheet.getUsedRange();
      range.format.autofitColumns();
      range.load('address');
      await context.sync();
      return { address: range.address, autoFitted: true };
    },
  },

  {
    name: 'auto_fit_rows',
    description:
      'Auto-fit row heights to fit their content. If no address is specified, fits all used rows.',
    params: {
      address: {
        type: 'string',
        required: false,
        description: 'Optional range address. Rows in this range will be auto-fitted.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string | undefined;
      const range = address ? sheet.getRange(address) : sheet.getUsedRange();
      range.format.autofitRows();
      range.load('address');
      await context.sync();
      return { address: range.address, autoFitted: true };
    },
  },

  // ─── Formulas ───────────────────────────────────────────
  {
    name: 'set_range_formulas',
    description:
      'Write Excel formulas to cells. Formulas must start with "=" (e.g., "=SUM(A1:A10)").',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "D1")' },
      formulas: {
        type: 'string[][]',
        description: '2D array of formula strings (e.g., [["=SUM(A1:A10)"]])',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const formulas = args.formulas as string[][];
      range.formulas = formulas;
      range.load('address');
      await context.sync();
      return {
        address: range.address,
        rowsWritten: formulas.length,
        columnsWritten: formulas[0]?.length ?? 0,
      };
    },
  },

  {
    name: 'get_range_formulas',
    description:
      'Read the formulas from a range. Returns formulas as strings (cells without formulas return their value).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "A1:D10")' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.load(['formulas', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        formulas: range.formulas,
      };
    },
  },

  // ─── Sort ───────────────────────────────────────────────
  {
    name: 'sort_range',
    description:
      'Sort a range of cells in-place by a specified column. Best for ad-hoc data ranges. For structured Excel tables, use sort_table instead.',
    params: {
      address: { type: 'string', description: 'The range address to sort (e.g., "A1:D20")' },
      column: {
        type: 'number',
        description: 'Zero-based column index to sort by (0 = first column)',
      },
      ascending: {
        type: 'boolean',
        required: false,
        description: 'Sort ascending (true) or descending (false). Default true.',
      },
      hasHeaders: {
        type: 'boolean',
        required: false,
        description: 'First row is a header row. Default true.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const column = args.column as number;
      const ascending = (args.ascending as boolean) ?? true;
      const hasHeaders = (args.hasHeaders as boolean) ?? true;
      range.sort.apply([{ key: column, ascending }], hasHeaders);
      range.load('address');
      await context.sync();
      return { address: range.address, sortedByColumn: column, ascending };
    },
  },

  {
    name: 'auto_fill_range',
    description:
      'Auto-fill a pattern from a source range into a destination range. Useful for extending formulas, dates, or sequences.',
    params: {
      sourceAddress: {
        type: 'string',
        description: 'Source range containing the pattern (e.g., "A1:A2")',
      },
      destinationAddress: {
        type: 'string',
        description: 'Destination range to fill (e.g., "A1:A20")',
      },
      autoFillType: {
        type: 'string',
        required: false,
        description: 'Auto-fill behavior. Default FillDefault.',
        enum: [
          'FillDefault',
          'FillCopy',
          'FillSeries',
          'FillFormats',
          'FillValues',
          'FillDays',
          'FillWeekdays',
          'FillMonths',
          'FillYears',
          'LinearTrend',
          'GrowthTrend',
          'FlashFill',
        ],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const sourceRange = sheet.getRange(args.sourceAddress as string);
      const destinationAddress = args.destinationAddress as string;
      const autoFillType = ((args.autoFillType as string | undefined) ??
        'FillDefault') as Excel.AutoFillType;

      sourceRange.autoFill(destinationAddress, autoFillType);
      const destinationRange = sheet.getRange(destinationAddress);
      destinationRange.load('address');
      await context.sync();

      return {
        sourceAddress: args.sourceAddress,
        destinationAddress: destinationRange.address,
        autoFillType,
        filled: true,
      };
    },
  },

  {
    name: 'flash_fill_range',
    description:
      'Apply Flash Fill to a range based on adjacent data patterns (Excel-detected extraction/transformation pattern).',
    params: {
      address: {
        type: 'string',
        description: 'Range address to flash-fill (e.g., "B2:B100")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.flashFill();
      range.load('address');
      await context.sync();
      return { address: range.address, flashFilled: true };
    },
  },

  {
    name: 'get_special_cells',
    description:
      'Get a RangeAreas result containing only special cells in a range (e.g., blanks, formulas, constants, visible cells, conditional formats).',
    params: {
      address: {
        type: 'string',
        description: 'Range address to inspect (e.g., "A1:D200")',
      },
      cellType: {
        type: 'string',
        description: 'Special cell category to return',
        enum: [
          'ConditionalFormats',
          'DataValidations',
          'Blanks',
          'Constants',
          'Formulas',
          'SameConditionalFormat',
          'SameDataValidation',
          'Visible',
        ],
      },
      cellValueType: {
        type: 'string',
        required: false,
        description: 'Optional value filter for constants/formulas.',
        enum: [
          'All',
          'Errors',
          'ErrorsLogical',
          'ErrorsNumbers',
          'ErrorsText',
          'ErrorsLogicalNumber',
          'ErrorsLogicalText',
          'ErrorsNumberText',
          'Logical',
          'LogicalNumbers',
          'LogicalText',
          'LogicalNumbersText',
          'Numbers',
          'NumbersText',
          'Text',
        ],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const cellType = args.cellType as Excel.SpecialCellType;
      const cellValueType = args.cellValueType as Excel.SpecialCellValueType | undefined;
      const special = range.getSpecialCells(cellType, cellValueType);
      special.load(['address', 'cellCount', 'areaCount']);
      await context.sync();
      return {
        sourceAddress: args.address,
        specialAddress: special.address,
        areaCount: special.areaCount,
        cellCount: special.cellCount,
        cellType,
        cellValueType: cellValueType ?? null,
      };
    },
  },

  {
    name: 'get_range_precedents',
    description:
      'Get direct precedent cells (cells referenced by formulas) for a range as cross-sheet addresses.',
    params: {
      address: {
        type: 'string',
        description: 'Range containing formulas to inspect (e.g., "D2:D20")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const precedents = range.getDirectPrecedents();
      precedents.load('addresses');
      await context.sync();
      return {
        sourceAddress: args.address,
        addresses: precedents.addresses,
        count: precedents.addresses.length,
      };
    },
  },

  {
    name: 'get_range_dependents',
    description:
      'Get direct dependent cells (formulas that reference this range) as cross-sheet addresses.',
    params: {
      address: {
        type: 'string',
        description: 'Range whose dependents should be found (e.g., "B2")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const dependents = range.getDirectDependents();
      dependents.load('addresses');
      await context.sync();
      return {
        sourceAddress: args.address,
        addresses: dependents.addresses,
        count: dependents.addresses.length,
      };
    },
  },

  {
    name: 'recalculate_range',
    description: 'Force recalculation for formulas in a specific range.',
    params: {
      address: {
        type: 'string',
        description: 'Range address to recalculate (e.g., "A1:Z100")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.calculate();
      await context.sync();
      return { address: args.address, recalculated: true };
    },
  },

  {
    name: 'get_tables_for_range',
    description: 'List tables that overlap (or are fully contained in) a given range.',
    params: {
      address: {
        type: 'string',
        description: 'Range address to check for intersecting tables (e.g., "A1:H500")',
      },
      fullyContained: {
        type: 'boolean',
        required: false,
        description:
          'If true, only include tables fully contained in the range. Default false (any overlap).',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const fullyContained = (args.fullyContained as boolean | undefined) ?? false;
      const tables = range.getTables(fullyContained);
      tables.load('items');
      await context.sync();

      for (const table of tables.items) {
        table.load(['name', 'id']);
      }
      await context.sync();

      const result = tables.items.map(table => ({
        name: table.name,
        id: table.id,
      }));

      return {
        address: args.address,
        fullyContained,
        tables: result,
        count: result.length,
      };
    },
  },

  // ─── Copy ───────────────────────────────────────────────
  {
    name: 'copy_range',
    description:
      'Copy values, formulas, and formatting from a source range to a destination range. Supports copying across worksheets by specifying sourceSheet and destinationSheet.',
    params: {
      sourceAddress: { type: 'string', description: 'Source range address (e.g., "A1:C10")' },
      destinationAddress: { type: 'string', description: 'Destination range address (e.g., "E1")' },
      sourceSheet: {
        type: 'string',
        required: false,
        description: 'Source worksheet name. Uses active sheet if omitted.',
      },
      destinationSheet: {
        type: 'string',
        required: false,
        description: 'Destination worksheet name. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const srcSheet = getSheet(context, args.sourceSheet as string | undefined);
      const dstSheet = getSheet(context, args.destinationSheet as string | undefined);
      const srcRange = srcSheet.getRange(args.sourceAddress as string);
      const dstRange = dstSheet.getRange(args.destinationAddress as string);
      dstRange.copyFrom(srcRange);
      srcRange.load('address');
      dstRange.load('address');
      await context.sync();
      return { source: srcRange.address, destination: dstRange.address, copied: true };
    },
  },

  // ─── Find ───────────────────────────────────────────────
  {
    name: 'find_values',
    description:
      'Search for text or a value across all cells in a worksheet. Returns the cell address and value of the first match, or a "not found" message if no match exists.',
    params: {
      searchText: { type: 'string', description: 'The text or value to search for' },
      matchCase: {
        type: 'boolean',
        required: false,
        description: 'Case-sensitive search. Default false.',
      },
      matchEntireCell: {
        type: 'boolean',
        required: false,
        description: 'Match the entire cell content. Default false.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const usedRange = sheet.getUsedRange();
      const matchCase = (args.matchCase as boolean) ?? false;
      const matchEntireCell = (args.matchEntireCell as boolean) ?? false;
      const searchText = args.searchText as string;
      const result = usedRange.find(searchText, {
        completeMatch: matchEntireCell,
        matchCase,
      });
      result.load(['address', 'values']);
      try {
        await context.sync();
        return {
          found: true,
          address: result.address,
          value: result.values[0]?.[0] as string | number | boolean | null | undefined,
        };
      } catch {
        return { found: false, searchText, message: 'No match found.' };
      }
    },
  },

  // ─── Insert / Delete ───────────────────────────────────
  {
    name: 'insert_range',
    description:
      'Insert blank cells at the specified address, shifting existing cells down or right to make room. Use row notation (e.g., "3:5") to insert entire rows, column notation (e.g., "B:D") for entire columns, or a cell range for a block.',
    params: {
      address: {
        type: 'string',
        description:
          'Range address where blank cells will be inserted (e.g., "3:5" for rows 3-5, "B:D" for columns B-D, "A1:C3" for a block)',
      },
      shift: {
        type: 'string',
        required: false,
        description: 'Direction to shift existing cells. Default "down".',
        enum: ['down', 'right'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const shift = (args.shift as string) ?? 'down';
      const range = sheet.getRange(address);
      const shiftDir =
        shift === 'right' ? Excel.InsertShiftDirection.right : Excel.InsertShiftDirection.down;
      range.insert(shiftDir);
      await context.sync();
      return { address, shift, inserted: true };
    },
  },

  {
    name: 'delete_range',
    description:
      'Delete cells at the specified address, shifting remaining cells up or left to fill the gap. Use row notation (e.g., "3:5") to delete entire rows, column notation (e.g., "B:D") for entire columns, or a cell range for a block. Data in deleted cells is permanently removed.',
    params: {
      address: {
        type: 'string',
        description:
          'Range address to delete (e.g., "3:5" for rows 3-5, "B:D" for columns, "A1:C3" for a block)',
      },
      shift: {
        type: 'string',
        required: false,
        description: 'Direction to shift remaining cells. Default "up".',
        enum: ['up', 'left'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const shift = (args.shift as string) ?? 'up';
      const range = sheet.getRange(address);
      const shiftDir =
        shift === 'left' ? Excel.DeleteShiftDirection.left : Excel.DeleteShiftDirection.up;
      range.delete(shiftDir);
      await context.sync();
      return { address, shift, deleted: true };
    },
  },

  // ─── Merge / Unmerge ────────────────────────────────────
  {
    name: 'merge_cells',
    description:
      'Merge a range of cells into a single cell. WARNING: Only the upper-left cell\'s value is preserved; values in all other cells are discarded. Use the "across" option to merge each row separately (e.g., for header rows that span columns).',
    params: {
      address: { type: 'string', description: 'The range address to merge (e.g., "A1:D1")' },
      across: {
        type: 'boolean',
        required: false,
        description: 'Merge each row independently (merge across). Default false.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const across = (args.across as boolean) ?? false;
      range.merge(across);
      range.load('address');
      await context.sync();
      return { address: range.address, merged: true, across };
    },
  },

  {
    name: 'unmerge_cells',
    description: 'Unmerge a previously merged range of cells.',
    params: {
      address: { type: 'string', description: 'The range address to unmerge' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.unmerge();
      range.load('address');
      await context.sync();
      return { address: range.address, unmerged: true };
    },
  },

  // ─── Find & Replace ──────────────────────────────────────

  {
    name: 'replace_values',
    description:
      'Find and replace text across a range or the entire used range on a sheet. Returns the number of replacements made.',
    params: {
      find: { type: 'string', description: 'The text to search for' },
      replace: { type: 'string', description: 'The replacement text' },
      address: {
        type: 'string',
        required: false,
        description:
          'Range to search in (e.g., "A1:C10"). If omitted, searches the entire used range.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string | undefined;
      const range = address ? sheet.getRange(address) : sheet.getUsedRange();
      const count = range.replaceAll(args.find as string, args.replace as string, {
        completeMatch: false,
        matchCase: false,
      });
      await context.sync();
      return { find: args.find, replace: args.replace, replacements: count.value };
    },
  },

  // ─── Remove Duplicates ───────────────────────────────────

  {
    name: 'remove_duplicates',
    description:
      'Remove duplicate rows from a range based on specified columns. Returns the number of rows removed and remaining.',
    params: {
      address: {
        type: 'string',
        description: 'The range address containing data (e.g., "A1:D100")',
      },
      columns: {
        type: 'string[]',
        description:
          'Zero-based column indices to compare for duplicates (e.g., ["0","2"] for columns A and C). Pass all column indices to check entire rows.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const columnsRaw = args.columns as string[];
      const columnIndices = columnsRaw.map(Number);
      const result = range.removeDuplicates(columnIndices, true);
      result.load(['removed', 'uniqueRemaining']);
      await context.sync();
      return {
        address: args.address,
        rowsRemoved: result.removed,
        rowsRemaining: result.uniqueRemaining,
      };
    },
  },

  // ─── Hyperlinks ──────────────────────────────────────────

  {
    name: 'set_hyperlink',
    description:
      'Set or remove a hyperlink on a single cell. Supports web URLs, email addresses, and links to other cells within the workbook.',
    params: {
      address: { type: 'string', description: 'The cell address (e.g., "A1")' },
      url: {
        type: 'string',
        description:
          'The hyperlink URL or cell reference (e.g., "https://example.com", "mailto:x@y.com", "Sheet2!A1"). Use "" to remove.',
      },
      textToDisplay: {
        type: 'string',
        required: false,
        description: 'Display text for the hyperlink. If omitted, the URL is shown.',
      },
      tooltip: {
        type: 'string',
        required: false,
        description: 'Tooltip shown on hover',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const url = args.url as string;
      if (url === '') {
        range.clear('Hyperlinks');
      } else {
        range.hyperlink = {
          address: url,
          textToDisplay: (args.textToDisplay as string) ?? url,
          screenTip: (args.tooltip as string) ?? '',
        };
      }
      range.load('address');
      await context.sync();
      return { address: range.address, url: url || null };
    },
  },

  // ─── Hide / Show Rows & Columns ──────────────────────────

  {
    name: 'toggle_row_column_visibility',
    description:
      'Hide or show rows and/or columns in a range. Hidden rows/columns are not visible in the UI but their data is preserved.',
    params: {
      address: {
        type: 'string',
        description:
          'Range specifying which rows/columns to affect (e.g., "A:C" for columns A-C, "3:5" for rows 3-5, "B2:D10" for specific rows and columns)',
      },
      hidden: {
        type: 'boolean',
        description: 'True to hide, false to show',
      },
      target: {
        type: 'string',
        description: 'What to hide/show',
        enum: ['rows', 'columns'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const hidden = args.hidden as boolean;
      const target = args.target as string;
      if (target === 'rows') {
        range.rowHidden = hidden;
      } else {
        range.columnHidden = hidden;
      }
      await context.sync();
      return { address: args.address, target, hidden };
    },
  },

  // ─── Grouping ────────────────────────────────────────────

  {
    name: 'group_rows_columns',
    description: 'Group rows or columns to create an outline for expanding/collapsing sections.',
    params: {
      address: {
        type: 'string',
        description:
          'Range of rows or columns to group (e.g., "3:5" for rows 3-5, "B:D" for columns B-D)',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.group('ByRows');
      await context.sync();
      return { address: args.address, grouped: true };
    },
  },

  {
    name: 'ungroup_rows_columns',
    description: 'Remove a grouping from rows or columns.',
    params: {
      address: {
        type: 'string',
        description: 'Range of grouped rows or columns to ungroup',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.ungroup('ByRows');
      await context.sync();
      return { address: args.address, ungrouped: true };
    },
  },

  // ─── Borders ──────────────────────────────────────────────

  {
    name: 'set_cell_borders',
    description: 'Apply borders to a range. Specify which sides to apply borders to and the style.',
    params: {
      address: {
        type: 'string',
        description: 'The range address to apply borders to',
      },
      borderStyle: {
        type: 'string',
        description: 'Border style (e.g., "Thin", "Medium", "Thick", "Dotted")',
        enum: ['Thin', 'Medium', 'Thick', 'Double', 'Dotted', 'Dashed', 'DashDot'],
      },
      borderColor: {
        type: 'string',
        required: false,
        description: 'Border color as hex (e.g., "#000000" for black)',
      },
      side: {
        type: 'string',
        description: 'Which side(s) to apply border to',
        enum: ['EdgeLeft', 'EdgeRight', 'EdgeTop', 'EdgeBottom', 'EdgeAll'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const borderStyle = args.borderStyle as string;
      const side = args.side as string;
      const color = (args.borderColor as string) ?? '000000';

      // Map user-friendly borderStyle to Excel API style + weight
      const styleMap: Record<string, { style: string; weight?: string }> = {
        Thin: { style: 'Continuous', weight: 'Thin' },
        Medium: { style: 'Continuous', weight: 'Medium' },
        Thick: { style: 'Continuous', weight: 'Thick' },
        Double: { style: 'Double' },
        Dotted: { style: 'Dot' },
        Dashed: { style: 'Dash' },
        DashDot: { style: 'DashDot' },
      };
      const mapped = styleMap[borderStyle] ?? { style: borderStyle };

      const sides: Excel.BorderIndex[] =
        side === 'EdgeAll'
          ? [
              Excel.BorderIndex.edgeLeft,
              Excel.BorderIndex.edgeRight,
              Excel.BorderIndex.edgeTop,
              Excel.BorderIndex.edgeBottom,
            ]
          : [side as Excel.BorderIndex];

      for (const s of sides) {
        const borderObj = range.format.borders.getItem(s);
        borderObj.style = mapped.style as Excel.BorderLineStyle;
        if (mapped.weight) {
          borderObj.weight = mapped.weight as Excel.BorderWeight;
        }
        borderObj.color = color;
      }

      await context.sync();
      return { address: args.address, borderStyle, side, color };
    },
  },
];
