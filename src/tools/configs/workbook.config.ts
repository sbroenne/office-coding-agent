/**
 * Workbook tool configs — 5 tools for workbook overview, selection, and named ranges.
 */

import type { ToolConfig } from '../codegen';

export const workbookConfigs: readonly ToolConfig[] = [
  {
    name: 'get_workbook_info',
    description:
      'Get a high-level overview of the entire workbook. Returns all sheet names, the active sheet, the used range dimensions of the active sheet, a list of all table names, and counts. This is the best starting point to understand what data exists before performing operations.',
    params: {},
    execute: async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      const usedRange = activeSheet.getUsedRangeOrNullObject();
      usedRange.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      const tables = context.workbook.tables;
      tables.load('items');
      await context.sync();

      for (const s of sheets.items) {
        s.load('name');
      }
      for (const t of tables.items) {
        t.load('name');
      }
      await context.sync();

      return {
        sheetNames: sheets.items.map(s => s.name),
        sheetCount: sheets.items.length,
        activeSheet: activeSheet.name,
        usedRange: usedRange.isNullObject ? null : usedRange.address,
        usedRangeRows: usedRange.isNullObject ? 0 : usedRange.rowCount,
        usedRangeColumns: usedRange.isNullObject ? 0 : usedRange.columnCount,
        tableNames: tables.items.map(t => t.name),
        tableCount: tables.items.length,
      };
    },
  },

  {
    name: 'get_selected_range',
    description:
      'Get the address and values of the range the user currently has selected (highlighted) in Excel. Useful when the user says "this data", "selected cells", or "what I have highlighted".',
    params: {},
    execute: async context => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'rowCount', 'columnCount', 'values', 'numberFormat']);
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
    name: 'define_named_range',
    description:
      'Create a named range — a human-readable alias for a cell range (e.g., "SalesData" → Sheet1!A1:D100). Named ranges make formulas more readable and can be used in formulas across the workbook.',
    params: {
      name: { type: 'string', description: 'Name for the range (e.g., "SalesData")' },
      address: { type: 'string', description: 'The range address (e.g., "A1:D100")' },
      comment: {
        type: 'string',
        required: false,
        description: 'Optional comment describing the named range.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.address as string);
      const comment = args.comment as string | undefined;
      const namedItem = context.workbook.names.add(args.name as string, range, comment);
      namedItem.load(['name', 'comment']);
      await context.sync();
      return { name: namedItem.name, address: args.address, comment: namedItem.comment ?? '' };
    },
  },

  {
    name: 'list_named_ranges',
    description:
      'List all named ranges defined in the workbook. Returns each name, the range address it refers to, and any comment.',
    params: {},
    execute: async context => {
      const names = context.workbook.names;
      names.load('items');
      await context.sync();

      for (const n of names.items) {
        n.load(['name', 'comment', 'value']);
      }
      await context.sync();

      const result = names.items.map(n => ({
        name: n.name,
        value: n.value as string,
        comment: n.comment ?? '',
      }));
      return { namedRanges: result, count: result.length };
    },
  },

  {
    name: 'recalculate_workbook',
    description: 'Force a full recalculation of all formulas in the entire workbook.',
    params: {
      recalcType: {
        type: 'string',
        required: false,
        description: 'Recalculation type',
        enum: ['Recalculate', 'Full'],
      },
    },
    execute: async (context, args) => {
      const recalcType = (args.recalcType as string) ?? 'Full';
      context.application.calculate(recalcType as Excel.CalculationType);
      await context.sync();
      return { recalculated: true, type: recalcType };
    },
  },
];
