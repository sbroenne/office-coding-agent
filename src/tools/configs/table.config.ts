/**
 * Table tool configs — 11 tools for managing structured Excel tables.
 *
 * Fixes applied (from tool audit):
 *   - list_tables: now loads and returns address, worksheet name, column names, row count
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const tableConfigs: readonly ToolConfig[] = [
  {
    name: 'list_tables',
    description:
      "List all structured Excel tables. Returns each table's name, address range, worksheet name, column names, and row count. Optionally filter to a specific sheet.",
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional: filter to tables on a specific sheet.',
      },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      let tables: Excel.TableCollection;
      if (sheetName) {
        const sheet = getSheet(context, sheetName);
        tables = sheet.tables;
      } else {
        tables = context.workbook.tables;
      }
      tables.load(['items/name', 'items/id']);
      await context.sync();

      // Load additional details per table — save proxy references for safe reuse
      const rangeProxies: Excel.Range[] = [];
      const bodyProxies: Excel.Range[] = [];
      for (const table of tables.items) {
        table.load(['name', 'id']);
        const r = table.getRange();
        r.load('address');
        rangeProxies.push(r);
        table.worksheet.load('name');
        table.columns.load('items/name');
        const b = table.getDataBodyRange();
        b.load('rowCount');
        bodyProxies.push(b);
      }
      await context.sync();

      const result = tables.items.map((table, i) => ({
        name: table.name,
        id: table.id,
        address: rangeProxies[i].address,
        worksheetName: table.worksheet.name,
        columns: table.columns.items.map(c => c.name),
        rowCount: bodyProxies[i].rowCount,
      }));

      return { tables: result, count: result.length };
    },
  },

  {
    name: 'create_table',
    description:
      'Convert a data range into a structured Excel table with sorting, filtering, and styled formatting. The range should include a header row by default.',
    params: {
      address: { type: 'string', description: 'Range address for the table (e.g., "A1:D10")' },
      hasHeaders: {
        type: 'boolean',
        required: false,
        description: 'Whether the first row contains headers. Default true.',
      },
      name: { type: 'string', required: false, description: 'Optional name for the table.' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const hasHeaders = (args.hasHeaders as boolean) ?? true;
      const table = sheet.tables.add(args.address as string, hasHeaders);
      if (args.name) table.name = args.name as string;
      table.load(['name', 'id']);
      await context.sync();
      return { name: table.name, id: table.id, address: args.address };
    },
  },

  {
    name: 'add_table_rows',
    description:
      'Append new rows to the bottom of an existing Excel table. Each row must have the same number of columns as the table.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to add rows to' },
      values: { type: 'any[][]', description: '2D array of row data to add' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const values = args.values as unknown[][];
      table.rows.add(undefined, values as (string | number | boolean)[][]);
      table.load('name');
      await context.sync();
      return { tableName: table.name, rowsAdded: values.length };
    },
  },

  {
    name: 'get_table_data',
    description:
      'Read all data from an Excel table. Returns the column headers and all row values as a 2D array. Use list_tables first to discover available table names.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to read' },
    },
    execute: async (context, args) => {
      const tableName = args.tableName as string;
      const table = context.workbook.tables.getItem(tableName);
      const headerRange = table.getHeaderRowRange();
      const bodyRange = table.getDataBodyRange();
      headerRange.load('values');
      bodyRange.load(['values', 'rowCount', 'columnCount']);
      await context.sync();
      return {
        tableName,
        headers: headerRange.values[0],
        data: bodyRange.values,
        rowCount: bodyRange.rowCount,
        columnCount: bodyRange.columnCount,
      };
    },
  },

  {
    name: 'delete_table',
    description:
      'Remove the Excel table structure (sorting, filtering, styling) from a range. The underlying cell data is preserved — only the table object is removed.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to delete' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      table.delete();
      await context.sync();
      return { deleted: args.tableName };
    },
  },

  {
    name: 'sort_table',
    description:
      'Sort a structured Excel table by a column in ascending or descending order. For ad-hoc data ranges that are not tables, use sort_range instead.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to sort' },
      column: {
        type: 'number',
        description: 'Zero-based column index to sort by (0 = first column)',
      },
      ascending: {
        type: 'boolean',
        required: false,
        description: 'Sort ascending (true) or descending (false). Default true.',
      },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const column = args.column as number;
      const ascending = (args.ascending as boolean) ?? true;
      table.sort.apply([{ key: column, ascending }]);
      table.load('name');
      await context.sync();
      return { tableName: table.name, sortedByColumn: column, ascending };
    },
  },

  {
    name: 'filter_table',
    description:
      "Apply a values filter to a table column, hiding rows that don't match. Only rows with values matching any of the specified filter values will be visible. Use clear_table_filters to remove all filters.",
    params: {
      tableName: { type: 'string', description: 'Name of the table to filter' },
      column: {
        type: 'number',
        description: 'Zero-based column index to filter (0 = first column)',
      },
      values: {
        type: 'string[]',
        description: 'Array of values to show (rows matching any of these are kept)',
      },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const column = args.column as number;
      const values = args.values as string[];
      const col = table.columns.getItemAt(column);
      col.filter.applyValuesFilter(values);
      table.load('name');
      await context.sync();
      return { tableName: table.name, filteredColumn: column, filterValues: values };
    },
  },

  {
    name: 'clear_table_filters',
    description: 'Clear all filters from an Excel table, showing all rows.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to clear filters from' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      table.clearFilters();
      table.load('name');
      await context.sync();
      return { tableName: table.name, filtersCleared: true };
    },
  },

  // ─── Table Columns ────────────────────────────────────────

  {
    name: 'add_table_column',
    description: 'Add a new column to the end of a table.',
    params: {
      tableName: { type: 'string', description: 'Name of the table' },
      columnName: {
        type: 'string',
        required: false,
        description: 'Optional column header name. If omitted, Excel generates a default name.',
      },
      columnData: {
        type: 'string[]',
        required: false,
        description: 'Optional array of values for the new column (one per data row)',
      },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const columnData = args.columnData as string[] | undefined;
      const values = columnData ? columnData.map(v => [v]) : undefined;
      const col = table.columns.add(undefined, values, args.columnName as string | undefined);
      col.load('name');
      await context.sync();
      return { tableName: args.tableName, columnName: col.name, added: true };
    },
  },

  {
    name: 'delete_table_column',
    description: 'Delete a column from a table.',
    params: {
      tableName: { type: 'string', description: 'Name of the table' },
      columnName: { type: 'string', description: 'Name of the column to delete' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const col = table.columns.getItem(args.columnName as string);
      col.delete();
      await context.sync();
      return { tableName: args.tableName, columnName: args.columnName, deleted: true };
    },
  },

  {
    name: 'convert_table_to_range',
    description: 'Convert a structured table back to a plain range (removes table formatting).',
    params: {
      tableName: { type: 'string', description: 'Name of the table to convert' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      const range = table.convertToRange();
      range.load('address');
      await context.sync();
      return { tableName: args.tableName, rangeAddress: range.address, converted: true };
    },
  },

  {
    name: 'resize_table',
    description:
      'Resize a table to a new range address. The new range must overlap the existing table and keep a valid table shape.',
    params: {
      tableName: { type: 'string', description: 'Name of the table to resize' },
      newAddress: {
        type: 'string',
        description: 'New table range address (e.g., "A1:F200")',
      },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      table.resize(args.newAddress as string);
      const range = table.getRange();
      range.load('address');
      await context.sync();
      return { tableName: args.tableName, address: range.address, resized: true };
    },
  },

  {
    name: 'set_table_style',
    description: 'Set or change a table style (e.g., "TableStyleMedium2").',
    params: {
      tableName: { type: 'string', description: 'Name of the table' },
      style: { type: 'string', description: 'Table style name' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      table.style = args.style as string;
      table.load(['name', 'style']);
      await context.sync();
      return { tableName: table.name, style: table.style };
    },
  },

  {
    name: 'set_table_header_totals_visibility',
    description: 'Show or hide table header row and totals row.',
    params: {
      tableName: { type: 'string', description: 'Name of the table' },
      showHeaders: {
        type: 'boolean',
        required: false,
        description: 'Set table header row visibility',
      },
      showTotals: {
        type: 'boolean',
        required: false,
        description: 'Set table totals row visibility',
      },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      if (args.showHeaders !== undefined) table.showHeaders = args.showHeaders as boolean;
      if (args.showTotals !== undefined) table.showTotals = args.showTotals as boolean;
      table.load(['name', 'showHeaders', 'showTotals']);
      await context.sync();
      return {
        tableName: table.name,
        showHeaders: table.showHeaders,
        showTotals: table.showTotals,
      };
    },
  },

  {
    name: 'reapply_table_filters',
    description: 'Reapply existing filters on a table after data changes.',
    params: {
      tableName: { type: 'string', description: 'Name of the table' },
    },
    execute: async (context, args) => {
      const table = context.workbook.tables.getItem(args.tableName as string);
      table.reapplyFilters();
      await context.sync();
      return { tableName: args.tableName, reapplied: true };
    },
  },
];
