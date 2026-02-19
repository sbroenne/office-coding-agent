/**
 * Table tool configs — 1 tool (table) with actions:
 * list, create, delete, get_data, add_rows, sort, filter,
 * clear_filters, reapply_filters, add_column, delete_column,
 * convert_to_range, resize, configure.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const tableConfigs: readonly ToolConfig[] = [
  {
    name: 'table',
    description:
      'Manage structured Excel tables. Actions: "list" (list tables), "create" (convert range to table), "delete" (remove table), "get_data" (read rows/headers), "add_rows" (append rows), "sort", "filter", "clear_filters", "reapply_filters", "add_column", "delete_column", "convert_to_range", "resize", "configure" (style/headers/totals).',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'list',
          'create',
          'delete',
          'get_data',
          'add_rows',
          'sort',
          'filter',
          'clear_filters',
          'reapply_filters',
          'add_column',
          'delete_column',
          'convert_to_range',
          'resize',
          'configure',
        ],
      },
      tableName: {
        type: 'string',
        required: false,
        description: 'Table name. Required for most actions except list/create.',
      },
      // create
      address: { type: 'string', required: false, description: 'Range address for create/resize.' },
      hasHeaders: {
        type: 'boolean',
        required: false,
        description: 'First row has headers (create). Default true.',
      },
      name: { type: 'string', required: false, description: 'Name for new table (create).' },
      // add_rows
      values: { type: 'any[][]', required: false, description: '2D array of row data (add_rows).' },
      // sort
      column: {
        type: 'number',
        required: false,
        description: 'Zero-based column index for sort/filter.',
      },
      ascending: { type: 'boolean', required: false, description: 'Ascending sort. Default true.' },
      // filter
      filterValues: { type: 'string[]', required: false, description: 'Values to show (filter).' },
      // add_column
      columnName: {
        type: 'string',
        required: false,
        description: 'Header name for new column (add_column/delete_column).',
      },
      columnData: {
        type: 'string[]',
        required: false,
        description: 'Values for new column (add_column).',
      },
      // configure
      style: {
        type: 'string',
        required: false,
        description: 'Table style name (configure), e.g. "TableStyleMedium2".',
      },
      showHeaders: {
        type: 'boolean',
        required: false,
        description: 'Show header row (configure).',
      },
      showTotals: { type: 'boolean', required: false, description: 'Show totals row (configure).' },
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name (used for list/create).',
      },
    },
    execute: async (context, args) => {
      const action = args.action as string;

      if (action === 'list') {
        const sheetName = args.sheetName as string | undefined;
        let tables: Excel.TableCollection;
        if (sheetName) {
          tables = getSheet(context, sheetName).tables;
        } else {
          tables = context.workbook.tables;
        }
        tables.load(['items/name', 'items/id']);
        await context.sync();
        const rangeProxies: Excel.Range[] = [];
        const bodyProxies: Excel.Range[] = [];
        for (const t of tables.items) {
          t.load(['name', 'id']);
          const r = t.getRange();
          r.load('address');
          rangeProxies.push(r);
          t.worksheet.load('name');
          t.columns.load('items/name');
          const b = t.getDataBodyRange();
          b.load('rowCount');
          bodyProxies.push(b);
        }
        await context.sync();
        const result = tables.items.map((t, i) => ({
          name: t.name,
          id: t.id,
          address: rangeProxies[i].address,
          worksheetName: t.worksheet.name,
          columns: t.columns.items.map(c => c.name),
          rowCount: bodyProxies[i].rowCount,
        }));
        return { tables: result, count: result.length };
      }

      if (action === 'create') {
        const sheet = getSheet(context, args.sheetName as string | undefined);
        const hasHeaders = (args.hasHeaders as boolean) ?? true;
        const t = sheet.tables.add(args.address as string, hasHeaders);
        if (args.name) t.name = args.name as string;
        t.load(['name', 'id']);
        await context.sync();
        return { name: t.name, id: t.id, address: args.address };
      }

      const tableName = args.tableName as string;
      const table = context.workbook.tables.getItem(tableName);

      if (action === 'delete') {
        table.delete();
        await context.sync();
        return { deleted: tableName };
      }

      if (action === 'get_data') {
        const header = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        header.load('values');
        body.load(['values', 'rowCount', 'columnCount']);
        await context.sync();
        return {
          tableName,
          headers: header.values[0],
          data: body.values,
          rowCount: body.rowCount,
          columnCount: body.columnCount,
        };
      }

      if (action === 'add_rows') {
        const values = args.values as unknown[][];
        table.rows.add(undefined, values as (string | number | boolean)[][]);
        table.load('name');
        await context.sync();
        return { tableName: table.name, rowsAdded: values.length };
      }

      if (action === 'sort') {
        const col = args.column as number;
        const asc = (args.ascending as boolean) ?? true;
        table.sort.apply([{ key: col, ascending: asc }]);
        table.load('name');
        await context.sync();
        return { tableName: table.name, sortedByColumn: col, ascending: asc };
      }

      if (action === 'filter') {
        const col = args.column as number;
        const values = args.filterValues as string[];
        table.columns.getItemAt(col).filter.applyValuesFilter(values);
        table.load('name');
        await context.sync();
        return { tableName: table.name, filteredColumn: col, filterValues: values };
      }

      if (action === 'clear_filters') {
        table.clearFilters();
        table.load('name');
        await context.sync();
        return { tableName: table.name, filtersCleared: true };
      }

      if (action === 'reapply_filters') {
        table.reapplyFilters();
        await context.sync();
        return { tableName, reapplied: true };
      }

      if (action === 'add_column') {
        const data = args.columnData as string[] | undefined;
        const vals = data ? data.map(v => [v]) : undefined;
        const col = table.columns.add(undefined, vals, args.columnName as string | undefined);
        col.load('name');
        await context.sync();
        return { tableName, columnName: col.name, added: true };
      }

      if (action === 'delete_column') {
        table.columns.getItem(args.columnName as string).delete();
        await context.sync();
        return { tableName, columnName: args.columnName, deleted: true };
      }

      if (action === 'convert_to_range') {
        const range = table.convertToRange();
        range.load('address');
        await context.sync();
        return { tableName, rangeAddress: range.address, converted: true };
      }

      if (action === 'resize') {
        table.resize(args.address as string);
        const range = table.getRange();
        range.load('address');
        await context.sync();
        return { tableName, address: range.address, resized: true };
      }

      // configure
      if (args.style !== undefined) table.style = args.style as string;
      if (args.showHeaders !== undefined) table.showHeaders = args.showHeaders as boolean;
      if (args.showTotals !== undefined) table.showTotals = args.showTotals as boolean;
      table.load(['name', 'style', 'showHeaders', 'showTotals']);
      await context.sync();
      return {
        tableName: table.name,
        style: table.style,
        showHeaders: table.showHeaders,
        showTotals: table.showTotals,
      };
    },
  },
];
