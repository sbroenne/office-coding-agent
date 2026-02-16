/**
 * PivotTable tool configs — 6 tools for managing PivotTables.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const pivotTableConfigs: readonly ToolConfig[] = [
  {
    name: 'list_pivot_tables',
    description:
      "List all PivotTables on a worksheet. Returns each PivotTable's name, row hierarchies, and data hierarchies.",
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTables = sheet.pivotTables;
      pivotTables.load('items');
      await context.sync();

      // Load per-item details and hierarchy names (second sync)
      for (const pt of pivotTables.items) {
        pt.load(['name', 'id']);
        pt.rowHierarchies.load('items/name');
        pt.dataHierarchies.load('items/name');
      }
      await context.sync();

      const result = pivotTables.items.map(pt => ({
        name: pt.name,
        id: pt.id,
        rowHierarchies: pt.rowHierarchies.items.map(h => h.name),
        dataHierarchies: pt.dataHierarchies.items.map(h => h.name),
      }));
      return { pivotTables: result, count: result.length };
    },
  },

  {
    name: 'refresh_pivot_table',
    description: 'Refresh a PivotTable to reflect changes in its source data.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable to refresh' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTableName = args.pivotTableName as string;
      const pt = sheet.pivotTables.getItem(pivotTableName);
      pt.refresh();
      await context.sync();
      return { pivotTableName, refreshed: true };
    },
  },

  {
    name: 'delete_pivot_table',
    description: 'Delete a PivotTable from the worksheet.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable to delete' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTableName = args.pivotTableName as string;
      const pt = sheet.pivotTables.getItem(pivotTableName);
      pt.delete();
      await context.sync();
      return { pivotTableName, deleted: true };
    },
  },

  // ─── Create PivotTable ──────────────────────────────────

  {
    name: 'create_pivot_table',
    description:
      'Create a new PivotTable from a data range. Specify which fields go into rows and values. Value fields default to SUM aggregation.',
    params: {
      name: { type: 'string', description: 'Name for the new PivotTable' },
      sourceAddress: {
        type: 'string',
        description:
          'Source data range with headers (e.g., "Sheet1!A1:D100"). Must include column headers in the first row.',
      },
      destinationAddress: {
        type: 'string',
        description:
          'Top-left cell where the PivotTable should be placed (e.g., "Sheet2!A1"). Must be on a different area from the source.',
      },
      rowFields: {
        type: 'string[]',
        description: 'Column names to use as row labels (e.g., ["Region", "Category"])',
      },
      valueFields: {
        type: 'string[]',
        description: 'Column names to aggregate as values (e.g., ["Sales", "Quantity"])',
      },
      sourceSheetName: {
        type: 'string',
        required: false,
        description: 'Sheet containing the source data. Uses active sheet if omitted.',
      },
      destinationSheetName: {
        type: 'string',
        required: false,
        description: 'Sheet for the PivotTable output. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const sourceSheet = getSheet(context, args.sourceSheetName as string | undefined);
      const destSheet = args.destinationSheetName
        ? context.workbook.worksheets.getItem(args.destinationSheetName as string)
        : sourceSheet;
      const sourceRange = sourceSheet.getRange(args.sourceAddress as string);
      const destRange = destSheet.getRange(args.destinationAddress as string);
      const pivotTableName = args.name as string;
      const pt = context.workbook.pivotTables.add(pivotTableName, sourceRange, destRange);

      // Add row fields
      const rowFields = args.rowFields as string[];
      for (const field of rowFields) {
        pt.rowHierarchies.add(pt.hierarchies.getItem(field));
      }

      // Add value fields with SUM aggregation
      const valueFields = args.valueFields as string[];
      for (const field of valueFields) {
        const dataHierarchy = pt.dataHierarchies.add(pt.hierarchies.getItem(field));
        dataHierarchy.summarizeBy = 'Sum' as Excel.AggregationFunction;
      }

      pt.load('name');
      await context.sync();
      return {
        pivotTableName: pt.name,
        rowFields,
        valueFields,
        created: true,
      };
    },
  },
  // ─── Pivot Table Fields ───────────────────────────────────

  {
    name: 'add_pivot_field',
    description:
      'Add a field to a pivot table as a row field, column field, data field, or filter field.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: {
        type: 'string',
        description: 'Name of the source data column to add as a field',
      },
      fieldType: {
        type: 'string',
        description: 'Where to add the field in the pivot layout',
        enum: ['row', 'column', 'data', 'filter'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const fieldType = args.fieldType as string;
      const fieldName = args.fieldName as string;

      switch (fieldType) {
        case 'row':
          pt.rowHierarchies.add(pt.hierarchies.getItem(fieldName));
          break;
        case 'column':
          pt.columnHierarchies.add(pt.hierarchies.getItem(fieldName));
          break;
        case 'data':
          pt.dataHierarchies.add(pt.hierarchies.getItem(fieldName));
          break;
        case 'filter':
          pt.filterHierarchies.add(pt.hierarchies.getItem(fieldName));
          break;
      }

      await context.sync();
      return { pivotTableName: args.pivotTableName, fieldName, fieldType, added: true };
    },
  },

  {
    name: 'remove_pivot_field',
    description: 'Remove a field from a pivot table.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the field to remove' },
      fieldType: {
        type: 'string',
        description: 'Location of the field in the pivot layout',
        enum: ['row', 'column', 'data', 'filter'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const fieldType = args.fieldType as string;
      const fieldName = args.fieldName as string;

      switch (fieldType) {
        case 'row':
          pt.rowHierarchies.remove(pt.rowHierarchies.getItem(fieldName));
          break;
        case 'column':
          pt.columnHierarchies.remove(pt.columnHierarchies.getItem(fieldName));
          break;
        case 'data':
          pt.dataHierarchies.remove(pt.dataHierarchies.getItem(fieldName));
          break;
        case 'filter':
          pt.filterHierarchies.remove(pt.filterHierarchies.getItem(fieldName));
          break;
      }

      await context.sync();
      return { pivotTableName: args.pivotTableName, fieldName, fieldType, removed: true };
    },
  },
];
