/**
 * PivotTable tool configs — 7 tools for managing PivotTables.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

function getPivotField(pt: Excel.PivotTable, fieldName: string): Excel.PivotField {
  const hierarchy = pt.hierarchies.getItem(fieldName);
  return hierarchy.fields.getItem(fieldName);
}

async function resolveDataHierarchy(
  context: Excel.RequestContext,
  pt: Excel.PivotTable,
  requestedName: string
): Promise<Excel.DataPivotHierarchy> {
  pt.dataHierarchies.load('items/name');
  await context.sync();

  const exact = pt.dataHierarchies.items.find(
    item => item.name.toLowerCase() === requestedName.toLowerCase()
  );
  if (exact) return pt.dataHierarchies.getItem(exact.name);

  const contains = pt.dataHierarchies.items.find(
    item =>
      item.name.toLowerCase().includes(requestedName.toLowerCase()) ||
      requestedName.toLowerCase().includes(item.name.toLowerCase())
  );
  if (contains) return pt.dataHierarchies.getItem(contains.name);

  if (pt.dataHierarchies.items.length > 0) {
    return pt.dataHierarchies.getItem(pt.dataHierarchies.items[0].name);
  }

  return pt.dataHierarchies.getItem(requestedName);
}

export const pivotTableConfigs: readonly ToolConfig[] = [
  {
    name: 'list_pivot_tables',
    description:
      "List all PivotTables on a worksheet. Returns each PivotTable's name plus row/column/filter/data hierarchy names.",
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
        pt.columnHierarchies.load('items/name');
        pt.filterHierarchies.load('items/name');
        pt.dataHierarchies.load('items/name');
      }
      await context.sync();

      const result = pivotTables.items.map(pt => ({
        name: pt.name,
        id: pt.id,
        rowHierarchies: pt.rowHierarchies.items.map(h => h.name),
        columnHierarchies: pt.columnHierarchies.items.map(h => h.name),
        filterHierarchies: pt.filterHierarchies.items.map(h => h.name),
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
    name: 'refresh_all_pivot_tables',
    description:
      'Refresh all PivotTables. If sheetName is provided, refreshes PivotTables on that worksheet; otherwise refreshes all workbook PivotTables.',
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name to scope refresh to one sheet.',
      },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      if (sheetName) {
        const sheet = getSheet(context, sheetName);
        sheet.pivotTables.refreshAll();
      } else {
        context.workbook.pivotTables.refreshAll();
      }
      await context.sync();
      return { refreshed: true, sheetName: sheetName ?? null };
    },
  },

  {
    name: 'get_pivot_table_source_info',
    description:
      'Get PivotTable source metadata, including source type and source address/connection string when available.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTableName = args.pivotTableName as string;
      const pt = sheet.pivotTables.getItem(pivotTableName);

      const sourceTypeResult = pt.getDataSourceType();
      const sourceStringResult = pt.getDataSourceString();
      await context.sync();

      return {
        pivotTableName,
        dataSourceType: sourceTypeResult.value,
        dataSourceString: sourceStringResult.value,
      };
    },
  },

  {
    name: 'get_pivot_hierarchy_counts',
    description:
      'Get counts of row, column, filter, and data hierarchies currently configured on a PivotTable.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);

      const rowCount = pt.rowHierarchies.getCount();
      const columnCount = pt.columnHierarchies.getCount();
      const filterCount = pt.filterHierarchies.getCount();
      const dataCount = pt.dataHierarchies.getCount();
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        rowHierarchyCount: rowCount.value,
        columnHierarchyCount: columnCount.value,
        filterHierarchyCount: filterCount.value,
        dataHierarchyCount: dataCount.value,
      };
    },
  },

  {
    name: 'get_pivot_hierarchies',
    description:
      'List row, column, filter, and data hierarchies for a PivotTable including names and ids.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);

      pt.rowHierarchies.load('items/name,id');
      pt.columnHierarchies.load('items/name,id');
      pt.filterHierarchies.load('items/name,id');
      pt.dataHierarchies.load('items/name,id');
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        rowHierarchies: pt.rowHierarchies.items.map(h => ({ name: h.name, id: h.id })),
        columnHierarchies: pt.columnHierarchies.items.map(h => ({ name: h.name, id: h.id })),
        filterHierarchies: pt.filterHierarchies.items.map(h => ({ name: h.name, id: h.id })),
        dataHierarchies: pt.dataHierarchies.items.map(h => ({ name: h.name, id: h.id })),
      };
    },
  },

  {
    name: 'set_pivot_table_options',
    description:
      'Set PivotTable behavior options such as multiple filters per field, custom sort lists, refresh on open, and data value editing.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      allowMultipleFiltersPerField: {
        type: 'boolean',
        required: false,
        description: 'Allow multiple filters (for different filter types) on the same PivotField',
      },
      useCustomSortLists: {
        type: 'boolean',
        required: false,
        description: 'Use workbook custom sort lists when sorting PivotFields',
      },
      refreshOnOpen: {
        type: 'boolean',
        required: false,
        description: 'Refresh this PivotTable when the workbook opens',
      },
      enableDataValueEditing: {
        type: 'boolean',
        required: false,
        description: 'Enable editing of data values directly in the PivotTable when supported',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTableName = args.pivotTableName as string;
      const pt = sheet.pivotTables.getItem(pivotTableName);

      if (args.allowMultipleFiltersPerField !== undefined) {
        pt.allowMultipleFiltersPerField = args.allowMultipleFiltersPerField as boolean;
      }
      if (args.useCustomSortLists !== undefined) {
        pt.useCustomSortLists = args.useCustomSortLists as boolean;
      }
      if (args.refreshOnOpen !== undefined) {
        pt.refreshOnOpen = args.refreshOnOpen as boolean;
      }
      if (args.enableDataValueEditing !== undefined) {
        pt.enableDataValueEditing = args.enableDataValueEditing as boolean;
      }

      pt.load([
        'allowMultipleFiltersPerField',
        'useCustomSortLists',
        'refreshOnOpen',
        'enableDataValueEditing',
      ]);
      await context.sync();

      return {
        pivotTableName,
        allowMultipleFiltersPerField: pt.allowMultipleFiltersPerField,
        useCustomSortLists: pt.useCustomSortLists,
        refreshOnOpen: pt.refreshOnOpen,
        enableDataValueEditing: pt.enableDataValueEditing,
        updated: true,
      };
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
    name: 'set_pivot_layout',
    description:
      'Set PivotTable layout and display options such as layout type, subtotal placement, field headers, and grand totals.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable to configure' },
      layoutType: {
        type: 'string',
        required: false,
        description: 'Pivot layout type',
        enum: ['Compact', 'Tabular', 'Outline'],
      },
      subtotalLocation: {
        type: 'string',
        required: false,
        description: 'Subtotal location for row fields',
        enum: ['AtTop', 'AtBottom', 'Off'],
      },
      showFieldHeaders: {
        type: 'boolean',
        required: false,
        description: 'Show or hide pivot field headers',
      },
      showRowGrandTotals: {
        type: 'boolean',
        required: false,
        description: 'Show or hide row grand totals',
      },
      showColumnGrandTotals: {
        type: 'boolean',
        required: false,
        description: 'Show or hide column grand totals',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;

      if (args.layoutType !== undefined) {
        layout.layoutType = args.layoutType as Excel.PivotLayoutType;
      }
      if (args.subtotalLocation !== undefined) {
        layout.subtotalLocation = args.subtotalLocation as Excel.SubtotalLocationType;
      }
      if (args.showFieldHeaders !== undefined) {
        layout.showFieldHeaders = args.showFieldHeaders as boolean;
      }
      if (args.showRowGrandTotals !== undefined) {
        layout.showRowGrandTotals = args.showRowGrandTotals as boolean;
      }
      if (args.showColumnGrandTotals !== undefined) {
        layout.showColumnGrandTotals = args.showColumnGrandTotals as boolean;
      }

      layout.load([
        'layoutType',
        'subtotalLocation',
        'showFieldHeaders',
        'showRowGrandTotals',
        'showColumnGrandTotals',
      ]);
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        layoutType: layout.layoutType,
        subtotalLocation: layout.subtotalLocation,
        showFieldHeaders: layout.showFieldHeaders,
        showRowGrandTotals: layout.showRowGrandTotals,
        showColumnGrandTotals: layout.showColumnGrandTotals,
        updated: true,
      };
    },
  },

  {
    name: 'get_pivot_field_filters',
    description:
      'Get active filter state for a PivotField, including whether any/date/label/manual/value filters are currently set.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      const filtersResult = field.getFilters();
      const hasAnyFilterResult = field.isFiltered();
      await context.sync();

      const filters = filtersResult.value;
      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        hasAnyFilter: hasAnyFilterResult.value,
        hasDateFilter: filters.dateFilter !== undefined,
        hasLabelFilter: filters.labelFilter !== undefined,
        hasManualFilter: filters.manualFilter !== undefined,
        hasValueFilter: filters.valueFilter !== undefined,
      };
    },
  },

  {
    name: 'get_pivot_field_items',
    description:
      'List PivotItems for a PivotField, including item names, ids, and expand state.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      field.items.load('items/name,id,isExpanded');
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        items: field.items.items.map(item => ({
          name: item.name,
          id: item.id,
          isExpanded: item.isExpanded,
        })),
        count: field.items.items.length,
      };
    },
  },

  {
    name: 'clear_pivot_field_filters',
    description:
      'Clear filters on a PivotField. If filterType is omitted, clears all filters; otherwise clears only the specified filter type.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      filterType: {
        type: 'string',
        required: false,
        description: 'Optional filter type to clear; omit to clear all filter types.',
        enum: ['Value', 'Manual', 'Label', 'Date'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      if (args.filterType === undefined) {
        field.clearAllFilters();
      } else {
        field.clearFilter(args.filterType as Excel.PivotFilterType);
      }

      await context.sync();
      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        filterType: (args.filterType as string | undefined) ?? null,
        cleared: true,
      };
    },
  },

  {
    name: 'apply_pivot_label_filter',
    description:
      'Apply a label filter to a PivotField. Provide condition and value1; optionally provide value2 for Between/NotBetween conditions.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      condition: {
        type: 'string',
        description: 'Label filter condition',
        enum: [
          'Equals',
          'DoesNotEqual',
          'BeginsWith',
          'DoesNotBeginWith',
          'EndsWith',
          'DoesNotEndWith',
          'Contains',
          'DoesNotContain',
          'GreaterThan',
          'GreaterThanOrEqualTo',
          'LessThan',
          'LessThanOrEqualTo',
          'Between',
          'NotBetween',
        ],
      },
      value1: { type: 'string', description: 'Primary comparator value' },
      value2: {
        type: 'string',
        required: false,
        description: 'Optional secondary comparator value for Between/NotBetween conditions',
      },
      substring: {
        type: 'string',
        required: false,
        description: 'Optional substring for contains/begins/ends conditions',
      },
      exclusive: {
        type: 'boolean',
        required: false,
        description: 'Optional flag for exclusive boundary behavior where supported',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      const labelFilter: Excel.PivotLabelFilter = {
        condition: args.condition as Excel.LabelFilterCondition,
        comparator: (args.value1 as string) ?? '',
      };

      if (args.value2 !== undefined) {
        labelFilter.lowerBound = args.value1 as string;
        labelFilter.upperBound = args.value2 as string;
      }
      if (args.substring !== undefined) {
        labelFilter.substring = args.substring as string;
      }
      if (args.exclusive !== undefined) {
        labelFilter.exclusive = args.exclusive as boolean;
      }

      field.applyFilter({ labelFilter });
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        condition: args.condition,
        value1: args.value1,
        value2: (args.value2 as string | undefined) ?? null,
        applied: true,
      };
    },
  },

  {
    name: 'sort_pivot_field_labels',
    description: 'Sort PivotField labels in ascending or descending order.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      sortBy: {
        type: 'string',
        description: 'Sort direction for labels',
        enum: ['Ascending', 'Descending'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      field.sortByLabels(args.sortBy as Excel.SortBy);
      await context.sync();
      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        sortBy: args.sortBy,
        sorted: true,
      };
    },
  },

  {
    name: 'apply_pivot_manual_filter',
    description:
      'Apply a manual item filter to a PivotField by explicitly selecting visible item names.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      selectedItems: {
        type: 'string[]',
        description: 'Pivot item names to keep visible for this field',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);
      const selectedItems = args.selectedItems as string[];

      field.items.load('items/name');
      await context.sync();

      const selectedSet = new Set(selectedItems.map(v => v.toLowerCase()));
      for (const item of field.items.items) {
        item.visible = selectedSet.has(item.name.toLowerCase());
      }
      await context.sync();
      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        selectedItems,
        applied: true,
      };
    },
  },

  {
    name: 'sort_pivot_field_values',
    description:
      'Sort a PivotField by data values for a specified value hierarchy (measure). Optionally scope sort to a pivot item.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      sortBy: {
        type: 'string',
        description: 'Sort direction for value-based sort',
        enum: ['Ascending', 'Descending'],
      },
      valuesHierarchyName: {
        type: 'string',
        description: 'Name of the data hierarchy/measure to sort by',
      },
      pivotItemScope: {
        type: 'string',
        required: false,
        description: 'Optional pivot item name to scope value sort',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);
      const valuesHierarchy = await resolveDataHierarchy(
        context,
        pt,
        args.valuesHierarchyName as string
      );
      const sortBy = args.sortBy as Excel.SortBy;
      const pivotItemScopeName = args.pivotItemScope as string | undefined;

      if (pivotItemScopeName) {
        const scopeItem = field.items.getItem(pivotItemScopeName);
        field.sortByValues(sortBy, valuesHierarchy, [scopeItem]);
      } else {
        field.sortByValues(sortBy, valuesHierarchy);
      }

      await context.sync();
      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        sortBy: args.sortBy,
        valuesHierarchyName: args.valuesHierarchyName,
        pivotItemScope: pivotItemScopeName ?? null,
        sorted: true,
      };
    },
  },

  {
    name: 'set_pivot_field_show_all_items',
    description: 'Set whether a PivotField shows all items, including those with no data.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      fieldName: { type: 'string', description: 'Name of the PivotField (source column name)' },
      showAllItems: {
        type: 'boolean',
        description: 'Whether to show all items for the field',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const field = getPivotField(pt, args.fieldName as string);

      field.showAllItems = args.showAllItems as boolean;
      field.load('showAllItems');
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        fieldName: args.fieldName,
        showAllItems: field.showAllItems,
        updated: true,
      };
    },
  },

  {
    name: 'get_pivot_layout_ranges',
    description:
      'Get key PivotLayout range addresses (table, row labels, column labels, data body, and filter axis when available).',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;

      const tableRange = layout.getRange();
      const rowLabelRange = layout.getRowLabelRange();
      const columnLabelRange = layout.getColumnLabelRange();
      const dataBodyRange = layout.getDataBodyRange();

      tableRange.load('address');
      rowLabelRange.load('address');
      columnLabelRange.load('address');
      dataBodyRange.load('address');
      await context.sync();

      let filterAxisAddress: string | null = null;
      try {
        const filterAxisRange = layout.getFilterAxisRange();
        filterAxisRange.load('address');
        await context.sync();
        filterAxisAddress = filterAxisRange.address;
      } catch {
        filterAxisAddress = null;
      }

      return {
        pivotTableName: args.pivotTableName,
        tableRangeAddress: tableRange.address,
        rowLabelRangeAddress: rowLabelRange.address,
        columnLabelRangeAddress: columnLabelRange.address,
        dataBodyRangeAddress: dataBodyRange.address,
        filterAxisRangeAddress: filterAxisAddress,
      };
    },
  },

  {
    name: 'set_pivot_layout_display_options',
    description:
      'Configure PivotLayout display and formatting behavior such as repeat labels, blank lines, auto-format, empty-cell text, and accessibility text.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      repeatAllItemLabels: {
        type: 'boolean',
        required: false,
        description: 'Repeat item labels across all pivot fields',
      },
      displayBlankLineAfterEachItem: {
        type: 'boolean',
        required: false,
        description: 'Display a blank line after each pivot item across fields',
      },
      autoFormat: {
        type: 'boolean',
        required: false,
        description: 'Auto-apply formatting when refreshed or fields move',
      },
      preserveFormatting: {
        type: 'boolean',
        required: false,
        description: 'Preserve formatting on refresh/recalculate',
      },
      fillEmptyCells: {
        type: 'boolean',
        required: false,
        description: 'Fill empty pivot cells using emptyCellText',
      },
      emptyCellText: {
        type: 'string',
        required: false,
        description: 'Text used when fillEmptyCells is enabled',
      },
      enableFieldList: {
        type: 'boolean',
        required: false,
        description: 'Enable/disable field list in UI',
      },
      altTextTitle: {
        type: 'string',
        required: false,
        description: 'Accessibility alt text title for the PivotTable',
      },
      altTextDescription: {
        type: 'string',
        required: false,
        description: 'Accessibility alt text description for the PivotTable',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;

      if (args.repeatAllItemLabels !== undefined) {
        layout.repeatAllItemLabels(args.repeatAllItemLabels as boolean);
      }
      if (args.displayBlankLineAfterEachItem !== undefined) {
        layout.displayBlankLineAfterEachItem(args.displayBlankLineAfterEachItem as boolean);
      }
      if (args.autoFormat !== undefined) {
        layout.autoFormat = args.autoFormat as boolean;
      }
      if (args.preserveFormatting !== undefined) {
        layout.preserveFormatting = args.preserveFormatting as boolean;
      }
      if (args.fillEmptyCells !== undefined) {
        layout.fillEmptyCells = args.fillEmptyCells as boolean;
      }
      if (args.emptyCellText !== undefined) {
        layout.emptyCellText = args.emptyCellText as string;
      }
      if (args.enableFieldList !== undefined) {
        layout.enableFieldList = args.enableFieldList as boolean;
      }
      if (args.altTextTitle !== undefined) {
        layout.altTextTitle = args.altTextTitle as string;
      }
      if (args.altTextDescription !== undefined) {
        layout.altTextDescription = args.altTextDescription as string;
      }

      layout.load([
        'autoFormat',
        'preserveFormatting',
        'fillEmptyCells',
        'emptyCellText',
        'enableFieldList',
        'altTextTitle',
        'altTextDescription',
      ]);
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        autoFormat: layout.autoFormat,
        preserveFormatting: layout.preserveFormatting,
        fillEmptyCells: layout.fillEmptyCells,
        emptyCellText: layout.emptyCellText,
        enableFieldList: layout.enableFieldList,
        altTextTitle: layout.altTextTitle,
        altTextDescription: layout.altTextDescription,
        updated: true,
      };
    },
  },

  {
    name: 'get_pivot_data_hierarchy_for_cell',
    description:
      'Get the DataPivotHierarchy (measure) used to compute a specific cell in the PivotTable data body.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      cellAddress: {
        type: 'string',
        description: 'Single cell address inside the PivotTable data body (e.g., "B5")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;
      const hierarchy = layout.getDataHierarchy(args.cellAddress as string);

      hierarchy.load(['name', 'id']);
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        cellAddress: args.cellAddress,
        dataHierarchyName: hierarchy.name,
        dataHierarchyId: hierarchy.id,
      };
    },
  },

  {
    name: 'get_pivot_items_for_cell',
    description:
      'Get PivotItems from a specified axis that contribute to a PivotTable data-body cell value.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      axis: {
        type: 'string',
        description: 'Pivot axis to inspect for contributing items',
        enum: ['Row', 'Column', 'Data', 'Filter'],
      },
      cellAddress: {
        type: 'string',
        description: 'Single cell address inside the PivotTable data body (e.g., "B5")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;
      const axis = args.axis as 'Row' | 'Column' | 'Data' | 'Filter';
      const items = layout.getPivotItems(axis, args.cellAddress as string);

      items.load('items/name,id,isExpanded');
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        axis,
        cellAddress: args.cellAddress,
        items: items.items.map(item => ({
          name: item.name,
          id: item.id,
          isExpanded: item.isExpanded,
        })),
        count: items.items.length,
      };
    },
  },

  {
    name: 'set_pivot_layout_auto_sort_on_cell',
    description:
      'Apply PivotTable autosort using a specific pivot data-body cell as context (equivalent to UI autosort behavior).',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      cellAddress: {
        type: 'string',
        description: 'Single cell address inside the PivotTable data body used as autosort context',
      },
      sortBy: {
        type: 'string',
        description: 'Sort direction for autosort',
        enum: ['Ascending', 'Descending'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pt = sheet.pivotTables.getItem(args.pivotTableName as string);
      const layout = pt.layout;

      layout.setAutoSortOnCell(args.cellAddress as string, args.sortBy as Excel.SortBy);
      await context.sync();

      return {
        pivotTableName: args.pivotTableName,
        cellAddress: args.cellAddress,
        sortBy: args.sortBy,
        sorted: true,
      };
    },
  },

  {
    name: 'get_pivot_table_count',
    description:
      'Get the PivotTable count for a worksheet or for the entire workbook when no sheet is provided.',
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name. If omitted, counts PivotTables in the workbook.',
      },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      const count = sheetName
        ? getSheet(context, sheetName).pivotTables.getCount()
        : context.workbook.pivotTables.getCount();
      await context.sync();

      return { count: count.value, scope: sheetName ?? 'workbook' };
    },
  },

  {
    name: 'pivot_table_exists',
    description:
      'Check whether a PivotTable exists by name on a worksheet or across workbook PivotTables.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable to check' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      const pivotTableName = args.pivotTableName as string;

      const collection = sheetName
        ? getSheet(context, sheetName).pivotTables
        : context.workbook.pivotTables;
      const pt = collection.getItemOrNullObject(pivotTableName);
      pt.load('isNullObject,name,id');
      await context.sync();

      return {
        pivotTableName,
        exists: !pt.isNullObject,
        sheetName: sheetName ?? null,
        id: pt.isNullObject ? null : pt.id,
        resolvedName: pt.isNullObject ? null : pt.name,
      };
    },
  },

  {
    name: 'get_pivot_table_location',
    description:
      'Get PivotTable worksheet name and layout range address for a named PivotTable on a worksheet.',
    params: {
      pivotTableName: { type: 'string', description: 'Name of the PivotTable' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      const pivotTableName = args.pivotTableName as string;
      const sheet = getSheet(context, sheetName);
      const pt = sheet.pivotTables.getItem(pivotTableName);
      const tableRange = pt.layout.getRange();

      pt.worksheet.load('name');
      tableRange.load('address');
      await context.sync();

      return {
        pivotTableName,
        worksheetName: pt.worksheet.name,
        rangeAddress: tableRange.address,
      };
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
