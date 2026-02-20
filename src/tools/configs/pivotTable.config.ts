/**
 * PivotTable tool configs — 1 tool (pivot) with actions:
 * list, create, delete, get_info, refresh, configure,
 * add_field, remove_field, filter, sort.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

function getPivotField(pt: Excel.PivotTable, fieldName: string): Excel.PivotField {
  return pt.hierarchies.getItem(fieldName).fields.getItem(fieldName);
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
  if (pt.dataHierarchies.items.length > 0)
    return pt.dataHierarchies.getItem(pt.dataHierarchies.items[0].name);
  return pt.dataHierarchies.getItem(requestedName);
}

export const pivotTableConfigs: readonly ToolConfig[] = [
  {
    name: 'pivot',
    description:
      'Manage PivotTables. Actions: "list" (all PivotTables on a sheet), "create" (new PivotTable), "delete", "get_info" (source, hierarchies), "refresh", "configure" (layout/options), "add_field" (add to row/column/data/filter), "remove_field", "filter" (apply/clear filters on a field), "sort" (by labels or values).',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'list',
          'create',
          'delete',
          'get_info',
          'refresh',
          'configure',
          'add_field',
          'remove_field',
          'filter',
          'sort',
        ],
      },
      pivotTableName: {
        type: 'string',
        required: false,
        description: 'PivotTable name. Required for most actions except list/create.',
      },
      // create
      name: { type: 'string', required: false, description: 'Name for new PivotTable (create).' },
      sourceAddress: {
        type: 'string',
        required: false,
        description: 'Source data range with headers, e.g. "Sheet1!A1:D100" (create).',
      },
      destinationAddress: {
        type: 'string',
        required: false,
        description: 'Top-left destination cell, e.g. "Sheet2!A1" (create).',
      },
      rowFields: {
        type: 'string[]',
        required: false,
        description: 'Column names for row labels (create).',
      },
      valueFields: {
        type: 'string[]',
        required: false,
        description: 'Column names to aggregate (create).',
      },
      sourceSheetName: {
        type: 'string',
        required: false,
        description: 'Source data sheet (create).',
      },
      destinationSheetName: {
        type: 'string',
        required: false,
        description: 'Destination sheet (create).',
      },
      // add_field / remove_field
      fieldName: {
        type: 'string',
        required: false,
        description: 'Source column name for field operations.',
      },
      fieldType: {
        type: 'string',
        required: false,
        description: 'Field area (add_field/remove_field).',
        enum: ['row', 'column', 'data', 'filter'],
      },
      // configure
      layoutType: {
        type: 'string',
        required: false,
        enum: ['Compact', 'Tabular', 'Outline'],
        description: 'Layout type (configure).',
      },
      subtotalLocation: {
        type: 'string',
        required: false,
        enum: ['AtTop', 'AtBottom', 'Off'],
        description: 'Subtotal location (configure).',
      },
      showFieldHeaders: {
        type: 'boolean',
        required: false,
        description: 'Show field headers (configure).',
      },
      showRowGrandTotals: {
        type: 'boolean',
        required: false,
        description: 'Show row grand totals (configure).',
      },
      showColumnGrandTotals: {
        type: 'boolean',
        required: false,
        description: 'Show column grand totals (configure).',
      },
      allowMultipleFiltersPerField: {
        type: 'boolean',
        required: false,
        description: 'Allow multiple filters per field (configure).',
      },
      useCustomSortLists: {
        type: 'boolean',
        required: false,
        description: 'Use custom sort lists for sorting (configure).',
      },
      refreshOnOpen: {
        type: 'boolean',
        required: false,
        description: 'Refresh on open (configure).',
      },
      // filter
      filterType: {
        type: 'string',
        required: false,
        description:
          'For filter action: "label" (label filter), "manual" (item selection), or "clear" (clear filters).',
        enum: ['label', 'manual', 'clear'],
      },
      // label filter
      labelCondition: {
        type: 'string',
        required: false,
        description: 'Label filter condition (filter action, filterType=label).',
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
      labelValue1: {
        type: 'string',
        required: false,
        description: 'Primary comparator for label filter.',
      },
      labelValue2: {
        type: 'string',
        required: false,
        description: 'Second comparator for Between/NotBetween.',
      },
      // manual filter
      selectedItems: {
        type: 'string[]',
        required: false,
        description: 'Items to show for manual filter.',
      },
      // sort
      sortBy: {
        type: 'string',
        required: false,
        enum: ['Ascending', 'Descending'],
        description: 'Sort direction (sort action).',
      },
      sortMode: {
        type: 'string',
        required: false,
        enum: ['labels', 'values'],
        description: 'Sort by labels or by values (sort action).',
      },
      valuesHierarchyName: {
        type: 'string',
        required: false,
        description: 'Data hierarchy name for value-based sort.',
      },
      pivotItemScope: {
        type: 'string',
        required: false,
        description: 'Pivot item scope for value sort.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const action = args.action as string;

      if (action === 'list') {
        const sheet = getSheet(context, args.sheetName as string | undefined);
        const pivotTables = sheet.pivotTables;
        pivotTables.load('items');
        await context.sync();
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
      }

      if (action === 'create') {
        const sourceSheet = getSheet(context, args.sourceSheetName as string | undefined);
        const destSheet = args.destinationSheetName
          ? context.workbook.worksheets.getItem(args.destinationSheetName as string)
          : sourceSheet;
        const srcRange = sourceSheet.getRange(args.sourceAddress as string);
        const dstRange = destSheet.getRange(args.destinationAddress as string);
        const pt = context.workbook.pivotTables.add(args.name as string, srcRange, dstRange);
        for (const field of (args.rowFields as string[]) ?? []) {
          pt.rowHierarchies.add(pt.hierarchies.getItem(field));
        }
        for (const field of (args.valueFields as string[]) ?? []) {
          const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(field));
          dh.summarizeBy = 'Sum' as Excel.AggregationFunction;
        }
        pt.load('name');
        await context.sync();
        return {
          pivotTableName: pt.name,
          rowFields: args.rowFields,
          valueFields: args.valueFields,
          created: true,
        };
      }

      const sheet = getSheet(context, args.sheetName as string | undefined);
      const pivotTableName = args.pivotTableName as string;
      const pt = sheet.pivotTables.getItem(pivotTableName);

      if (action === 'delete') {
        pt.delete();
        await context.sync();
        return { pivotTableName, deleted: true };
      }

      if (action === 'get_info') {
        const srcType = pt.getDataSourceType();
        const srcStr = pt.getDataSourceString();
        pt.rowHierarchies.load('items/name,id');
        pt.columnHierarchies.load('items/name,id');
        pt.filterHierarchies.load('items/name,id');
        pt.dataHierarchies.load('items/name,id');
        const rowCount = pt.rowHierarchies.getCount();
        const colCount = pt.columnHierarchies.getCount();
        const filterCount = pt.filterHierarchies.getCount();
        const dataCount = pt.dataHierarchies.getCount();
        await context.sync();
        return {
          pivotTableName,
          dataSourceType: srcType.value,
          dataSourceString: srcStr.value,
          rowHierarchyCount: rowCount.value,
          columnHierarchyCount: colCount.value,
          filterHierarchyCount: filterCount.value,
          dataHierarchyCount: dataCount.value,
          rowHierarchies: pt.rowHierarchies.items.map(h => ({ name: h.name, id: h.id })),
          columnHierarchies: pt.columnHierarchies.items.map(h => ({ name: h.name, id: h.id })),
          filterHierarchies: pt.filterHierarchies.items.map(h => ({ name: h.name, id: h.id })),
          dataHierarchies: pt.dataHierarchies.items.map(h => ({ name: h.name, id: h.id })),
        };
      }

      if (action === 'refresh') {
        pt.refresh();
        await context.sync();
        return { pivotTableName, refreshed: true };
      }

      if (action === 'configure') {
        const layout = pt.layout;
        if (args.layoutType !== undefined)
          layout.layoutType = args.layoutType as Excel.PivotLayoutType;
        if (args.subtotalLocation !== undefined)
          layout.subtotalLocation = args.subtotalLocation as Excel.SubtotalLocationType;
        if (args.showFieldHeaders !== undefined)
          layout.showFieldHeaders = args.showFieldHeaders as boolean;
        if (args.showRowGrandTotals !== undefined)
          layout.showRowGrandTotals = args.showRowGrandTotals as boolean;
        if (args.showColumnGrandTotals !== undefined)
          layout.showColumnGrandTotals = args.showColumnGrandTotals as boolean;
        if (args.allowMultipleFiltersPerField !== undefined)
          pt.allowMultipleFiltersPerField = args.allowMultipleFiltersPerField as boolean;
        if (args.useCustomSortLists !== undefined)
          pt.useCustomSortLists = args.useCustomSortLists as boolean;
        if (args.refreshOnOpen !== undefined) pt.refreshOnOpen = args.refreshOnOpen as boolean;
        layout.load([
          'layoutType',
          'subtotalLocation',
          'showFieldHeaders',
          'showRowGrandTotals',
          'showColumnGrandTotals',
        ]);
        pt.load(['name', 'allowMultipleFiltersPerField', 'useCustomSortLists']);
        await context.sync();
        return {
          pivotTableName,
          layoutType: layout.layoutType,
          subtotalLocation: layout.subtotalLocation,
          showFieldHeaders: layout.showFieldHeaders,
          showRowGrandTotals: layout.showRowGrandTotals,
          showColumnGrandTotals: layout.showColumnGrandTotals,
          allowMultipleFiltersPerField: pt.allowMultipleFiltersPerField,
          useCustomSortLists: pt.useCustomSortLists,
          updated: true,
        };
      }

      if (action === 'add_field') {
        const fieldName = args.fieldName as string;
        const fieldType = args.fieldType as string;
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
        return { pivotTableName, fieldName, fieldType, added: true };
      }

      if (action === 'remove_field') {
        const fieldName = args.fieldName as string;
        const fieldType = args.fieldType as string;
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
        return { pivotTableName, fieldName, fieldType, removed: true };
      }

      if (action === 'filter') {
        const field = getPivotField(pt, args.fieldName as string);
        const filterType = args.filterType as string;
        if (filterType === 'clear') {
          field.clearAllFilters();
          await context.sync();
          return { pivotTableName, fieldName: args.fieldName, cleared: true };
        }
        if (filterType === 'manual') {
          const selectedItems = args.selectedItems as string[];
          field.items.load('items/name');
          await context.sync();
          const selectedSet = new Set(selectedItems.map(v => v.toLowerCase()));
          for (const item of field.items.items) {
            item.visible = selectedSet.has(item.name.toLowerCase());
          }
          await context.sync();
          return { pivotTableName, fieldName: args.fieldName, selectedItems, applied: true };
        }
        // label filter
        const labelFilter: Excel.PivotLabelFilter = {
          condition: args.labelCondition as Excel.LabelFilterCondition,
          comparator: (args.labelValue1 as string) ?? '',
        };
        if (args.labelValue2 !== undefined) {
          labelFilter.lowerBound = args.labelValue1 as string;
          labelFilter.upperBound = args.labelValue2 as string;
        }
        field.applyFilter({ labelFilter });
        await context.sync();
        return { pivotTableName, fieldName: args.fieldName, filterType, applied: true };
      }

      // sort
      const field = getPivotField(pt, args.fieldName as string);
      const sortBy = args.sortBy as Excel.SortBy;
      const sortMode = args.sortMode as string;
      if (sortMode === 'values') {
        const valuesHierarchy = await resolveDataHierarchy(
          context,
          pt,
          args.valuesHierarchyName as string
        );
        const pivotItemScopeName = args.pivotItemScope as string | undefined;
        if (pivotItemScopeName) {
          field.sortByValues(sortBy, valuesHierarchy, [field.items.getItem(pivotItemScopeName)]);
        } else {
          field.sortByValues(sortBy, valuesHierarchy);
        }
      } else {
        field.sortByLabels(sortBy);
      }
      await context.sync();
      return { pivotTableName, fieldName: args.fieldName, sortBy, sortMode, sorted: true };
    },
  },
];
