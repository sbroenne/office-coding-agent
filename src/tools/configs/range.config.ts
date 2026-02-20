/**
 * Range tool configs — 1 tool (range) for data and manipulation operations.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const rangeConfigs: readonly ToolConfig[] = [
  {
    name: 'range',
    description:
      'Operate on cell ranges. Actions: "get_values", "set_values", "get_formulas", "set_formulas", "get_used" (used range + optional paged values), "clear", "copy", "sort", "fill" (auto-fill), "flash_fill", "find", "replace", "get_special_cells", "get_precedents", "get_dependents", "get_tables", "remove_duplicates", "merge", "unmerge", "group", "ungroup", "insert", "delete", "recalculate".',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'get_values',
          'set_values',
          'get_formulas',
          'set_formulas',
          'get_used',
          'clear',
          'copy',
          'sort',
          'fill',
          'flash_fill',
          'find',
          'replace',
          'get_special_cells',
          'get_precedents',
          'get_dependents',
          'get_tables',
          'remove_duplicates',
          'merge',
          'unmerge',
          'group',
          'ungroup',
          'insert',
          'delete',
          'recalculate',
        ],
      },
      address: {
        type: 'string',
        required: false,
        description: 'Range address (e.g. "A1:C10"). Required for most actions.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
      // paging (get_values / get_used)
      maxRows: {
        type: 'number',
        required: false,
        description: 'Max rows to return (get_values/get_used).',
      },
      maxColumns: { type: 'number', required: false, description: 'Max columns to return.' },
      startRow: {
        type: 'number',
        required: false,
        description: '1-based starting row offset. Default 1.',
      },
      startColumn: {
        type: 'number',
        required: false,
        description: '1-based starting column offset. Default 1.',
      },
      // set_values
      values: { type: 'any[][]', required: false, description: '2D array of values (set_values).' },
      // set_formulas
      formulas: {
        type: 'string[][]',
        required: false,
        description: '2D array of formula strings (set_formulas).',
      },
      // sort
      column: {
        type: 'number',
        required: false,
        description: 'Zero-based column index to sort by.',
      },
      ascending: { type: 'boolean', required: false, description: 'Ascending sort. Default true.' },
      hasHeaders: {
        type: 'boolean',
        required: false,
        description: 'First row is header (sort). Default true.',
      },
      // copy
      destinationAddress: { type: 'string', required: false, description: 'Destination for copy.' },
      sourceSheet: { type: 'string', required: false, description: 'Source sheet name (copy).' },
      destinationSheet: {
        type: 'string',
        required: false,
        description: 'Destination sheet name (copy).',
      },
      // fill
      sourceAddress: { type: 'string', required: false, description: 'Source range for fill.' },
      autoFillType: {
        type: 'string',
        required: false,
        description: 'Auto-fill behavior (fill). Default FillDefault.',
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
      // find
      searchText: { type: 'string', required: false, description: 'Text to find (find).' },
      matchCase: {
        type: 'boolean',
        required: false,
        description: 'Case-sensitive (find). Default false.',
      },
      matchEntireCell: {
        type: 'boolean',
        required: false,
        description: 'Match entire cell (find). Default false.',
      },
      // replace
      find: { type: 'string', required: false, description: 'Text to find (replace).' },
      replace: { type: 'string', required: false, description: 'Replacement text (replace).' },
      // get_special_cells
      cellType: {
        type: 'string',
        required: false,
        description: 'Special cell type (get_special_cells).',
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
        description: 'Value filter for get_special_cells.',
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
      // get_tables
      fullyContained: {
        type: 'boolean',
        required: false,
        description: 'Only tables fully inside range (get_tables). Default false.',
      },
      // remove_duplicates
      columns: {
        type: 'string[]',
        required: false,
        description: 'Zero-based column indices to compare (remove_duplicates).',
      },
      // merge
      across: {
        type: 'boolean',
        required: false,
        description: 'Merge each row separately (merge). Default false.',
      },
      // insert / delete
      shift: {
        type: 'string',
        required: false,
        description:
          'Shift direction for insert ("down","right") or delete ("up","left"). Default "down"/"up".',
        enum: ['down', 'right', 'up', 'left'],
      },
    },
    execute: async (context, args) => {
      const action = args.action as string;

      if (action === 'get_used') {
        const sheet = getSheet(context, args.sheetName as string | undefined);
        const usedRange = sheet.getUsedRange();
        const maxRows = args.maxRows as number | undefined;
        const maxCols = args.maxColumns as number | undefined;
        const startRow = Math.max(1, (args.startRow as number | undefined) ?? 1);
        const startCol = Math.max(1, (args.startColumn as number | undefined) ?? 1);
        const includeValues = maxRows != null || startRow > 1 || startCol > 1 || maxCols != null;
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
          const rowStart = startRow - 1;
          const colStart = startCol - 1;
          const rowEnd =
            maxRows != null ? Math.min(rowStart + maxRows, usedRange.rowCount) : usedRange.rowCount;
          const colEnd =
            maxCols != null
              ? Math.min(colStart + maxCols, usedRange.columnCount)
              : usedRange.columnCount;
          result.values = usedRange.values
            .slice(rowStart, rowEnd)
            .map((row: unknown[]) => row.slice(colStart, colEnd));
          result.rowsReturned = (result.values as unknown[][]).length;
          result.columnsReturned = (result.values as unknown[][])[0]?.length ?? 0;
          if (rowEnd < usedRange.rowCount) {
            result.moreRows = true;
            result.nextStartRow = rowEnd + 1;
          }
          if (colEnd < usedRange.columnCount) {
            result.moreColumns = true;
            result.nextStartColumn = colEnd + 1;
          }
        }
        return result;
      }

      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;

      if (action === 'get_values') {
        const range = sheet.getRange(address);
        range.load(['values', 'address', 'rowCount', 'columnCount']);
        await context.sync();
        const startRow = Math.max(1, (args.startRow as number | undefined) ?? 1);
        const startCol = Math.max(1, (args.startColumn as number | undefined) ?? 1);
        const maxRows = args.maxRows as number | undefined;
        const maxCols = args.maxColumns as number | undefined;
        const rowStart = startRow - 1;
        const colStart = startCol - 1;
        const rowEnd =
          maxRows != null ? Math.min(rowStart + maxRows, range.rowCount) : range.rowCount;
        const colEnd =
          maxCols != null ? Math.min(colStart + maxCols, range.columnCount) : range.columnCount;
        const sliced = range.values
          .slice(rowStart, rowEnd)
          .map((row: unknown[]) => row.slice(colStart, colEnd));
        const result: Record<string, unknown> = {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          values: sliced,
        };
        const isPaged = maxRows != null || maxCols != null || startRow > 1 || startCol > 1;
        if (isPaged) {
          result.rowsReturned = sliced.length;
          result.columnsReturned = sliced[0]?.length ?? 0;
          if (rowEnd < range.rowCount) {
            result.moreRows = true;
            result.nextStartRow = rowEnd + 1;
          }
          if (colEnd < range.columnCount) {
            result.moreColumns = true;
            result.nextStartColumn = colEnd + 1;
          }
        }
        return result;
      }

      if (action === 'set_values') {
        const range = sheet.getRange(address);
        const values = args.values as unknown[][];
        range.values = values;
        range.load('address');
        await context.sync();
        return {
          address: range.address,
          rowsWritten: values.length,
          columnsWritten: values[0]?.length ?? 0,
        };
      }

      if (action === 'get_formulas') {
        const range = sheet.getRange(address);
        range.load(['formulas', 'address', 'rowCount', 'columnCount']);
        await context.sync();
        return {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          formulas: range.formulas,
        };
      }

      if (action === 'set_formulas') {
        const range = sheet.getRange(address);
        const formulas = args.formulas as string[][];
        range.formulas = formulas;
        range.load('address');
        await context.sync();
        return {
          address: range.address,
          rowsWritten: formulas.length,
          columnsWritten: formulas[0]?.length ?? 0,
        };
      }

      if (action === 'clear') {
        const range = sheet.getRange(address);
        range.clear(Excel.ClearApplyTo.all);
        range.load('address');
        await context.sync();
        return { address: range.address, cleared: true };
      }

      if (action === 'copy') {
        const srcSheet = getSheet(context, args.sourceSheet as string | undefined);
        const dstSheet = getSheet(context, args.destinationSheet as string | undefined);
        const src = srcSheet.getRange(address);
        const dst = dstSheet.getRange(args.destinationAddress as string);
        dst.copyFrom(src);
        src.load('address');
        dst.load('address');
        await context.sync();
        return { source: src.address, destination: dst.address, copied: true };
      }

      if (action === 'sort') {
        const range = sheet.getRange(address);
        const col = args.column as number;
        const asc = (args.ascending as boolean) ?? true;
        const hdrs = (args.hasHeaders as boolean) ?? true;
        range.sort.apply([{ key: col, ascending: asc }], hdrs);
        range.load('address');
        await context.sync();
        return { address: range.address, sortedByColumn: col, ascending: asc };
      }

      if (action === 'fill') {
        const srcSheet = getSheet(context, args.sheetName as string | undefined);
        const src = srcSheet.getRange(args.sourceAddress as string);
        const fillType = ((args.autoFillType as string | undefined) ??
          'FillDefault') as Excel.AutoFillType;
        src.autoFill(args.destinationAddress as string, fillType);
        const dst = srcSheet.getRange(args.destinationAddress as string);
        dst.load('address');
        await context.sync();
        return {
          sourceAddress: args.sourceAddress,
          destinationAddress: dst.address,
          autoFillType: fillType,
          filled: true,
        };
      }

      if (action === 'flash_fill') {
        const range = sheet.getRange(address);
        range.flashFill();
        range.load('address');
        await context.sync();
        return { address: range.address, flashFilled: true };
      }

      if (action === 'find') {
        const usedRange = sheet.getUsedRange();
        const result = usedRange.find(args.searchText as string, {
          completeMatch: (args.matchEntireCell as boolean) ?? false,
          matchCase: (args.matchCase as boolean) ?? false,
        });
        result.load(['address', 'values']);
        try {
          await context.sync();
          return {
            found: true,
            address: result.address,
            value: result.values[0]?.[0] as string | number | boolean | null,
          };
        } catch {
          return { found: false, searchText: args.searchText, message: 'No match found.' };
        }
      }

      if (action === 'replace') {
        const range = args.address ? sheet.getRange(address) : sheet.getUsedRange();
        const count = range.replaceAll(args.find as string, args.replace as string, {
          completeMatch: false,
          matchCase: false,
        });
        await context.sync();
        return { find: args.find, replace: args.replace, replacements: count.value };
      }

      if (action === 'get_special_cells') {
        const range = sheet.getRange(address);
        const special = range.getSpecialCells(
          args.cellType as Excel.SpecialCellType,
          args.cellValueType as Excel.SpecialCellValueType | undefined
        );
        special.load(['address', 'cellCount', 'areaCount']);
        await context.sync();
        return {
          sourceAddress: address,
          specialAddress: special.address,
          areaCount: special.areaCount,
          cellCount: special.cellCount,
        };
      }

      if (action === 'get_precedents') {
        const range = sheet.getRange(address);
        const prec = range.getDirectPrecedents();
        prec.load('addresses');
        await context.sync();
        return { sourceAddress: address, addresses: prec.addresses, count: prec.addresses.length };
      }

      if (action === 'get_dependents') {
        const range = sheet.getRange(address);
        const dep = range.getDirectDependents();
        dep.load('addresses');
        await context.sync();
        return { sourceAddress: address, addresses: dep.addresses, count: dep.addresses.length };
      }

      if (action === 'get_tables') {
        const range = sheet.getRange(address);
        const fullyContained = (args.fullyContained as boolean | undefined) ?? false;
        const tables = range.getTables(fullyContained);
        tables.load('items');
        await context.sync();
        for (const t of tables.items) t.load(['name', 'id']);
        await context.sync();
        return {
          address,
          fullyContained,
          tables: tables.items.map(t => ({ name: t.name, id: t.id })),
          count: tables.items.length,
        };
      }

      if (action === 'remove_duplicates') {
        const range = sheet.getRange(address);
        const cols = (args.columns as string[]).map(Number);
        const result = range.removeDuplicates(cols, true);
        result.load(['removed', 'uniqueRemaining']);
        await context.sync();
        return { address, rowsRemoved: result.removed, rowsRemaining: result.uniqueRemaining };
      }

      if (action === 'merge') {
        const range = sheet.getRange(address);
        const across = (args.across as boolean) ?? false;
        range.merge(across);
        range.load('address');
        await context.sync();
        return { address: range.address, merged: true, across };
      }

      if (action === 'unmerge') {
        const range = sheet.getRange(address);
        range.unmerge();
        range.load('address');
        await context.sync();
        return { address: range.address, unmerged: true };
      }

      if (action === 'group') {
        const range = sheet.getRange(address);
        range.group('ByRows');
        await context.sync();
        return { address, grouped: true };
      }

      if (action === 'ungroup') {
        const range = sheet.getRange(address);
        range.ungroup('ByRows');
        await context.sync();
        return { address, ungrouped: true };
      }

      if (action === 'insert') {
        const range = sheet.getRange(address);
        const shiftStr = (args.shift as string) ?? 'down';
        range.insert(
          shiftStr === 'right' ? Excel.InsertShiftDirection.right : Excel.InsertShiftDirection.down
        );
        await context.sync();
        return { address, shift: shiftStr, inserted: true };
      }

      if (action === 'delete') {
        const range = sheet.getRange(address);
        const shiftStr = (args.shift as string) ?? 'up';
        range.delete(
          shiftStr === 'left' ? Excel.DeleteShiftDirection.left : Excel.DeleteShiftDirection.up
        );
        await context.sync();
        return { address, shift: shiftStr, deleted: true };
      }

      // recalculate
      const range = sheet.getRange(address);
      range.calculate();
      await context.sync();
      return { address, recalculated: true };
    },
  },
];
