import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

function requiredString(value: unknown, name: string): string {
  if (typeof value !== 'string' || !value.trim())
    throw new Error(`${name} must be a nonempty string.`);
  return value;
}

export const slicerConfigs: readonly ToolConfig[] = [
  {
    name: 'slicer',
    description:
      'Manage Excel table and PivotTable slicers. Actions: list, create, get_info (items and selected keys), select_items, clear_filters, configure (caption/style/position/size), delete. Requires ExcelApi 1.10.',
    params: {
      action: {
        type: 'string',
        enum: [
          'list',
          'create',
          'get_info',
          'select_items',
          'clear_filters',
          'configure',
          'delete',
        ],
        description: 'Operation to perform.',
      },
      slicerName: {
        type: 'string',
        required: false,
        description: 'Existing slicer name or ID. Required except list/create.',
      },
      sourceType: {
        type: 'string',
        required: false,
        enum: ['table', 'pivot'],
        description: 'Source type for create.',
      },
      sourceName: {
        type: 'string',
        required: false,
        description: 'Table or PivotTable name for create.',
      },
      sourceField: {
        type: 'string',
        required: false,
        description: 'Table column name or PivotTable hierarchy/field name for create.',
      },
      sheetName: {
        type: 'string',
        required: false,
        description:
          'Destination sheet for create (default active); optional sheet scope for list.',
      },
      name: { type: 'string', required: false, description: 'New slicer name for create.' },
      itemKeys: {
        type: 'string[]',
        required: false,
        description:
          'Nonempty item keys from get_info for select_items. Replaces existing selection. Use clear_filters to select all.',
      },
      caption: { type: 'string', required: false, description: 'Caption for create/configure.' },
      style: {
        type: 'string',
        required: false,
        description: 'Excel slicer style name for create/configure.',
      },
      left: {
        type: 'number',
        required: false,
        description: 'Left position in points, nonnegative (create/configure).',
      },
      top: {
        type: 'number',
        required: false,
        description: 'Top position in points, nonnegative (create/configure).',
      },
      width: {
        type: 'number',
        required: false,
        description: 'Width in points, positive (create/configure).',
      },
      height: {
        type: 'number',
        required: false,
        description: 'Height in points, positive (create/configure).',
      },
    },
    execute: async (context, args) => {
      if (!Office.context.requirements.isSetSupported('ExcelApi', '1.10')) {
        throw new Error(
          'Slicers require ExcelApi 1.10 or later. This Excel version does not support them.'
        );
      }
      const action = args.action;
      if (action === 'list') {
        const collection = args.sheetName
          ? getSheet(context, requiredString(args.sheetName, 'sheetName')).slicers
          : context.workbook.slicers;
        collection.load('items/name,items/id,items/caption,items/isFilterCleared');
        await context.sync();
        for (const slicer of collection.items) slicer.worksheet.load('name');
        await context.sync();
        return {
          count: collection.items.length,
          slicers: collection.items.map(s => ({
            name: s.name,
            id: s.id,
            caption: s.caption,
            sheetName: s.worksheet.name,
            isFilterCleared: s.isFilterCleared,
          })),
        };
      }
      if (
        !['create', 'get_info', 'select_items', 'clear_filters', 'configure', 'delete'].includes(
          String(action)
        )
      ) {
        throw new Error('Unsupported slicer action.');
      }
      if (action === 'create' || action === 'configure') {
        for (const key of ['left', 'top', 'width', 'height'] as const) {
          const value = args[key];
          if (
            value !== undefined &&
            (typeof value !== 'number' ||
              !Number.isFinite(value) ||
              (key === 'left' || key === 'top' ? value < 0 : value <= 0))
          ) {
            throw new Error(
              `${key} must be a finite ${key === 'left' || key === 'top' ? 'nonnegative' : 'positive'} number of points.`
            );
          }
        }
        for (const key of ['caption', 'style', 'name'] as const) {
          if (args[key] !== undefined) requiredString(args[key], key);
        }
      }
      let slicer: Excel.Slicer;
      if (action === 'create') {
        const sourceName = requiredString(args.sourceName, 'sourceName');
        const sourceField = requiredString(args.sourceField, 'sourceField');
        const sheet = getSheet(context, args.sheetName as string | undefined);
        if (args.sourceType === 'table') {
          const table = context.workbook.tables.getItem(sourceName);
          slicer = sheet.slicers.add(table, table.columns.getItem(sourceField));
        } else if (args.sourceType === 'pivot') {
          const pivot = context.workbook.pivotTables.getItem(sourceName);
          slicer = sheet.slicers.add(
            pivot,
            pivot.hierarchies.getItem(sourceField).fields.getItem(sourceField)
          );
        } else {
          throw new Error('create requires sourceType table or pivot.');
        }
        if (args.name !== undefined) slicer.name = args.name as string;
      } else {
        slicer = context.workbook.slicers.getItem(requiredString(args.slicerName, 'slicerName'));
      }
      if (action === 'delete') {
        slicer.delete();
        await context.sync();
        return { slicerName: args.slicerName, deleted: true };
      }
      if (action === 'create' || action === 'configure') {
        if (args.caption !== undefined) slicer.caption = args.caption as string;
        if (args.style !== undefined) slicer.style = args.style as string;
        for (const key of ['left', 'top', 'width', 'height'] as const) {
          if (args[key] !== undefined) slicer[key] = args[key] as number;
        }
      }
      if (action === 'select_items') {
        if (
          !Array.isArray(args.itemKeys) ||
          !args.itemKeys.length ||
          !args.itemKeys.every(k => typeof k === 'string' && k.length > 0) ||
          new Set(args.itemKeys).size !== args.itemKeys.length
        ) {
          throw new Error('select_items requires nonempty unique itemKeys from get_info.');
        }
        slicer.slicerItems.load('items/key');
        await context.sync();
        const keys = new Set(slicer.slicerItems.items.map(item => item.key));
        if (args.itemKeys.some(key => !keys.has(key)))
          throw new Error('Unknown slicer item key. Refresh get_info before selecting.');
        slicer.selectItems(args.itemKeys);
      }
      if (action === 'clear_filters') slicer.clearFilters();
      slicer.load('name,id,caption,style,left,top,width,height,isFilterCleared');
      slicer.worksheet.load('name');
      slicer.slicerItems.load('items/key,items/name,items/isSelected,items/hasData');
      const selected = slicer.getSelectedItems();
      await context.sync();
      return {
        name: slicer.name,
        id: slicer.id,
        caption: slicer.caption,
        style: slicer.style,
        sheetName: slicer.worksheet.name,
        left: slicer.left,
        top: slicer.top,
        width: slicer.width,
        height: slicer.height,
        isFilterCleared: slicer.isFilterCleared,
        selectedKeys: selected.value,
        items: slicer.slicerItems.items.map(item => ({
          key: item.key,
          name: item.name,
          isSelected: item.isSelected,
          hasData: item.hasData,
        })),
      };
    },
  },
];
