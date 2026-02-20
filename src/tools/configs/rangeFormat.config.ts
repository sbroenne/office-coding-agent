/**
 * Range format tool configs — 1 tool (range_format) for formatting operations.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const rangeFormatConfigs: readonly ToolConfig[] = [
  {
    name: 'range_format',
    description:
      'Apply formatting to cell ranges. Actions: "format" (bold/italic/font/fill/alignment/wrap), "set_number_format" (number display formats), "auto_fit" (column/row widths), "set_borders", "set_hyperlink", "toggle_visibility" (hide/show rows or columns).',
    params: {
      action: {
        type: 'string',
        description: 'Formatting operation to perform',
        enum: [
          'format',
          'set_number_format',
          'auto_fit',
          'set_borders',
          'set_hyperlink',
          'toggle_visibility',
        ],
      },
      address: {
        type: 'string',
        required: false,
        description: 'Range address. Required for most actions.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
      // format
      bold: { type: 'boolean', required: false, description: 'Bold text.' },
      italic: { type: 'boolean', required: false, description: 'Italic text.' },
      underline: { type: 'boolean', required: false, description: 'Underline text.' },
      fontSize: { type: 'number', required: false, description: 'Font size in points.' },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color hex (e.g. "#FF0000").',
      },
      fillColor: { type: 'string', required: false, description: 'Cell background color hex.' },
      horizontalAlignment: {
        type: 'string',
        required: false,
        enum: ['General', 'Left', 'Center', 'Right', 'Fill', 'Justify', 'Distributed'],
        description: 'Horizontal alignment.',
      },
      wrapText: { type: 'boolean', required: false, description: 'Enable text wrap.' },
      // set_number_format
      format: {
        type: 'string',
        required: false,
        description: 'Excel number format string, e.g. "#,##0.00", "0.00%", "yyyy-mm-dd".',
      },
      // auto_fit
      fitTarget: {
        type: 'string',
        required: false,
        enum: ['columns', 'rows', 'both'],
        description: 'What to auto-fit (auto_fit). Default "columns".',
      },
      // set_borders
      borderStyle: {
        type: 'string',
        required: false,
        enum: ['Thin', 'Medium', 'Thick', 'Double', 'Dotted', 'Dashed', 'DashDot'],
        description: 'Border style.',
      },
      borderColor: { type: 'string', required: false, description: 'Border color hex.' },
      side: {
        type: 'string',
        required: false,
        enum: ['EdgeLeft', 'EdgeRight', 'EdgeTop', 'EdgeBottom', 'EdgeAll'],
        description: 'Which side(s) to apply border.',
      },
      // set_hyperlink
      url: {
        type: 'string',
        required: false,
        description: 'URL or cell ref for hyperlink. Use "" to remove.',
      },
      textToDisplay: { type: 'string', required: false, description: 'Hyperlink display text.' },
      tooltip: { type: 'string', required: false, description: 'Hyperlink tooltip.' },
      // toggle_visibility
      hidden: {
        type: 'boolean',
        required: false,
        description: 'True to hide (toggle_visibility).',
      },
      target: {
        type: 'string',
        required: false,
        enum: ['rows', 'columns'],
        description: 'Whether to hide rows or columns.',
      },
    },
    execute: async (context, args) => {
      const action = args.action as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string | undefined;
      const addr = address ?? '';

      if (action === 'format') {
        const range = sheet.getRange(addr);
        if (args.bold !== undefined) range.format.font.bold = args.bold as boolean;
        if (args.italic !== undefined) range.format.font.italic = args.italic as boolean;
        if (args.fontSize !== undefined) range.format.font.size = args.fontSize as number;
        if (args.fontColor !== undefined) range.format.font.color = args.fontColor as string;
        if (args.fillColor !== undefined) range.format.fill.color = args.fillColor as string;
        if (args.horizontalAlignment !== undefined)
          range.format.horizontalAlignment = args.horizontalAlignment as Excel.HorizontalAlignment;
        if (args.underline !== undefined)
          range.format.font.underline = (args.underline as boolean)
            ? Excel.RangeUnderlineStyle.single
            : Excel.RangeUnderlineStyle.none;
        if (args.wrapText !== undefined) range.format.wrapText = args.wrapText as boolean;
        range.load('address');
        await context.sync();
        return { address: range.address, formatted: true };
      }

      if (action === 'set_number_format') {
        const range = sheet.getRange(addr);
        range.numberFormat = [[args.format as string]];
        range.load(['address', 'rowCount', 'columnCount']);
        await context.sync();
        return {
          address: range.address,
          format: args.format,
          cellCount: range.rowCount * range.columnCount,
        };
      }

      if (action === 'auto_fit') {
        const range = address ? sheet.getRange(address) : sheet.getUsedRange();
        const fitTarget = (args.fitTarget as string) ?? 'columns';
        if (fitTarget === 'columns' || fitTarget === 'both') range.format.autofitColumns();
        if (fitTarget === 'rows' || fitTarget === 'both') range.format.autofitRows();
        range.load('address');
        await context.sync();
        return { address: range.address, autoFitted: true, fitTarget };
      }

      if (action === 'set_borders') {
        const range = sheet.getRange(addr);
        const borderStyle = args.borderStyle as string;
        const side = args.side as string;
        const color = (args.borderColor as string) ?? '000000';
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
          const border = range.format.borders.getItem(s);
          border.style = mapped.style as Excel.BorderLineStyle;
          if (mapped.weight) border.weight = mapped.weight as Excel.BorderWeight;
          border.color = color;
        }
        await context.sync();
        return { address, borderStyle, side, color };
      }

      if (action === 'set_hyperlink') {
        const range = sheet.getRange(addr);
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
      }

      // toggle_visibility
      const range = sheet.getRange(addr);
      if ((args.target as string) === 'rows') {
        range.rowHidden = args.hidden as boolean;
      } else {
        range.columnHidden = args.hidden as boolean;
      }
      await context.sync();
      return { address, target: args.target, hidden: args.hidden };
    },
  },
];
