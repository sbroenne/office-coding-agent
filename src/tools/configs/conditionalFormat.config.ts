/**
 * Conditional format tool configs — 1 tool (conditional_format) with actions: add, list, clear.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const conditionalFormatConfigs: readonly ToolConfig[] = [
  {
    name: 'conditional_format',
    description:
      'Manage conditional formatting rules. Use action "add" to create a rule (specify type), "list" to show existing rules, or "clear" to remove all rules from a range.',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['add', 'list', 'clear'],
      },
      address: {
        type: 'string',
        required: false,
        description: 'Range address (e.g. "A1:D20"). Required for add/clear; optional for list.',
      },
      type: {
        type: 'string',
        required: false,
        description: 'Rule type for action=add.',
        enum: ['colorScale', 'dataBar', 'cellValue', 'topBottom', 'containsText', 'custom'],
      },
      // colorScale
      minColor: { type: 'string', required: false, description: 'Min color hex (colorScale).' },
      midColor: {
        type: 'string',
        required: false,
        description: 'Mid color hex (colorScale, optional).',
      },
      maxColor: { type: 'string', required: false, description: 'Max color hex (colorScale).' },
      // dataBar
      fillColor: { type: 'string', required: false, description: 'Fill color hex (dataBar).' },
      showDataBarOnly: {
        type: 'boolean',
        required: false,
        description: 'Hide cell value (dataBar).',
      },
      // cellValue
      operator: {
        type: 'string',
        required: false,
        description: 'Comparison operator (cellValue).',
        enum: [
          'Between',
          'NotBetween',
          'EqualTo',
          'NotEqualTo',
          'GreaterThan',
          'LessThan',
          'GreaterThanOrEqualTo',
          'LessThanOrEqualTo',
        ],
      },
      formula1: {
        type: 'string',
        required: false,
        description: 'First value/formula (cellValue/custom).',
      },
      formula2: {
        type: 'string',
        required: false,
        description: 'Second value for Between/NotBetween (cellValue).',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color hex for matching cells.',
      },
      backgroundColor: {
        type: 'string',
        required: false,
        description: 'Background color hex for matching cells.',
      },
      // topBottom
      topBottomRank: {
        type: 'number',
        required: false,
        description: 'Number of items (topBottom).',
      },
      topBottomType: {
        type: 'string',
        required: false,
        description: 'Top or bottom, absolute or percent (topBottom).',
        enum: ['TopItems', 'TopPercent', 'BottomItems', 'BottomPercent'],
      },
      // containsText
      containsText: {
        type: 'string',
        required: false,
        description: 'Text to match (containsText).',
      },
      containsOperator: {
        type: 'string',
        required: false,
        description: 'Match mode (containsText).',
        enum: ['Contains', 'NotContains', 'BeginsWith', 'EndsWith'],
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const action = args.action as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);

      if (action === 'list') {
        const addr = args.address as string | undefined;
        const range = addr ? sheet.getRange(addr) : sheet.getUsedRange();
        const formats = range.conditionalFormats;
        formats.load('items');
        await context.sync();
        for (const f of formats.items) {
          f.load(['type', 'priority', 'stopIfTrue']);
        }
        await context.sync();
        const result = formats.items.map(f => ({
          type: f.type,
          priority: f.priority,
          stopIfTrue: f.stopIfTrue,
        }));
        return { conditionalFormats: result, count: result.length };
      }

      if (action === 'clear') {
        const range = sheet.getRange(args.address as string);
        range.conditionalFormats.clearAll();
        await context.sync();
        return { address: args.address, cleared: true };
      }

      // action === 'add'
      const range = sheet.getRange(args.address as string);
      const type = args.type as string;

      if (type === 'colorScale') {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const cs = cf.colorScale;
        cs.criteria = {
          minimum: {
            color: args.minColor as string,
            type: Excel.ConditionalFormatColorCriterionType.lowestValue,
          },
          maximum: {
            color: args.maxColor as string,
            type: Excel.ConditionalFormatColorCriterionType.highestValue,
          },
          ...(args.midColor
            ? {
                midpoint: {
                  color: args.midColor as string,
                  type: Excel.ConditionalFormatColorCriterionType.percentile,
                  formula: '50',
                },
              }
            : {}),
        };
        await context.sync();
        return { action, type, address: args.address, added: true };
      }

      if (type === 'dataBar') {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        const db = cf.dataBar;
        if (args.fillColor) db.positiveFormat.fillColor = args.fillColor as string;
        if (args.showDataBarOnly !== undefined)
          db.showDataBarOnly = args.showDataBarOnly as boolean;
        await context.sync();
        return { action, type, address: args.address, added: true };
      }

      if (type === 'cellValue') {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        const cv = cf.cellValue;
        cv.rule = {
          operator: args.operator as Excel.ConditionalCellValueOperator,
          formula1: String(args.formula1),
          ...(args.formula2 ? { formula2: args.formula2 as string } : {}),
        };
        if (args.fontColor) cv.format.font.color = args.fontColor as string;
        if (args.backgroundColor) cv.format.fill.color = args.backgroundColor as string;
        await context.sync();
        return { action, type, address: args.address, added: true };
      }

      if (type === 'topBottom') {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
        const tb = cf.topBottom;
        tb.rule = {
          rank: args.topBottomRank as number,
          type: args.topBottomType as Excel.ConditionalTopBottomCriterionType,
        };
        if (args.fontColor) tb.format.font.color = args.fontColor as string;
        if (args.backgroundColor) tb.format.fill.color = args.backgroundColor as string;
        await context.sync();
        return { action, type, address: args.address, added: true };
      }

      if (type === 'containsText') {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        const ct = cf.textComparison;
        ct.rule = {
          operator: (args.containsOperator as Excel.ConditionalTextOperator) ?? 'Contains',
          text: args.containsText as string,
        };
        if (args.fontColor) ct.format.font.color = args.fontColor as string;
        if (args.backgroundColor) ct.format.fill.color = args.backgroundColor as string;
        await context.sync();
        return { action, type, address: args.address, added: true };
      }

      // custom
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      const custom = cf.custom;
      custom.rule.formula = args.formula1 as string;
      if (args.fontColor) custom.format.font.color = args.fontColor as string;
      if (args.backgroundColor) custom.format.fill.color = args.backgroundColor as string;
      await context.sync();
      return { action, type, address: args.address, added: true };
    },
  },
];
