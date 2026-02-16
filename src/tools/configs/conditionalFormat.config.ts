/**
 * Conditional format tool configs — 8 tools (decomposed from 3).
 *
 * The old mega-tool `add_conditional_format` with 6 ruleType branches is now
 * 6 single-purpose tools. Each tool has exactly the parameters it needs —
 * no silently-ignored params, no ambiguity for the LLM.
 *
 * Fixes applied (from tool audit):
 *   - fillColor/fontColor now applied correctly for all rule types
 *   - Each tool only exposes the parameters it actually uses
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const conditionalFormatConfigs: readonly ToolConfig[] = [
  // ─── 6 decomposed add tools ────────────────────────────

  {
    name: 'add_color_scale',
    description:
      'Add a color scale (gradient) conditional format to a range. Cells are colored on a gradient from minColor (lowest value) through optional midColor to maxColor (highest value).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "B2:B100")' },
      minColor: {
        type: 'string',
        required: false,
        description: 'Color for lowest values (e.g., "blue"). Default "blue".',
      },
      midColor: {
        type: 'string',
        required: false,
        description: 'Optional color for midpoint values (e.g., "yellow").',
      },
      maxColor: {
        type: 'string',
        required: false,
        description: 'Color for highest values (e.g., "red"). Default "red".',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
      const minColor = (args.minColor as string) ?? 'blue';
      const maxColor = (args.maxColor as string) ?? 'red';
      const midColor = args.midColor as string | undefined;
      const criteria: Excel.ConditionalColorScaleCriteria = {
        minimum: {
          formula: undefined,
          type: Excel.ConditionalFormatColorCriterionType.lowestValue,
          color: minColor,
        },
        maximum: {
          formula: undefined,
          type: Excel.ConditionalFormatColorCriterionType.highestValue,
          color: maxColor,
        },
      };
      if (midColor) {
        criteria.midpoint = {
          formula: '50',
          type: Excel.ConditionalFormatColorCriterionType.percent,
          color: midColor,
        };
      }
      cf.colorScale.criteria = criteria;
      await context.sync();
      return { address, ruleType: 'colorScale', applied: true };
    },
  },

  {
    name: 'add_data_bar',
    description:
      'Add data bars to cells, showing a horizontal bar whose length represents the cell value relative to other cells in the range.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "C2:C50")' },
      barColor: {
        type: 'string',
        required: false,
        description: 'Bar fill color (e.g., "#638EC6"). Default "#638EC6".',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
      const barColor = (args.barColor as string) ?? '#638EC6';
      cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
      cf.dataBar.positiveFormat.fillColor = barColor;
      await context.sync();
      return { address, ruleType: 'dataBar', applied: true };
    },
  },

  {
    name: 'add_cell_value_format',
    description:
      'Add conditional formatting based on cell numeric values. Highlight cells that meet a condition (e.g., greater than 100, between 0 and 50).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "B2:B100")' },
      operator: {
        type: 'string',
        description: 'Comparison operator',
        enum: [
          'GreaterThan',
          'GreaterThanOrEqual',
          'LessThan',
          'LessThanOrEqual',
          'EqualTo',
          'NotEqualTo',
          'Between',
          'NotBetween',
        ],
      },
      formula1: { type: 'string', description: 'First comparison value (e.g., "100", "=0")' },
      formula2: {
        type: 'string',
        required: false,
        description: 'Second value for Between/NotBetween operators.',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color when condition is met (e.g., "#FF0000").',
      },
      fillColor: {
        type: 'string',
        required: false,
        description: 'Fill color when condition is met (e.g., "#00FF00").',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      const fontColor = args.fontColor as string | undefined;
      const fillColor = args.fillColor as string | undefined;
      if (fontColor) cf.cellValue.format.font.color = fontColor;
      if (fillColor) cf.cellValue.format.fill.color = fillColor;
      const rule: Excel.ConditionalCellValueRule = {
        formula1: args.formula1 as string,
        operator: args.operator as Excel.ConditionalCellValueOperator,
      };
      if (args.formula2) rule.formula2 = args.formula2 as string;
      cf.cellValue.rule = rule;
      await context.sync();
      return { address, ruleType: 'cellValue', applied: true };
    },
  },

  {
    name: 'add_top_bottom_format',
    description:
      'Highlight the top or bottom N items (or percent) in a range. For example, highlight the top 10 values or the bottom 25%.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "B2:B100")' },
      rank: {
        type: 'number',
        required: false,
        description: 'Number of top/bottom items. Default 10.',
      },
      topOrBottom: {
        type: 'string',
        required: false,
        description: 'Which items to highlight. Default "TopItems".',
        enum: ['TopItems', 'BottomItems', 'TopPercent', 'BottomPercent'],
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color for highlighted items.',
      },
      fillColor: {
        type: 'string',
        required: false,
        description: 'Fill color for highlighted items. Default "green".',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
      const fontColor = args.fontColor as string | undefined;
      const fillColor = (args.fillColor as string) ?? 'green';
      if (fontColor) cf.topBottom.format.font.color = fontColor;
      cf.topBottom.format.fill.color = fillColor;
      cf.topBottom.rule = {
        rank: (args.rank as number) ?? 10,
        type: ((args.topOrBottom as string) ?? 'TopItems') as
          | 'TopItems'
          | 'BottomItems'
          | 'TopPercent'
          | 'BottomPercent',
      };
      await context.sync();
      return { address, ruleType: 'topBottom', applied: true };
    },
  },

  {
    name: 'add_contains_text_format',
    description:
      'Highlight cells that contain specific text. For example, highlight all cells containing "Error" in red.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "A1:D100")' },
      text: { type: 'string', description: 'The text to search for in cells' },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color for matching cells. Default "red".',
      },
      fillColor: { type: 'string', required: false, description: 'Fill color for matching cells.' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
      const fontColor = (args.fontColor as string) ?? 'red';
      const fillColor = args.fillColor as string | undefined;
      cf.textComparison.format.font.color = fontColor;
      if (fillColor) cf.textComparison.format.fill.color = fillColor;
      cf.textComparison.rule = {
        operator: Excel.ConditionalTextOperator.contains,
        text: args.text as string,
      };
      await context.sync();
      return { address, ruleType: 'containsText', applied: true };
    },
  },

  {
    name: 'add_custom_format',
    description:
      'Add a custom formula-based conditional format. The formula must evaluate to TRUE for the formatting to apply. For example, "=A1>B1" to highlight cells where column A exceeds column B.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "A1:D20")' },
      formula: {
        type: 'string',
        description: 'Excel formula that evaluates to TRUE/FALSE (e.g., "=A1>B1")',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color when formula is TRUE.',
      },
      fillColor: {
        type: 'string',
        required: false,
        description: 'Fill color when formula is TRUE.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = args.formula as string;
      const fontColor = args.fontColor as string | undefined;
      const fillColor = args.fillColor as string | undefined;
      if (fontColor) cf.custom.format.font.color = fontColor;
      if (fillColor) cf.custom.format.fill.color = fillColor;
      await context.sync();
      return { address, ruleType: 'custom', applied: true };
    },
  },

  // ─── List & Clear (unchanged) ──────────────────────────

  {
    name: 'list_conditional_formats',
    description:
      "List all conditional formatting rules applied to a range. Returns each rule's type, priority, and stopIfTrue setting.",
    params: {
      address: { type: 'string', description: 'The range address to inspect (e.g., "A1:D20")' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      const cfs = range.conditionalFormats;
      cfs.load('items');
      await context.sync();

      for (const cf of cfs.items) {
        cf.load(['type', 'priority', 'stopIfTrue']);
      }
      await context.sync();

      const result = cfs.items.map(cf => ({
        type: cf.type,
        priority: cf.priority,
        stopIfTrue: cf.stopIfTrue,
      }));
      return { conditionalFormats: result, count: result.length };
    },
  },

  {
    name: 'clear_conditional_formats',
    description:
      'Remove all conditional formatting rules from a range, or from the entire sheet if no address is specified.',
    params: {
      address: {
        type: 'string',
        required: false,
        description: 'Optional range address. If omitted, clears the entire sheet.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string | undefined;
      const range = address ? sheet.getRange(address) : sheet.getRange();
      range.conditionalFormats.clearAll();
      await context.sync();
      return { address: address ?? 'entire sheet', cleared: true };
    },
  },
];
