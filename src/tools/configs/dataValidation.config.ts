/**
 * Data validation tool configs — 7 tools (decomposed from 3).
 *
 * The old mega-tool `set_data_validation` with 6 ruleType branches is now
 * 5 single-purpose tools. Each tool has exactly the parameters it needs.
 * All 5 share optional errorAlert and prompt params.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

/** Shared params for error alert and prompt (used by all set_*_validation tools) */
const validationAlertParams = {
  errorMessage: {
    type: 'string' as const,
    required: false,
    description: 'Error message shown when invalid data is entered.',
  },
  errorTitle: {
    type: 'string' as const,
    required: false,
    description: 'Title of the error alert dialog. Default "Invalid Input".',
  },
  promptMessage: {
    type: 'string' as const,
    required: false,
    description: 'Help message shown when the cell is selected.',
  },
  promptTitle: {
    type: 'string' as const,
    required: false,
    description: 'Title of the input prompt. Default "Input Help".',
  },
};

/** Apply optional error alert and prompt to a range's dataValidation */
function applyAlertAndPrompt(range: Excel.Range, args: Record<string, unknown>): void {
  const errorMessage = args.errorMessage as string | undefined;
  if (errorMessage) {
    range.dataValidation.errorAlert = {
      message: errorMessage,
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: (args.errorTitle as string) ?? 'Invalid Input',
    };
  }
  const promptMessage = args.promptMessage as string | undefined;
  if (promptMessage) {
    range.dataValidation.prompt = {
      message: promptMessage,
      showPrompt: true,
      title: (args.promptTitle as string) ?? 'Input Help',
    };
  }
}

export const dataValidationConfigs: readonly ToolConfig[] = [
  // ─── 5 decomposed set tools ────────────────────────────

  {
    name: 'set_list_validation',
    description:
      'Set a dropdown list validation on a range. Users will see a dropdown with the specified values. Use comma-separated values or a range reference as the source.',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "B2:B100")' },
      source: {
        type: 'string',
        description:
          'Comma-separated values or a range reference (e.g., "Yes,No,Maybe" or "=Sheet2!A1:A5")',
      },
      inCellDropDown: {
        type: 'boolean',
        required: false,
        description: 'Show a dropdown in the cell. Default true.',
      },
      ...validationAlertParams,
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      range.dataValidation.rule = {
        list: {
          inCellDropDown: (args.inCellDropDown as boolean) ?? true,
          source: args.source as string,
        },
      };
      applyAlertAndPrompt(range, args);
      await context.sync();
      return { address, ruleType: 'list', applied: true };
    },
  },

  {
    name: 'set_number_validation',
    description:
      'Restrict cell input to whole numbers or decimals matching a condition (e.g., greater than 0, between 1 and 100).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "C2:C50")' },
      numberType: {
        type: 'string',
        description: 'Whether to allow whole numbers or decimals.',
        enum: ['wholeNumber', 'decimal'],
      },
      operator: {
        type: 'string',
        description: 'Comparison operator',
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
      formula1: { type: 'string', description: 'First value or formula (e.g., "0", "=A1")' },
      formula2: {
        type: 'string',
        required: false,
        description: 'Second value for Between/NotBetween operators.',
      },
      ...validationAlertParams,
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const numberType = args.numberType as 'wholeNumber' | 'decimal';
      const rule: Excel.BasicDataValidation = {
        formula1: args.formula1 as string | number,
        operator: args.operator as Excel.DataValidationOperator,
      };
      if (args.formula2 !== undefined) rule.formula2 = args.formula2 as string | number;
      range.dataValidation.rule = { [numberType]: rule };
      applyAlertAndPrompt(range, args);
      await context.sync();
      return { address, ruleType: numberType, applied: true };
    },
  },

  {
    name: 'set_date_validation',
    description:
      'Restrict cell input to dates matching a condition (e.g., after 2024-01-01, between two dates).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "D2:D50")' },
      operator: {
        type: 'string',
        description: 'Comparison operator',
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
      formula1: { type: 'string', description: 'First date or formula (e.g., "2024-01-01")' },
      formula2: {
        type: 'string',
        required: false,
        description: 'Second date for Between/NotBetween operators.',
      },
      ...validationAlertParams,
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const rule: Excel.DateTimeDataValidation = {
        formula1: args.formula1 as string,
        operator: args.operator as Excel.DataValidationOperator,
      };
      if (args.formula2) rule.formula2 = args.formula2 as string;
      range.dataValidation.rule = { date: rule };
      applyAlertAndPrompt(range, args);
      await context.sync();
      return { address, ruleType: 'date', applied: true };
    },
  },

  {
    name: 'set_text_length_validation',
    description:
      'Restrict cell input based on text length (e.g., maximum 100 characters, between 5 and 50 characters).',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "E2:E100")' },
      operator: {
        type: 'string',
        description: 'Comparison operator',
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
      formula1: { type: 'string', description: 'First value (e.g., "0", "5")' },
      formula2: {
        type: 'string',
        required: false,
        description: 'Second value for Between/NotBetween operators.',
      },
      ...validationAlertParams,
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      const rule: Excel.BasicDataValidation = {
        formula1: args.formula1 as string | number,
        operator: args.operator as Excel.DataValidationOperator,
      };
      if (args.formula2 !== undefined) rule.formula2 = args.formula2 as string | number;
      range.dataValidation.rule = { textLength: rule };
      applyAlertAndPrompt(range, args);
      await context.sync();
      return { address, ruleType: 'textLength', applied: true };
    },
  },

  {
    name: 'set_custom_validation',
    description:
      'Set a custom formula-based validation rule. The formula must evaluate to TRUE for the input to be accepted (e.g., "=LEN(A1)<=100").',
    params: {
      address: { type: 'string', description: 'The range address (e.g., "F2:F50")' },
      formula: {
        type: 'string',
        description: 'Excel formula that must evaluate to TRUE (e.g., "=LEN(A1)<=100")',
      },
      ...validationAlertParams,
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const address = args.address as string;
      const range = sheet.getRange(address);
      range.dataValidation.rule = { custom: { formula: args.formula as string } };
      applyAlertAndPrompt(range, args);
      await context.sync();
      return { address, ruleType: 'custom', applied: true };
    },
  },

  // ─── Get & Clear (unchanged) ───────────────────────────

  {
    name: 'get_data_validation',
    description:
      'Get the data validation rule currently applied to a range. Returns the rule type, rule details, error alert, and prompt configuration.',
    params: {
      address: { type: 'string', description: 'The range address to inspect (e.g., "B2")' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.dataValidation.load(['type', 'rule', 'errorAlert', 'prompt', 'ignoreBlanks']);
      await context.sync();
      return {
        address: args.address,
        type: range.dataValidation.type,
        rule: range.dataValidation.rule,
        errorAlert: range.dataValidation.errorAlert,
        prompt: range.dataValidation.prompt,
        ignoreBlanks: range.dataValidation.ignoreBlanks,
      };
    },
  },

  {
    name: 'clear_data_validation',
    description: 'Remove all data validation rules from a range, allowing any input.',
    params: {
      address: {
        type: 'string',
        description: 'The range address to clear validation from (e.g., "B2:B100")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);
      range.dataValidation.clear();
      await context.sync();
      return { address: args.address, cleared: true };
    },
  },
];
