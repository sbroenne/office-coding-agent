/**
 * Data validation tool configs — 1 tool (data_validation) with actions: get, set, clear.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const dataValidationConfigs: readonly ToolConfig[] = [
  {
    name: 'data_validation',
    description:
      'Manage data validation rules on cells. Use action "get" to read the current rule, "set" to apply a new rule (specify type), or "clear" to remove all validation.',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['get', 'set', 'clear'],
      },
      address: { type: 'string', description: 'Range address (e.g. "A1:A100").' },
      type: {
        type: 'string',
        required: false,
        description: 'Validation type for action=set.',
        enum: ['list', 'number', 'date', 'textLength', 'custom'],
      },
      // list
      listValues: {
        type: 'string[]',
        required: false,
        description: 'Allowed values for type=list (e.g. ["Yes","No","Maybe"]).',
      },
      // number / date / textLength
      operator: {
        type: 'string',
        required: false,
        description: 'Comparison operator (number/date/textLength).',
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
      formula1: { type: 'string', required: false, description: 'Min value or formula.' },
      formula2: { type: 'string', required: false, description: 'Max value for Between.' },
      // custom
      customFormula: {
        type: 'string',
        required: false,
        description: 'Custom validation formula (type=custom).',
      },
      // error alert
      showError: {
        type: 'boolean',
        required: false,
        description: 'Show error alert on invalid input. Default true.',
      },
      errorTitle: { type: 'string', required: false, description: 'Error alert title.' },
      errorMessage: { type: 'string', required: false, description: 'Error alert message.' },
      errorStyle: {
        type: 'string',
        required: false,
        description: 'Error alert style.',
        enum: ['Stop', 'Warning', 'Information'],
      },
      // prompt
      showPrompt: { type: 'boolean', required: false, description: 'Show input prompt message.' },
      promptTitle: { type: 'string', required: false, description: 'Input prompt title.' },
      promptMessage: { type: 'string', required: false, description: 'Input prompt message.' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const action = args.action as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const range = sheet.getRange(args.address as string);

      if (action === 'get') {
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
      }

      if (action === 'clear') {
        range.dataValidation.clear();
        await context.sync();
        return { address: args.address, cleared: true };
      }

      // action === 'set'
      const type = args.type as string;
      const dv = range.dataValidation;

      if (type === 'list') {
        const values = args.listValues as string[];
        dv.rule = { list: { inCellDropDown: true, source: values.join(',') } };
      } else if (type === 'number') {
        dv.rule = {
          wholeNumber: {
            operator: args.operator as Excel.DataValidationOperator,
            formula1: String(args.formula1),
            ...(args.formula2 ? { formula2: args.formula2 as string } : {}),
          },
        };
      } else if (type === 'decimal') {
        dv.rule = {
          decimal: {
            operator: args.operator as Excel.DataValidationOperator,
            formula1: String(args.formula1),
            ...(args.formula2 ? { formula2: args.formula2 as string } : {}),
          },
        };
      } else if (type === 'date') {
        dv.rule = {
          date: {
            operator: args.operator as Excel.DataValidationOperator,
            formula1: String(args.formula1),
            ...(args.formula2 ? { formula2: args.formula2 as string } : {}),
          },
        };
      } else if (type === 'textLength') {
        dv.rule = {
          textLength: {
            operator: args.operator as Excel.DataValidationOperator,
            formula1: String(args.formula1),
            ...(args.formula2 ? { formula2: args.formula2 as string } : {}),
          },
        };
      } else {
        // custom
        dv.rule = { custom: { formula: args.customFormula as string } };
      }

      if (args.showError !== false) {
        dv.errorAlert = {
          showAlert: (args.showError as boolean) ?? true,
          title: (args.errorTitle as string) ?? '',
          message: (args.errorMessage as string) ?? '',
          style: (args.errorStyle as Excel.DataValidationAlertStyle) ?? 'Stop',
        };
      }
      if (args.showPrompt) {
        dv.prompt = {
          showPrompt: true,
          title: (args.promptTitle as string) ?? '',
          message: (args.promptMessage as string) ?? '',
        };
      }

      await context.sync();
      return { address: args.address, type, set: true };
    },
  },
];
