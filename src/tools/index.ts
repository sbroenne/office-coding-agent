import { createTools } from './codegen';
import {
  rangeConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
} from './configs';
import type { ToolConfig } from './codegen/types';
import { tool, type ToolSet, type ToolExecutionOptions } from 'ai';
import { z } from 'zod';
import type { OfficeHostApp } from '@/services/office/host';

export { getGeneralTools, webFetchTool, createRunSubagentTool } from './general';

export const MAX_TOOLS_PER_REQUEST = 128;

/** All tool configs combined for manifest generation */
export const allConfigs: readonly (readonly ToolConfig[])[] = [
  rangeConfigs,
  tableConfigs,
  chartConfigs,
  sheetConfigs,
  workbookConfigs,
  commentConfigs,
  conditionalFormatConfigs,
  dataValidationConfigs,
  pivotTableConfigs,
];

/** All Excel tools combined into a single record for AI SDK */
export const excelTools: ToolSet = allConfigs.reduce<ToolSet>((acc, configs) => {
  const generatedTools = createTools(configs);
  return { ...acc, ...generatedTools };
}, {});

export const powerPointTools: ToolSet = {};

const CONSOLIDATED_COMMENT_TOOL_NAMES = [
  'add_comment',
  'list_comments',
  'edit_comment',
  'delete_comment',
] as const;
const CONSOLIDATED_WORKBOOK_PROTECTION_TOOL_NAMES = [
  'get_workbook_protection',
  'protect_workbook',
  'unprotect_workbook',
] as const;
const CONSOLIDATED_QUERY_TOOL_NAMES = ['list_queries', 'get_query', 'get_query_count'] as const;
const CONSOLIDATED_NAMED_RANGE_TOOL_NAMES = ['define_named_range', 'list_named_ranges'] as const;
const CONSOLIDATED_WORKBOOK_ADMIN_TOOL_NAMES = [
  'recalculate_workbook',
  'save_workbook',
  'get_workbook_properties',
  'set_workbook_properties',
] as const;
const CONSOLIDATED_TABLE_ADVANCED_TOOL_NAMES = [
  'add_table_column',
  'delete_table_column',
  'convert_table_to_range',
  'resize_table',
  'set_table_style',
  'set_table_header_totals_visibility',
  'reapply_table_filters',
] as const;
const CONSOLIDATED_RANGE_ADVANCED_TOOL_NAMES = [
  'auto_fill_range',
  'flash_fill_range',
  'get_special_cells',
  'get_range_precedents',
  'get_range_dependents',
  'recalculate_range',
  'get_tables_for_range',
  'toggle_row_column_visibility',
  'group_rows_columns',
  'ungroup_rows_columns',
  'set_cell_borders',
] as const;

function getRequiredTool(name: string) {
  const toolDef = excelTools[name];
  if (!toolDef) {
    throw new Error(`Missing required consolidated source tool: ${name}`);
  }
  return toolDef;
}

function buildConsolidatedExcelTools(): ToolSet {
  const merged = { ...excelTools } as ToolSet;

  const addComment = getRequiredTool('add_comment');
  const listComments = getRequiredTool('list_comments');
  const editComment = getRequiredTool('edit_comment');
  const deleteComment = getRequiredTool('delete_comment');

  const getWorkbookProtection = getRequiredTool('get_workbook_protection');
  const protectWorkbook = getRequiredTool('protect_workbook');
  const unprotectWorkbook = getRequiredTool('unprotect_workbook');

  const listQueries = getRequiredTool('list_queries');
  const getQuery = getRequiredTool('get_query');
  const getQueryCount = getRequiredTool('get_query_count');

  const defineNamedRange = getRequiredTool('define_named_range');
  const listNamedRanges = getRequiredTool('list_named_ranges');

  const recalculateWorkbook = getRequiredTool('recalculate_workbook');
  const saveWorkbook = getRequiredTool('save_workbook');
  const getWorkbookProperties = getRequiredTool('get_workbook_properties');
  const setWorkbookProperties = getRequiredTool('set_workbook_properties');

  const addTableColumn = getRequiredTool('add_table_column');
  const deleteTableColumn = getRequiredTool('delete_table_column');
  const convertTableToRange = getRequiredTool('convert_table_to_range');
  const resizeTable = getRequiredTool('resize_table');
  const setTableStyle = getRequiredTool('set_table_style');
  const setTableHeaderTotalsVisibility = getRequiredTool('set_table_header_totals_visibility');
  const reapplyTableFilters = getRequiredTool('reapply_table_filters');

  const autoFillRange = getRequiredTool('auto_fill_range');
  const flashFillRange = getRequiredTool('flash_fill_range');
  const getSpecialCells = getRequiredTool('get_special_cells');
  const getRangePrecedents = getRequiredTool('get_range_precedents');
  const getRangeDependents = getRequiredTool('get_range_dependents');
  const recalculateRange = getRequiredTool('recalculate_range');
  const getTablesForRange = getRequiredTool('get_tables_for_range');
  const toggleRowColumnVisibility = getRequiredTool('toggle_row_column_visibility');
  const groupRowsColumns = getRequiredTool('group_rows_columns');
  const ungroupRowsColumns = getRequiredTool('ungroup_rows_columns');
  const setCellBorders = getRequiredTool('set_cell_borders');

  for (const name of [
    ...CONSOLIDATED_COMMENT_TOOL_NAMES,
    ...CONSOLIDATED_WORKBOOK_PROTECTION_TOOL_NAMES,
    ...CONSOLIDATED_QUERY_TOOL_NAMES,
    ...CONSOLIDATED_NAMED_RANGE_TOOL_NAMES,
    ...CONSOLIDATED_WORKBOOK_ADMIN_TOOL_NAMES,
    ...CONSOLIDATED_TABLE_ADVANCED_TOOL_NAMES,
    ...CONSOLIDATED_RANGE_ADVANCED_TOOL_NAMES,
  ]) {
    delete merged[name as keyof ToolSet];
  }

  merged.manage_comments = tool({
    description:
      'Manage cell comments with a single tool. Actions: add, list, edit, delete. Use this instead of separate comment tools.',
    inputSchema: z.object({
      action: z.enum(['add', 'list', 'edit', 'delete']).describe('Comment action to perform'),
      cellAddress: z
        .string()
        .optional()
        .describe('Cell address for add/edit/delete actions (e.g., A1)'),
      text: z.string().optional().describe('Comment text for add action'),
      newText: z.string().optional().describe('Replacement comment text for edit action'),
      sheetName: z.string().optional().describe('Optional worksheet name'),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      if (args.action === 'list') {
        return listComments.execute!({ sheetName: args.sheetName }, options) as Promise<unknown>;
      }

      if (!args.cellAddress) {
        throw new Error('cellAddress is required for add/edit/delete comment actions.');
      }

      if (args.action === 'add') {
        if (!args.text) throw new Error('text is required for add comment action.');
        return addComment.execute!(
          {
            cellAddress: args.cellAddress,
            text: args.text,
            sheetName: args.sheetName,
          },
          options
        ) as Promise<unknown>;
      }

      if (args.action === 'edit') {
        if (!args.newText) throw new Error('newText is required for edit comment action.');
        return editComment.execute!(
          {
            cellAddress: args.cellAddress,
            newText: args.newText,
            sheetName: args.sheetName,
          },
          options
        ) as Promise<unknown>;
      }

      return deleteComment.execute!(
        {
          cellAddress: args.cellAddress,
          sheetName: args.sheetName,
        },
        options
      ) as Promise<unknown>;
    },
  });

  merged.manage_workbook_protection = tool({
    description:
      'Manage workbook protection state with one tool. Actions: get, protect, unprotect.',
    inputSchema: z.object({
      action: z
        .enum(['get', 'protect', 'unprotect'])
        .describe('Workbook protection action to perform'),
      password: z.string().optional().describe('Optional protection password'),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      if (args.action === 'get') {
        return getWorkbookProtection.execute!({}, options) as Promise<unknown>;
      }

      if (args.action === 'protect') {
        return protectWorkbook.execute!({ password: args.password }, options) as Promise<unknown>;
      }

      return unprotectWorkbook.execute!({ password: args.password }, options) as Promise<unknown>;
    },
  });

  merged.manage_power_queries = tool({
    description: 'Manage Power Query metadata with one tool. Actions: list, get, count.',
    inputSchema: z.object({
      action: z.enum(['list', 'get', 'count']).describe('Power Query metadata action to perform'),
      queryName: z.string().optional().describe('Required for get action: query name to retrieve'),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      if (args.action === 'list') {
        return listQueries.execute!({}, options) as Promise<unknown>;
      }

      if (args.action === 'count') {
        return getQueryCount.execute!({}, options) as Promise<unknown>;
      }

      if (!args.queryName) {
        throw new Error('queryName is required for get action.');
      }

      return getQuery.execute!({ queryName: args.queryName }, options) as Promise<unknown>;
    },
  });

  merged.manage_named_ranges = tool({
    description: 'Manage named ranges with one tool. Actions: define, list.',
    inputSchema: z.object({
      action: z.enum(['define', 'list']).describe('Named range action to perform'),
      name: z.string().optional().describe('Name for define action'),
      address: z.string().optional().describe('Address for define action (e.g., A1:D10)'),
      comment: z.string().optional().describe('Optional comment for define action'),
      sheetName: z.string().optional().describe('Optional worksheet name for define action'),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      if (args.action === 'list') {
        return listNamedRanges.execute!({}, options) as Promise<unknown>;
      }

      if (!args.name || !args.address) {
        throw new Error('name and address are required for define action.');
      }

      return defineNamedRange.execute!(
        {
          name: args.name,
          address: args.address,
          comment: args.comment,
          sheetName: args.sheetName,
        },
        options
      ) as Promise<unknown>;
    },
  });

  merged.manage_workbook_admin = tool({
    description:
      'Manage workbook administrative operations with one tool. Actions: recalculate, save, get-properties, set-properties.',
    inputSchema: z.object({
      action: z
        .enum(['recalculate', 'save', 'get-properties', 'set-properties'])
        .describe('Workbook admin action to perform'),
      recalcType: z.enum(['Recalculate', 'Full']).optional().describe('For recalculate action'),
      saveBehavior: z.enum(['Save', 'Prompt']).optional().describe('For save action'),
      author: z.string().optional(),
      category: z.string().optional(),
      comments: z.string().optional(),
      company: z.string().optional(),
      keywords: z.string().optional(),
      manager: z.string().optional(),
      revisionNumber: z.number().optional(),
      subject: z.string().optional(),
      title: z.string().optional(),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      if (args.action === 'recalculate') {
        return recalculateWorkbook.execute!(
          { recalcType: args.recalcType },
          options
        ) as Promise<unknown>;
      }

      if (args.action === 'save') {
        return saveWorkbook.execute!(
          { saveBehavior: args.saveBehavior },
          options
        ) as Promise<unknown>;
      }

      if (args.action === 'get-properties') {
        return getWorkbookProperties.execute!({}, options) as Promise<unknown>;
      }

      return setWorkbookProperties.execute!(
        {
          author: args.author,
          category: args.category,
          comments: args.comments,
          company: args.company,
          keywords: args.keywords,
          manager: args.manager,
          revisionNumber: args.revisionNumber,
          subject: args.subject,
          title: args.title,
        },
        options
      ) as Promise<unknown>;
    },
  });

  merged.manage_table_advanced = tool({
    description:
      'Manage advanced table operations with one tool. Actions: add-column, delete-column, convert-to-range, resize, set-style, set-header-totals-visibility, reapply-filters.',
    inputSchema: z.object({
      action: z
        .enum([
          'add-column',
          'delete-column',
          'convert-to-range',
          'resize',
          'set-style',
          'set-header-totals-visibility',
          'reapply-filters',
        ])
        .describe('Advanced table action to perform'),
      tableName: z.string().describe('Target table name'),
      columnName: z.string().optional(),
      columnData: z.array(z.string()).optional(),
      newAddress: z.string().optional(),
      style: z.string().optional(),
      showHeaders: z.boolean().optional(),
      showTotals: z.boolean().optional(),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      switch (args.action) {
        case 'add-column':
          return addTableColumn.execute!(
            {
              tableName: args.tableName,
              columnName: args.columnName,
              columnData: args.columnData,
            },
            options
          ) as Promise<unknown>;
        case 'delete-column':
          if (!args.columnName) throw new Error('columnName is required for delete-column action.');
          return deleteTableColumn.execute!(
            { tableName: args.tableName, columnName: args.columnName },
            options
          ) as Promise<unknown>;
        case 'convert-to-range':
          return convertTableToRange.execute!(
            { tableName: args.tableName },
            options
          ) as Promise<unknown>;
        case 'resize':
          if (!args.newAddress) throw new Error('newAddress is required for resize action.');
          return resizeTable.execute!(
            { tableName: args.tableName, newAddress: args.newAddress },
            options
          ) as Promise<unknown>;
        case 'set-style':
          if (!args.style) throw new Error('style is required for set-style action.');
          return setTableStyle.execute!(
            { tableName: args.tableName, style: args.style },
            options
          ) as Promise<unknown>;
        case 'set-header-totals-visibility':
          return setTableHeaderTotalsVisibility.execute!(
            {
              tableName: args.tableName,
              showHeaders: args.showHeaders,
              showTotals: args.showTotals,
            },
            options
          ) as Promise<unknown>;
        default:
          return reapplyTableFilters.execute!(
            { tableName: args.tableName },
            options
          ) as Promise<unknown>;
      }
    },
  });

  merged.manage_range_advanced = tool({
    description:
      'Manage advanced range operations with one tool. Actions: auto-fill, flash-fill, get-special-cells, get-precedents, get-dependents, recalculate, get-tables, visibility, group, ungroup, set-borders.',
    inputSchema: z.object({
      action: z
        .enum([
          'auto-fill',
          'flash-fill',
          'get-special-cells',
          'get-precedents',
          'get-dependents',
          'recalculate',
          'get-tables',
          'visibility',
          'group',
          'ungroup',
          'set-borders',
        ])
        .describe('Advanced range action to perform'),
      address: z.string().optional(),
      sourceAddress: z.string().optional(),
      destinationAddress: z.string().optional(),
      autoFillType: z.string().optional(),
      cellType: z.string().optional(),
      valueType: z.string().optional(),
      target: z.enum(['rows', 'columns']).optional(),
      visible: z.boolean().optional(),
      groupBy: z.enum(['Rows', 'Columns']).optional(),
      borderStyle: z.string().optional(),
      borderColor: z.string().optional(),
      sheetName: z.string().optional(),
    }),
    execute: (args, options: ToolExecutionOptions) => {
      switch (args.action) {
        case 'auto-fill':
          if (!args.sourceAddress || !args.destinationAddress) {
            throw new Error(
              'sourceAddress and destinationAddress are required for auto-fill action.'
            );
          }
          return autoFillRange.execute!(
            {
              sourceAddress: args.sourceAddress,
              destinationAddress: args.destinationAddress,
              autoFillType: args.autoFillType,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
        case 'flash-fill':
          if (!args.address) throw new Error('address is required for flash-fill action.');
          return flashFillRange.execute!(
            { address: args.address, sheetName: args.sheetName },
            options
          ) as Promise<unknown>;
        case 'get-special-cells':
          if (!args.address || !args.cellType) {
            throw new Error('address and cellType are required for get-special-cells action.');
          }
          return getSpecialCells.execute!(
            {
              address: args.address,
              cellType: args.cellType,
              valueType: args.valueType,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
        case 'get-precedents':
          if (!args.address) throw new Error('address is required for get-precedents action.');
          return getRangePrecedents.execute!(
            { address: args.address, sheetName: args.sheetName },
            options
          ) as Promise<unknown>;
        case 'get-dependents':
          if (!args.address) throw new Error('address is required for get-dependents action.');
          return getRangeDependents.execute!(
            { address: args.address, sheetName: args.sheetName },
            options
          ) as Promise<unknown>;
        case 'recalculate':
          if (!args.address) throw new Error('address is required for recalculate action.');
          return recalculateRange.execute!(
            { address: args.address, sheetName: args.sheetName },
            options
          ) as Promise<unknown>;
        case 'get-tables':
          if (!args.address) throw new Error('address is required for get-tables action.');
          return getTablesForRange.execute!(
            { address: args.address, sheetName: args.sheetName },
            options
          ) as Promise<unknown>;
        case 'visibility':
          if (!args.address || !args.target || args.visible === undefined) {
            throw new Error('address, target, and visible are required for visibility action.');
          }
          return toggleRowColumnVisibility.execute!(
            {
              address: args.address,
              target: args.target,
              visible: args.visible,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
        case 'group':
          if (!args.address || !args.groupBy) {
            throw new Error('address and groupBy are required for group action.');
          }
          return groupRowsColumns.execute!(
            {
              address: args.address,
              groupBy: args.groupBy,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
        case 'ungroup':
          if (!args.address || !args.groupBy) {
            throw new Error('address and groupBy are required for ungroup action.');
          }
          return ungroupRowsColumns.execute!(
            {
              address: args.address,
              groupBy: args.groupBy,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
        default:
          if (!args.address || !args.borderStyle) {
            throw new Error('address and borderStyle are required for set-borders action.');
          }
          return setCellBorders.execute!(
            {
              address: args.address,
              borderStyle: args.borderStyle,
              borderColor: args.borderColor,
              sheetName: args.sheetName,
            },
            options
          ) as Promise<unknown>;
      }
    },
  });

  return merged;
}

const excelConsolidatedTools = buildConsolidatedExcelTools();

function clampToolSet(toolSet: ToolSet, maxTools = MAX_TOOLS_PER_REQUEST): ToolSet {
  const entries = Object.entries(toolSet);
  if (entries.length <= maxTools) return toolSet;

  return Object.fromEntries(entries.slice(0, maxTools)) as ToolSet;
}

export function getToolsForHost(host: OfficeHostApp): ToolSet {
  switch (host) {
    case 'excel':
      return clampToolSet(excelConsolidatedTools);
    case 'powerpoint':
      return clampToolSet(powerPointTools);
    default:
      return {};
  }
}
