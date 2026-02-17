/**
 * Workbook tool configs — 8 tools for workbook overview, selection, named ranges,
 * and Power Query query metadata.
 */

import type { ToolConfig } from '../codegen';

export const workbookConfigs: readonly ToolConfig[] = [
  {
    name: 'get_workbook_info',
    description:
      'Get a high-level overview of the entire workbook. Returns all sheet names, the active sheet, the used range dimensions of the active sheet, a list of all table names, and counts. This is the best starting point to understand what data exists before performing operations.',
    params: {},
    execute: async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      const usedRange = activeSheet.getUsedRangeOrNullObject();
      usedRange.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      const tables = context.workbook.tables;
      tables.load('items');
      await context.sync();

      for (const s of sheets.items) {
        s.load('name');
      }
      for (const t of tables.items) {
        t.load('name');
      }
      await context.sync();

      return {
        sheetNames: sheets.items.map(s => s.name),
        sheetCount: sheets.items.length,
        activeSheet: activeSheet.name,
        usedRange: usedRange.isNullObject ? null : usedRange.address,
        usedRangeRows: usedRange.isNullObject ? 0 : usedRange.rowCount,
        usedRangeColumns: usedRange.isNullObject ? 0 : usedRange.columnCount,
        tableNames: tables.items.map(t => t.name),
        tableCount: tables.items.length,
      };
    },
  },

  {
    name: 'get_selected_range',
    description:
      'Get the address and values of the range the user currently has selected (highlighted) in Excel. Useful when the user says "this data", "selected cells", or "what I have highlighted".',
    params: {},
    execute: async context => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'rowCount', 'columnCount', 'values', 'numberFormat']);
      await context.sync();
      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
      };
    },
  },

  {
    name: 'define_named_range',
    description:
      'Create a named range — a human-readable alias for a cell range (e.g., "SalesData" → Sheet1!A1:D100). Named ranges make formulas more readable and can be used in formulas across the workbook.',
    params: {
      name: { type: 'string', description: 'Name for the range (e.g., "SalesData")' },
      address: { type: 'string', description: 'The range address (e.g., "A1:D100")' },
      comment: {
        type: 'string',
        required: false,
        description: 'Optional comment describing the named range.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheetName = args.sheetName as string | undefined;
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.address as string);
      const comment = args.comment as string | undefined;
      const namedItem = context.workbook.names.add(args.name as string, range, comment);
      namedItem.load(['name', 'comment']);
      await context.sync();
      return { name: namedItem.name, address: args.address, comment: namedItem.comment ?? '' };
    },
  },

  {
    name: 'list_named_ranges',
    description:
      'List all named ranges defined in the workbook. Returns each name, the range address it refers to, and any comment.',
    params: {},
    execute: async context => {
      const names = context.workbook.names;
      names.load('items');
      await context.sync();

      for (const n of names.items) {
        n.load(['name', 'comment', 'value']);
      }
      await context.sync();

      const result = names.items.map(n => ({
        name: n.name,
        value: n.value as string,
        comment: n.comment ?? '',
      }));
      return { namedRanges: result, count: result.length };
    },
  },

  {
    name: 'recalculate_workbook',
    description: 'Force a full recalculation of all formulas in the entire workbook.',
    params: {
      recalcType: {
        type: 'string',
        required: false,
        description: 'Recalculation type',
        enum: ['Recalculate', 'Full'],
      },
    },
    execute: async (context, args) => {
      const recalcType = (args.recalcType as string) ?? 'Full';
      context.application.calculate(recalcType as Excel.CalculationType);
      await context.sync();
      return { recalculated: true, type: recalcType };
    },
  },

  {
    name: 'save_workbook',
    description: 'Save the current workbook. Optionally prompt the user before saving.',
    params: {
      saveBehavior: {
        type: 'string',
        required: false,
        description: 'Save behavior',
        enum: ['Save', 'Prompt'],
      },
    },
    execute: async (context, args) => {
      const saveBehavior = (args.saveBehavior as string) ?? 'Save';
      context.workbook.save(saveBehavior as Excel.SaveBehavior);
      await context.sync();
      return { saved: true, saveBehavior };
    },
  },

  {
    name: 'get_workbook_properties',
    description: 'Get workbook document properties such as title, author, subject, and category.',
    params: {},
    execute: async context => {
      const properties = context.workbook.properties;
      properties.load([
        'author',
        'category',
        'comments',
        'company',
        'keywords',
        'lastAuthor',
        'manager',
        'revisionNumber',
        'subject',
        'title',
      ]);
      await context.sync();
      return {
        author: properties.author,
        category: properties.category,
        comments: properties.comments,
        company: properties.company,
        keywords: properties.keywords,
        lastAuthor: properties.lastAuthor,
        manager: properties.manager,
        revisionNumber: properties.revisionNumber,
        subject: properties.subject,
        title: properties.title,
      };
    },
  },

  {
    name: 'set_workbook_properties',
    description:
      'Set workbook document properties such as title, author, subject, category, and keywords.',
    params: {
      author: { type: 'string', required: false, description: 'Workbook author' },
      category: { type: 'string', required: false, description: 'Workbook category' },
      comments: { type: 'string', required: false, description: 'Workbook comments metadata' },
      company: { type: 'string', required: false, description: 'Workbook company' },
      keywords: { type: 'string', required: false, description: 'Workbook keywords' },
      manager: { type: 'string', required: false, description: 'Workbook manager' },
      revisionNumber: {
        type: 'number',
        required: false,
        description: 'Workbook revision number',
      },
      subject: { type: 'string', required: false, description: 'Workbook subject' },
      title: { type: 'string', required: false, description: 'Workbook title' },
    },
    execute: async (context, args) => {
      const properties = context.workbook.properties;

      if (args.author !== undefined) properties.author = args.author as string;
      if (args.category !== undefined) properties.category = args.category as string;
      if (args.comments !== undefined) properties.comments = args.comments as string;
      if (args.company !== undefined) properties.company = args.company as string;
      if (args.keywords !== undefined) properties.keywords = args.keywords as string;
      if (args.manager !== undefined) properties.manager = args.manager as string;
      if (args.revisionNumber !== undefined)
        properties.revisionNumber = args.revisionNumber as number;
      if (args.subject !== undefined) properties.subject = args.subject as string;
      if (args.title !== undefined) properties.title = args.title as string;

      properties.load([
        'author',
        'category',
        'comments',
        'company',
        'keywords',
        'manager',
        'revisionNumber',
        'subject',
        'title',
      ]);
      await context.sync();

      return {
        updated: true,
        author: properties.author,
        category: properties.category,
        comments: properties.comments,
        company: properties.company,
        keywords: properties.keywords,
        manager: properties.manager,
        revisionNumber: properties.revisionNumber,
        subject: properties.subject,
        title: properties.title,
      };
    },
  },

  {
    name: 'get_workbook_protection',
    description: 'Get current workbook protection state.',
    params: {},
    execute: async context => {
      const protection = context.workbook.protection;
      protection.load('protected');
      await context.sync();
      return { protected: protection.protected };
    },
  },

  {
    name: 'protect_workbook',
    description: 'Protect the workbook structure. Optionally provide a password.',
    params: {
      password: {
        type: 'string',
        required: false,
        description: 'Optional workbook protection password',
      },
    },
    execute: async (context, args) => {
      context.workbook.protection.protect(args.password as string | undefined);
      const protection = context.workbook.protection;
      protection.load('protected');
      await context.sync();
      return { protected: protection.protected };
    },
  },

  {
    name: 'unprotect_workbook',
    description: 'Remove workbook protection. Provide password if one was set.',
    params: {
      password: {
        type: 'string',
        required: false,
        description: 'Workbook protection password if required',
      },
    },
    execute: async (context, args) => {
      context.workbook.protection.unprotect(args.password as string | undefined);
      const protection = context.workbook.protection;
      protection.load('protected');
      await context.sync();
      return { protected: protection.protected };
    },
  },

  {
    name: 'refresh_data_connections',
    description:
      'Refresh all workbook data connections supported by the Excel data connection API.',
    params: {},
    execute: async context => {
      context.workbook.dataConnections.refreshAll();
      await context.sync();
      return { refreshed: true };
    },
  },

  // ─── Power Query ───────────────────────────────────────

  {
    name: 'list_queries',
    description:
      'List all Power Query queries in the workbook. Returns query name, load target type, whether it loads to data model, last refresh time, rows loaded, and latest error status.',
    params: {},
    execute: async context => {
      const queries = context.workbook.queries;
      queries.load('items');
      await context.sync();

      for (const query of queries.items) {
        query.load([
          'name',
          'loadedTo',
          'loadedToDataModel',
          'refreshDate',
          'rowsLoadedCount',
          'error',
        ]);
      }
      await context.sync();

      const result = queries.items.map(query => ({
        name: query.name,
        loadedTo: query.loadedTo,
        loadedToDataModel: query.loadedToDataModel,
        refreshDate: query.refreshDate,
        rowsLoadedCount: query.rowsLoadedCount,
        error: query.error,
      }));

      return { queries: result, count: result.length };
    },
  },

  {
    name: 'get_query',
    description:
      'Get details for a single Power Query query by name. Returns load target type, data-model loading flag, last refresh time, rows loaded, and latest error status.',
    params: {
      queryName: {
        type: 'string',
        description: 'Name of the Power Query query to retrieve (case-insensitive).',
      },
    },
    execute: async (context, args) => {
      const queryName = args.queryName as string;
      const query = context.workbook.queries.getItem(queryName);
      query.load([
        'name',
        'loadedTo',
        'loadedToDataModel',
        'refreshDate',
        'rowsLoadedCount',
        'error',
      ]);
      await context.sync();

      return {
        name: query.name,
        loadedTo: query.loadedTo,
        loadedToDataModel: query.loadedToDataModel,
        refreshDate: query.refreshDate,
        rowsLoadedCount: query.rowsLoadedCount,
        error: query.error,
      };
    },
  },

  {
    name: 'get_query_count',
    description: 'Get the number of Power Query queries in the workbook.',
    params: {},
    execute: async context => {
      const count = context.workbook.queries.getCount();
      await context.sync();
      return { count: count.value };
    },
  },
];
