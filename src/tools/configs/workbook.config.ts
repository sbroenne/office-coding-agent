/**
 * Workbook tool configs — 1 tool (workbook) with actions:
 * get_info, get_selected_range, get_properties, set_properties,
 * protect, unprotect, save, recalculate, refresh_connections,
 * define_named_range, list_named_ranges, list_queries, get_query.
 */

import type { ToolConfig } from '../codegen';

export const workbookConfigs: readonly ToolConfig[] = [
  {
    name: 'workbook',
    description:
      'Manage the workbook. Actions: "get_info" (overview), "get_selected_range" (user selection), "get_properties" / "set_properties" (doc metadata), "protect" / "unprotect", "save", "recalculate", "refresh_connections", "define_named_range", "list_named_ranges", "list_queries", "get_query".',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'get_info',
          'get_selected_range',
          'get_properties',
          'set_properties',
          'get_protection',
          'protect',
          'unprotect',
          'save',
          'recalculate',
          'refresh_connections',
          'define_named_range',
          'list_named_ranges',
          'list_queries',
          'get_query',
        ],
      },
      // set_properties
      author: { type: 'string', required: false, description: 'Workbook author (set_properties).' },
      category: { type: 'string', required: false, description: 'Workbook category.' },
      comments: { type: 'string', required: false, description: 'Workbook comments metadata.' },
      company: { type: 'string', required: false, description: 'Workbook company.' },
      keywords: { type: 'string', required: false, description: 'Workbook keywords.' },
      manager: { type: 'string', required: false, description: 'Workbook manager.' },
      subject: { type: 'string', required: false, description: 'Workbook subject.' },
      title: { type: 'string', required: false, description: 'Workbook title.' },
      // protect / unprotect / save
      password: { type: 'string', required: false, description: 'Protection password.' },
      saveBehavior: {
        type: 'string',
        required: false,
        description: 'Save behavior (save).',
        enum: ['Save', 'Prompt'],
      },
      // recalculate
      recalcType: {
        type: 'string',
        required: false,
        description: 'Recalc type.',
        enum: ['Recalculate', 'Full'],
      },
      // define_named_range
      name: {
        type: 'string',
        required: false,
        description: 'Named range name (define_named_range / get_query).',
      },
      address: {
        type: 'string',
        required: false,
        description: 'Range address for define_named_range.',
      },
      comment: { type: 'string', required: false, description: 'Comment for define_named_range.' },
      sheetName: { type: 'string', required: false, description: 'Sheet for define_named_range.' },
      // get_query
      queryName: {
        type: 'string',
        required: false,
        description: 'Power Query name for get_query.',
      },
    },
    execute: async (context, args) => {
      const action = args.action as string;

      if (action === 'get_info') {
        const sheets = context.workbook.worksheets;
        sheets.load('items');
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load('name');
        const usedRange = activeSheet.getUsedRangeOrNullObject();
        usedRange.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
        const tables = context.workbook.tables;
        tables.load('items');
        await context.sync();
        for (const s of sheets.items) s.load('name');
        for (const t of tables.items) t.load('name');
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
      }

      if (action === 'get_selected_range') {
        const range = context.workbook.getSelectedRange();
        range.load(['address', 'rowCount', 'columnCount', 'values']);
        await context.sync();
        return {
          address: range.address,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          values: range.values,
        };
      }

      if (action === 'get_properties') {
        const props = context.workbook.properties;
        props.load([
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
          author: props.author,
          category: props.category,
          comments: props.comments,
          company: props.company,
          keywords: props.keywords,
          lastAuthor: props.lastAuthor,
          manager: props.manager,
          revisionNumber: props.revisionNumber,
          subject: props.subject,
          title: props.title,
        };
      }

      if (action === 'set_properties') {
        const props = context.workbook.properties;
        if (args.author !== undefined) props.author = args.author as string;
        if (args.category !== undefined) props.category = args.category as string;
        if (args.comments !== undefined) props.comments = args.comments as string;
        if (args.company !== undefined) props.company = args.company as string;
        if (args.keywords !== undefined) props.keywords = args.keywords as string;
        if (args.manager !== undefined) props.manager = args.manager as string;
        if (args.subject !== undefined) props.subject = args.subject as string;
        if (args.title !== undefined) props.title = args.title as string;
        await context.sync();
        return { updated: true };
      }

      if (action === 'get_protection') {
        const prot = context.workbook.protection;
        prot.load('protected');
        await context.sync();
        return { protected: prot.protected };
      }

      if (action === 'protect') {
        context.workbook.protection.protect(args.password as string | undefined);
        const prot = context.workbook.protection;
        prot.load('protected');
        await context.sync();
        return { protected: prot.protected };
      }

      if (action === 'unprotect') {
        context.workbook.protection.unprotect(args.password as string | undefined);
        const prot = context.workbook.protection;
        prot.load('protected');
        await context.sync();
        return { protected: prot.protected };
      }

      if (action === 'save') {
        const saveBehavior = (args.saveBehavior as string) ?? 'Save';
        context.workbook.save(saveBehavior as Excel.SaveBehavior);
        await context.sync();
        return { saved: true, saveBehavior };
      }

      if (action === 'recalculate') {
        const recalcType = (args.recalcType as string) ?? 'Full';
        context.application.calculate(recalcType as Excel.CalculationType);
        await context.sync();
        return { recalculated: true, type: recalcType };
      }

      if (action === 'refresh_connections') {
        context.workbook.dataConnections.refreshAll();
        await context.sync();
        return { refreshed: true };
      }

      if (action === 'define_named_range') {
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
      }

      if (action === 'list_named_ranges') {
        const names = context.workbook.names;
        names.load('items');
        await context.sync();
        for (const n of names.items) n.load(['name', 'comment', 'value']);
        await context.sync();
        const result = names.items.map(n => ({
          name: n.name,
          value: n.value as string,
          comment: n.comment ?? '',
        }));
        return { namedRanges: result, count: result.length };
      }

      if (action === 'list_queries') {
        const queries = context.workbook.queries;
        queries.load('items');
        await context.sync();
        for (const q of queries.items)
          q.load([
            'name',
            'loadedTo',
            'loadedToDataModel',
            'refreshDate',
            'rowsLoadedCount',
            'error',
          ]);
        await context.sync();
        const result = queries.items.map(q => ({
          name: q.name,
          loadedTo: q.loadedTo,
          loadedToDataModel: q.loadedToDataModel,
          refreshDate: q.refreshDate,
          rowsLoadedCount: q.rowsLoadedCount,
          error: q.error,
        }));
        return { queries: result, count: result.length };
      }

      // get_query
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
];
