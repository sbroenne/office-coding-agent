/**
 * Sheet tool configs — 1 tool (sheet) with actions:
 * list, create, delete, rename, copy, move, activate,
 * protect, unprotect, freeze, set_visibility, set_gridlines,
 * set_headings, set_page_layout, recalculate.
 */

import type { ToolConfig } from '../codegen';

export const sheetConfigs: readonly ToolConfig[] = [
  {
    name: 'sheet',
    description:
      'Manage worksheets. Actions: "list", "create", "delete", "rename", "copy", "move", "activate", "protect", "unprotect", "freeze" (panes), "set_visibility", "set_gridlines", "set_headings", "set_page_layout", "recalculate".',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: [
          'list',
          'create',
          'delete',
          'rename',
          'copy',
          'move',
          'activate',
          'protect',
          'unprotect',
          'freeze',
          'set_visibility',
          'set_gridlines',
          'set_headings',
          'set_page_layout',
          'recalculate',
        ],
      },
      name: {
        type: 'string',
        required: false,
        description: 'Worksheet name (most actions). For create, this is the new sheet name.',
      },
      newName: { type: 'string', required: false, description: 'New name for rename/copy.' },
      currentName: { type: 'string', required: false, description: 'Current name for rename.' },
      position: { type: 'number', required: false, description: 'New 0-based position (move).' },
      password: {
        type: 'string',
        required: false,
        description: 'Protection password (protect/unprotect).',
      },
      freezeAt: {
        type: 'string',
        required: false,
        description: 'Cell address to freeze at (freeze). Omit to unfreeze.',
      },
      visibility: {
        type: 'string',
        required: false,
        description: 'Visibility state (set_visibility).',
        enum: ['Visible', 'Hidden', 'VeryHidden'],
      },
      tabColor: { type: 'string', required: false, description: 'Tab color hex (set_visibility).' },
      showGridlines: {
        type: 'boolean',
        required: false,
        description: 'True to show gridlines (set_gridlines).',
      },
      showHeadings: {
        type: 'boolean',
        required: false,
        description: 'True to show row/col headings (set_headings).',
      },
      // set_page_layout params
      orientation: {
        type: 'string',
        required: false,
        enum: ['Portrait', 'Landscape'],
        description: 'Page orientation (set_page_layout).',
      },
      paperSize: {
        type: 'string',
        required: false,
        enum: [
          'Letter',
          'LetterSmall',
          'Tabloid',
          'Ledger',
          'Legal',
          'Statement',
          'Executive',
          'A3',
          'A4',
          'A4Small',
          'A5',
        ],
        description: 'Paper size (set_page_layout).',
      },
      leftMargin: {
        type: 'number',
        required: false,
        description: 'Left margin inches (set_page_layout).',
      },
      rightMargin: { type: 'number', required: false, description: 'Right margin inches.' },
      topMargin: { type: 'number', required: false, description: 'Top margin inches.' },
      bottomMargin: { type: 'number', required: false, description: 'Bottom margin inches.' },
      recalcType: {
        type: 'string',
        required: false,
        enum: ['Recalculate', 'Full'],
        description: 'Recalc type (recalculate).',
      },
    },
    execute: async (context, args) => {
      const action = args.action as string;

      if (action === 'list') {
        const sheets = context.workbook.worksheets;
        sheets.load('items');
        const active = context.workbook.worksheets.getActiveWorksheet();
        active.load('name');
        await context.sync();
        for (const s of sheets.items) s.load(['name', 'id', 'position', 'visibility']);
        await context.sync();
        return {
          sheets: sheets.items.map(s => ({
            name: s.name,
            id: s.id,
            position: s.position,
            visibility: s.visibility,
            isActive: s.name === active.name,
          })),
          count: sheets.items.length,
        };
      }

      if (action === 'create') {
        const sheet = context.workbook.worksheets.add(args.name as string);
        sheet.load(['name', 'id', 'position']);
        await context.sync();
        return { name: sheet.name, id: sheet.id, position: sheet.position };
      }

      if (action === 'delete') {
        const name = args.name as string;
        context.workbook.worksheets.getItem(name).delete();
        await context.sync();
        return { deleted: name };
      }

      if (action === 'rename') {
        const sheet = context.workbook.worksheets.getItem(args.currentName as string);
        sheet.name = args.newName as string;
        sheet.load('name');
        await context.sync();
        return { previousName: args.currentName, newName: sheet.name };
      }

      if (action === 'copy') {
        const src = context.workbook.worksheets.getItem(args.name as string);
        const copied = src.copy('After', src);
        if (args.newName) copied.name = args.newName as string;
        copied.load(['name', 'id', 'position']);
        await context.sync();
        return { sourceSheet: args.name, copiedSheet: copied.name, position: copied.position };
      }

      if (action === 'move') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        sheet.position = args.position as number;
        sheet.load(['name', 'position']);
        await context.sync();
        return { name: sheet.name, position: sheet.position };
      }

      if (action === 'activate') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        sheet.activate();
        sheet.load('name');
        await context.sync();
        return { activated: sheet.name };
      }

      if (action === 'protect') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        const pw = args.password as string | undefined;
        if (pw) {
          sheet.protection.protect({ allowAutoFilter: true, allowSort: true }, pw);
        } else {
          sheet.protection.protect({ allowAutoFilter: true, allowSort: true });
        }
        await context.sync();
        return { sheet: args.name, protected: true };
      }

      if (action === 'unprotect') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        sheet.protection.unprotect(args.password as string | undefined);
        await context.sync();
        return { sheet: args.name, protected: false };
      }

      if (action === 'freeze') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        const freezeAt = args.freezeAt as string | undefined;
        if (freezeAt) {
          sheet.freezePanes.freezeAt(sheet.getRange(freezeAt));
        } else {
          sheet.freezePanes.unfreeze();
        }
        await context.sync();
        return { sheet: args.name, frozenAt: freezeAt ?? null, unfrozen: !freezeAt };
      }

      if (action === 'set_visibility') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        if (args.visibility !== undefined)
          sheet.visibility = args.visibility as Excel.SheetVisibility;
        if (args.tabColor !== undefined) sheet.tabColor = args.tabColor as string;
        sheet.load(['name', 'visibility', 'tabColor']);
        await context.sync();
        return { name: sheet.name, visibility: sheet.visibility, tabColor: sheet.tabColor };
      }

      if (action === 'set_gridlines') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        sheet.showGridlines = args.showGridlines as boolean;
        sheet.load(['name', 'showGridlines']);
        await context.sync();
        return { name: sheet.name, showGridlines: sheet.showGridlines };
      }

      if (action === 'set_headings') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        sheet.showHeadings = args.showHeadings as boolean;
        sheet.load(['name', 'showHeadings']);
        await context.sync();
        return { name: sheet.name, showHeadings: sheet.showHeadings };
      }

      if (action === 'set_page_layout') {
        const sheet = context.workbook.worksheets.getItem(args.name as string);
        const pl = sheet.pageLayout;
        if (args.orientation) pl.orientation = args.orientation as Excel.PageOrientation;
        if (args.paperSize) pl.paperSize = args.paperSize as Excel.PaperType;
        if (args.leftMargin !== undefined) pl.leftMargin = args.leftMargin as number;
        if (args.rightMargin !== undefined) pl.rightMargin = args.rightMargin as number;
        if (args.topMargin !== undefined) pl.topMargin = args.topMargin as number;
        if (args.bottomMargin !== undefined) pl.bottomMargin = args.bottomMargin as number;
        pl.load([
          'orientation',
          'paperSize',
          'leftMargin',
          'rightMargin',
          'topMargin',
          'bottomMargin',
        ]);
        await context.sync();
        return {
          sheet: args.name,
          orientation: pl.orientation,
          paperSize: pl.paperSize,
          margins: {
            left: pl.leftMargin,
            right: pl.rightMargin,
            top: pl.topMargin,
            bottom: pl.bottomMargin,
          },
        };
      }

      // recalculate
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const recalcType = (args.recalcType as string) ?? 'Full';
      sheet.calculate(recalcType === 'Full');
      await context.sync();
      return { name: args.name, recalculated: true, type: recalcType };
    },
  },
];
