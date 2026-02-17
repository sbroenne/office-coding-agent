/**
 * Sheet tool configs — 13 tools for managing worksheets.
 */

import type { ToolConfig } from '../codegen';

export const sheetConfigs: readonly ToolConfig[] = [
  {
    name: 'list_sheets',
    description:
      "List all worksheets in the workbook. Returns each sheet's name, position, visibility status, and whether it is the active sheet.",
    params: {},
    execute: async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      await context.sync();

      for (const s of sheets.items) {
        s.load(['name', 'id', 'position', 'visibility']);
      }
      await context.sync();

      const result = sheets.items.map(s => ({
        name: s.name,
        id: s.id,
        position: s.position,
        visibility: s.visibility,
        isActive: s.name === activeSheet.name,
      }));
      return { sheets: result, count: result.length };
    },
  },

  {
    name: 'create_sheet',
    description:
      'Create a new blank worksheet and add it to the workbook. The sheet becomes the active sheet.',
    params: {
      name: { type: 'string', description: 'Name for the new worksheet' },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.add(args.name as string);
      sheet.load(['name', 'id', 'position']);
      await context.sync();
      return { name: sheet.name, id: sheet.id, position: sheet.position };
    },
  },

  {
    name: 'rename_sheet',
    description: 'Rename an existing worksheet.',
    params: {
      currentName: { type: 'string', description: 'Current name of the worksheet' },
      newName: { type: 'string', description: 'New name for the worksheet' },
    },
    execute: async (context, args) => {
      const currentName = args.currentName as string;
      const sheet = context.workbook.worksheets.getItem(currentName);
      sheet.name = args.newName as string;
      sheet.load('name');
      await context.sync();
      return { previousName: currentName, newName: sheet.name };
    },
  },

  {
    name: 'delete_sheet',
    description:
      'Permanently delete a worksheet and all its data from the workbook. This cannot be undone.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to delete' },
    },
    execute: async (context, args) => {
      const name = args.name as string;
      const sheet = context.workbook.worksheets.getItem(name);
      sheet.delete();
      await context.sync();
      return { deleted: name };
    },
  },

  {
    name: 'activate_sheet',
    description: 'Activate (switch to) a specific worksheet.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to activate' },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      sheet.activate();
      sheet.load('name');
      await context.sync();
      return { activated: sheet.name };
    },
  },

  // ─── Freeze Panes ────────────────────────────────────────

  {
    name: 'freeze_panes',
    description:
      'Freeze rows and/or columns on a worksheet so they remain visible when scrolling. Specify a cell address to freeze all rows above and columns to the left of that cell. Use "unfreeze" to remove all frozen panes.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      freezeAt: {
        type: 'string',
        required: false,
        description:
          'Cell address to freeze at (e.g., "B3" freezes row 1-2 and column A). Omit to unfreeze all panes.',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const freezeAt = args.freezeAt as string | undefined;
      if (freezeAt) {
        const range = sheet.getRange(freezeAt);
        sheet.freezePanes.freezeAt(range);
      } else {
        sheet.freezePanes.unfreeze();
      }
      await context.sync();
      return { sheet: args.name, frozenAt: freezeAt ?? null, unfrozen: !freezeAt };
    },
  },

  // ─── Sheet Protection ────────────────────────────────────

  {
    name: 'protect_sheet',
    description:
      'Protect a worksheet to prevent editing. Optionally provide a password. Protected sheets block cell edits, row/column insertions, and other modifications.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to protect' },
      password: {
        type: 'string',
        required: false,
        description: 'Optional password to protect the sheet with',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const password = args.password as string | undefined;
      if (password) {
        sheet.protection.protect({ allowAutoFilter: true, allowSort: true }, password);
      } else {
        sheet.protection.protect({ allowAutoFilter: true, allowSort: true });
      }
      await context.sync();
      return { sheet: args.name, protected: true };
    },
  },

  {
    name: 'unprotect_sheet',
    description: 'Remove protection from a worksheet. Provide the password if one was set.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to unprotect' },
      password: {
        type: 'string',
        required: false,
        description: 'Password used when protecting (if any)',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      sheet.protection.unprotect(args.password as string | undefined);
      await context.sync();
      return { sheet: args.name, protected: false };
    },
  },

  // ─── Sheet Visibility & Tab Color ────────────────────────

  {
    name: 'set_sheet_visibility',
    description:
      'Set sheet visibility (visible, hidden, or very hidden) and/or tab color. Very hidden sheets cannot be unhidden from the Excel UI — only via code.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      visibility: {
        type: 'string',
        required: false,
        description: 'Visibility state for the sheet',
        enum: ['Visible', 'Hidden', 'VeryHidden'],
      },
      tabColor: {
        type: 'string',
        required: false,
        description: 'Tab color as a hex string (e.g., "#FF0000" for red). Use "" to clear.',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const visibility = args.visibility as string | undefined;
      const tabColor = args.tabColor as string | undefined;
      if (visibility !== undefined) {
        sheet.visibility = visibility as Excel.SheetVisibility;
      }
      if (tabColor !== undefined) {
        sheet.tabColor = tabColor;
      }
      sheet.load(['name', 'visibility', 'tabColor']);
      await context.sync();
      return { name: sheet.name, visibility: sheet.visibility, tabColor: sheet.tabColor };
    },
  },

  // ─── Copy / Move Sheet ───────────────────────────────────

  {
    name: 'copy_sheet',
    description:
      'Create a copy of a worksheet. The copy is placed after the source sheet by default.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to copy' },
      newName: {
        type: 'string',
        required: false,
        description: 'Name for the copied sheet. If omitted, Excel generates a default name.',
      },
    },
    execute: async (context, args) => {
      const name = args.name as string;
      const sheet = context.workbook.worksheets.getItem(name);
      const copied = sheet.copy('After', sheet);
      const newName = args.newName as string | undefined;
      if (newName) {
        copied.name = newName;
      }
      copied.load(['name', 'id', 'position']);
      await context.sync();
      return { sourceSheet: name, copiedSheet: copied.name, position: copied.position };
    },
  },

  {
    name: 'move_sheet',
    description:
      'Move a worksheet to a new position in the workbook. Position is 0-based (0 = first).',
    params: {
      name: { type: 'string', description: 'Name of the worksheet to move' },
      position: { type: 'number', description: 'New 0-based position index' },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      sheet.position = args.position as number;
      sheet.load(['name', 'position']);
      await context.sync();
      return { name: sheet.name, position: sheet.position };
    },
  },

  // ─── Page Layout ───────────────────────────────────────────

  {
    name: 'set_page_layout',
    description:
      'Configure page layout settings for printing: orientation (portrait/landscape), margins, paper size.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      orientation: {
        type: 'string',
        required: false,
        description: 'Page orientation',
        enum: ['Portrait', 'Landscape'],
      },
      paperSize: {
        type: 'string',
        required: false,
        description: 'Paper size (letter, legal, A4, etc.)',
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
      },
      leftMargin: {
        type: 'number',
        required: false,
        description: 'Left margin in inches',
      },
      rightMargin: {
        type: 'number',
        required: false,
        description: 'Right margin in inches',
      },
      topMargin: {
        type: 'number',
        required: false,
        description: 'Top margin in inches',
      },
      bottomMargin: {
        type: 'number',
        required: false,
        description: 'Bottom margin in inches',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const pageLayout = sheet.pageLayout;

      if (args.orientation) {
        pageLayout.orientation = args.orientation as Excel.PageOrientation;
      }
      if (args.paperSize) {
        pageLayout.paperSize = args.paperSize as Excel.PaperType;
      }

      // PageLayout has direct margin properties (no sub-object)
      if (args.leftMargin !== undefined) pageLayout.leftMargin = args.leftMargin as number;
      if (args.rightMargin !== undefined) pageLayout.rightMargin = args.rightMargin as number;
      if (args.topMargin !== undefined) pageLayout.topMargin = args.topMargin as number;
      if (args.bottomMargin !== undefined) pageLayout.bottomMargin = args.bottomMargin as number;

      pageLayout.load([
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
        orientation: pageLayout.orientation,
        paperSize: pageLayout.paperSize,
        margins: {
          left: pageLayout.leftMargin,
          right: pageLayout.rightMargin,
          top: pageLayout.topMargin,
          bottom: pageLayout.bottomMargin,
        },
      };
    },
  },

  {
    name: 'set_sheet_gridlines',
    description: 'Show or hide worksheet gridlines for a specific sheet.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      showGridlines: { type: 'boolean', description: 'True to show gridlines, false to hide' },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      sheet.showGridlines = args.showGridlines as boolean;
      sheet.load(['name', 'showGridlines']);
      await context.sync();
      return { name: sheet.name, showGridlines: sheet.showGridlines };
    },
  },

  {
    name: 'set_sheet_headings',
    description: 'Show or hide row/column headings (A, B, C and 1, 2, 3) for a worksheet.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      showHeadings: {
        type: 'boolean',
        description: 'True to show row/column headings, false to hide',
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      sheet.showHeadings = args.showHeadings as boolean;
      sheet.load(['name', 'showHeadings']);
      await context.sync();
      return { name: sheet.name, showHeadings: sheet.showHeadings };
    },
  },

  {
    name: 'recalculate_sheet',
    description:
      'Force recalculation of formulas on a specific worksheet. Defaults to full recalculation.',
    params: {
      name: { type: 'string', description: 'Name of the worksheet' },
      recalcType: {
        type: 'string',
        required: false,
        description: 'Recalculation type',
        enum: ['Recalculate', 'Full'],
      },
    },
    execute: async (context, args) => {
      const sheet = context.workbook.worksheets.getItem(args.name as string);
      const recalcType = (args.recalcType as string) ?? 'Full';
      const markAllDirty = recalcType === 'Full';
      sheet.calculate(markAllDirty);
      await context.sync();
      return { name: args.name, recalculated: true, type: recalcType };
    },
  },
];
