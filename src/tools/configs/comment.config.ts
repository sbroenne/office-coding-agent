/**
 * Comment tool configs — 4 tools for managing cell comments.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const commentConfigs: readonly ToolConfig[] = [
  {
    name: 'add_comment',
    description:
      'Add a threaded comment to a specific cell. The comment appears in the Comments pane and is attributed to the current user. Only one comment thread per cell is allowed.',
    params: {
      cellAddress: {
        type: 'string',
        description: 'The single cell address to attach the comment to (e.g., "A1", "B5")',
      },
      text: { type: 'string', description: 'The comment text content' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      sheet.comments.add(args.cellAddress as string, args.text as string);
      await context.sync();
      return { cellAddress: args.cellAddress, text: args.text, added: true };
    },
  },

  {
    name: 'list_comments',
    description:
      "List all comments on a worksheet. Returns each comment's text, author, creation date, and cell address.",
    params: {
      sheetName: {
        type: 'string',
        required: false,
        description: 'Optional worksheet name. Uses active sheet if omitted.',
      },
    },
    execute: async (context, args) => {
      const sheet = getSheet(context, args.sheetName as string | undefined);
      const comments = sheet.comments;
      comments.load('items');
      await context.sync();

      // Load per-item properties and cell address — save proxy refs for safe reuse
      const locationProxies: Excel.Range[] = [];
      for (const comment of comments.items) {
        comment.load(['content', 'authorName', 'creationDate']);
        const loc = comment.getLocation();
        loc.load('address');
        locationProxies.push(loc);
      }
      await context.sync();

      const result = comments.items.map((comment, i) => ({
        content: comment.content,
        authorName: comment.authorName,
        creationDate: comment.creationDate,
        cellAddress: locationProxies[i].address,
      }));
      return { comments: result, count: result.length };
    },
  },

  {
    name: 'edit_comment',
    description:
      'Edit the text of an existing comment on a specific cell. Replaces the entire comment content.',
    params: {
      cellAddress: {
        type: 'string',
        description: 'The cell address of the comment to edit (e.g., "A1")',
      },
      newText: { type: 'string', description: 'The new text for the comment' },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const cellAddress = args.cellAddress as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);
      let sheetRef = args.sheetName as string | undefined;
      if (!sheetRef) {
        sheet.load('name');
        await context.sync();
        sheetRef = sheet.name;
      }
      const fullAddress = cellAddress.includes('!') ? cellAddress : `${sheetRef}!${cellAddress}`;
      const comment = context.workbook.comments.getItemByCell(fullAddress);
      comment.content = args.newText as string;
      await context.sync();
      return { cellAddress, newText: args.newText, updated: true };
    },
  },

  {
    name: 'delete_comment',
    description: 'Delete a comment and all its replies from a specific cell.',
    params: {
      cellAddress: {
        type: 'string',
        description: 'The cell address of the comment to delete (e.g., "A1")',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const cellAddress = args.cellAddress as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);
      let sheetRef = args.sheetName as string | undefined;
      if (!sheetRef) {
        sheet.load('name');
        await context.sync();
        sheetRef = sheet.name;
      }
      const fullAddress = cellAddress.includes('!') ? cellAddress : `${sheetRef}!${cellAddress}`;
      const comment = context.workbook.comments.getItemByCell(fullAddress);
      comment.delete();
      await context.sync();
      return { cellAddress, deleted: true };
    },
  },
];
