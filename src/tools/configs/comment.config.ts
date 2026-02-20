/**
 * Comment tool configs — 1 tool (comment) with actions: list, add, edit, delete.
 */

import type { ToolConfig } from '../codegen';
import { getSheet } from '../codegen';

export const commentConfigs: readonly ToolConfig[] = [
  {
    name: 'comment',
    description:
      'Manage cell comments. Use action "list" to list all comments, "add" to create, "edit" to update, "delete" to remove.',
    params: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['list', 'add', 'edit', 'delete'],
      },
      cellAddress: {
        type: 'string',
        required: false,
        description: 'Cell address (e.g. "A1"). Required for add/edit/delete.',
      },
      text: {
        type: 'string',
        required: false,
        description: 'Comment text. Required for add; replacement text for edit.',
      },
      sheetName: { type: 'string', required: false, description: 'Optional worksheet name.' },
    },
    execute: async (context, args) => {
      const action = args.action as string;
      const sheet = getSheet(context, args.sheetName as string | undefined);

      if (action === 'list') {
        const comments = sheet.comments;
        comments.load('items');
        await context.sync();
        const locs: Excel.Range[] = [];
        for (const c of comments.items) {
          c.load(['content', 'authorName', 'creationDate']);
          const loc = c.getLocation();
          loc.load('address');
          locs.push(loc);
        }
        await context.sync();
        const result = comments.items.map((c, i) => ({
          content: c.content,
          authorName: c.authorName,
          creationDate: c.creationDate,
          cellAddress: locs[i].address,
        }));
        return { comments: result, count: result.length };
      }

      if (action === 'add') {
        sheet.comments.add(args.cellAddress as string, args.text as string);
        await context.sync();
        return { action, cellAddress: args.cellAddress, added: true };
      }

      // edit / delete — need sheet!cell format
      let sheetRef = args.sheetName as string | undefined;
      if (!sheetRef) {
        sheet.load('name');
        await context.sync();
        sheetRef = sheet.name;
      }
      const cell = args.cellAddress as string;
      const fullAddr = cell.includes('!') ? cell : `${sheetRef}!${cell}`;
      const comment = context.workbook.comments.getItemByCell(fullAddr);

      if (action === 'edit') {
        comment.content = args.text as string;
        await context.sync();
        return { action, cellAddress: cell, updated: true };
      }
      comment.delete();
      await context.sync();
      return { action, cellAddress: cell, deleted: true };
    },
  },
];
