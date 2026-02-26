import type { WordToolConfig } from '../codegen';
import { createWordTools } from '../codegen';

// ΓöÇΓöÇΓöÇ Tool Configs ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ

export const wordConfigs: readonly WordToolConfig[] = [
  {
    name: 'get_document_overview',
    description:
      'Get a structural overview of the Word document including heading hierarchy, paragraph count, tables, and content controls. Use this first to understand the document structure.',
    params: {},
    execute: async context => {
      const body = context.document.body;
      body.load('text');

      const paragraphs = body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const para of paragraphs.items) {
        para.load(['text', 'style', 'isListItem']);
      }
      await context.sync();

      const headings: string[] = [];
      let paragraphCount = 0;
      let tableCount = 0;

      for (const para of paragraphs.items) {
        paragraphCount++;
        const style = para.style ?? '';
        if (style.startsWith('Heading')) {
          const level = style.replace('Heading ', 'H');
          headings.push(`${level}: ${para.text.trim()}`);
        }
      }

      const tables = body.tables;
      tables.load('items');
      await context.sync();
      tableCount = tables.items.length;

      return [
        `Document Overview`,
        `${'='.repeat(40)}`,
        `Paragraphs: ${String(paragraphCount)}`,
        `Tables: ${String(tableCount)}`,
        '',
        headings.length > 0 ? `Headings:\n${headings.join('\n')}` : '(no headings found)',
      ].join('\n');
    },
  },

  {
    name: 'get_document_content',
    description:
      'Get the full HTML content of the Word document body. Returns rich formatted content that preserves structure such as headings, lists, and tables.',
    params: {},
    execute: async context => {
      const body = context.document.body;
      const htmlResult = body.getHtml();
      await context.sync();
      return htmlResult.value;
    },
  },

  {
    name: 'get_document_section',
    description:
      'Get the HTML content of a specific section identified by a heading. Finds the heading by partial text match and returns the content until the next heading of the same or higher level.',
    params: {
      headingText: {
        type: 'string',
        description: 'Partial or full text of the heading that starts the section.',
      },
    },
    execute: async (context, args) => {
      const { headingText } = args as { headingText: string };

      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const para of paragraphs.items) {
        para.load(['text', 'style']);
      }
      await context.sync();

      const headingPara = paragraphs.items.find(
        p =>
          p.style?.startsWith('Heading') && p.text.toLowerCase().includes(headingText.toLowerCase())
      );

      if (!headingPara) {
        return `No heading found containing "${headingText}".`;
      }

      const headingLevel = parseInt(headingPara.style.replace('Heading ', ''), 10) || 1;
      const range = headingPara.getRange();

      const allParas = paragraphs.items;
      const startIdx = allParas.indexOf(headingPara);

      let endPara: Word.Paragraph | undefined;
      for (let i = startIdx + 1; i < allParas.length; i++) {
        const p = allParas[i];
        if (p.style?.startsWith('Heading')) {
          const level = parseInt(p.style.replace('Heading ', ''), 10) || 1;
          if (level <= headingLevel) {
            endPara = p;
            break;
          }
        }
      }

      let sectionRange: Word.Range;
      if (endPara) {
        const endRange = endPara.getRange(Word.RangeLocation.start);
        sectionRange = range.expandTo(endRange);
      } else {
        const bodyEnd = body.getRange(Word.RangeLocation.end);
        sectionRange = range.expandTo(bodyEnd);
      }

      const htmlResult = sectionRange.getHtml();
      await context.sync();
      return htmlResult.value;
    },
  },

  {
    name: 'set_document_content',
    description:
      'Replace the entire document body with new HTML content. WARNING: This clears all existing content.',
    params: {
      html: {
        type: 'string',
        description: 'HTML content to set as the full document body.',
      },
    },
    execute: async (context, args) => {
      const { html } = args as { html: string };

      const body = context.document.body;
      body.clear();
      body.insertHtml(html, Word.InsertLocation.start);
      await context.sync();
      return 'Document content replaced successfully.';
    },
  },

  {
    name: 'get_selection',
    description:
      'Get the currently selected content in the Word document as OOXML. Useful for inspecting the structure of the selection.',
    params: {},
    execute: async context => {
      const selection = context.document.getSelection();
      const ooxmlResult = selection.getOoxml();
      await context.sync();

      const ooxml = ooxmlResult.value;
      const docMatch = /<w:document[^>]*>[\s\S]*<\/w:document>/i.exec(ooxml);
      return docMatch ? docMatch[0] : ooxml;
    },
  },

  {
    name: 'get_selection_text',
    description: 'Get the plain text of the currently selected content in the Word document.',
    params: {},
    execute: async context => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();
      return selection.text.length > 0 ? selection.text : '(no text selected)';
    },
  },

  {
    name: 'insert_content_at_selection',
    description:
      'Insert HTML content at or relative to the current selection in the Word document.',
    params: {
      html: { type: 'string', description: 'HTML content to insert.' },
      location: {
        type: 'string',
        required: false,
        enum: ['Replace', 'Before', 'After', 'Start', 'End'],
        description:
          'Where to insert relative to the selection. Replace overwrites the selection. Default: Replace.',
      },
    },
    execute: async (context, args) => {
      const { html, location = 'Replace' } = args as { html: string; location?: string };

      const locationMap: Record<string, Word.InsertLocation> = {
        Replace: Word.InsertLocation.replace,
        Before: Word.InsertLocation.before,
        After: Word.InsertLocation.after,
        Start: Word.InsertLocation.start,
        End: Word.InsertLocation.end,
      };

      const insertLocation = locationMap[location] ?? Word.InsertLocation.replace;

      const selection = context.document.getSelection();
      selection.insertHtml(html, insertLocation);
      await context.sync();
      return `Content inserted at selection (location: ${location}).`;
    },
  },

  {
    name: 'find_and_replace',
    description: 'Find text in the Word document and replace all occurrences with new text.',
    params: {
      find: { type: 'string', description: 'Text to search for.' },
      replace: { type: 'string', description: 'Replacement text.' },
      matchCase: {
        type: 'boolean',
        required: false,
        description: 'Whether the search is case-sensitive. Default: false.',
      },
      matchWholeWord: {
        type: 'boolean',
        required: false,
        description: 'Whether to match whole words only. Default: false.',
      },
    },
    execute: async (context, args) => {
      const {
        find,
        replace,
        matchCase = false,
        matchWholeWord = false,
      } = args as {
        find: string;
        replace: string;
        matchCase?: boolean;
        matchWholeWord?: boolean;
      };

      const body = context.document.body;
      const results = body.search(find, { matchCase, matchWholeWord });
      results.load('items');
      await context.sync();

      const count = results.items.length;
      if (count === 0) {
        return `No occurrences of "${find}" found.`;
      }

      for (const result of results.items) {
        result.insertText(replace, Word.InsertLocation.replace);
      }
      await context.sync();

      return `Replaced ${String(count)} occurrence(s) of "${find}" with "${replace}".`;
    },
  },

  {
    name: 'insert_table',
    description:
      'Insert a table at the current selection in the Word document. Supports grid, striped, and plain styles.',
    params: {
      rows: { type: 'number', description: 'Number of rows.' },
      columns: { type: 'number', description: 'Number of columns.' },
      data: {
        type: 'string[][]',
        required: false,
        description: 'Optional 2D array of cell values (row-major). Omit for an empty table.',
      },
      style: {
        type: 'string',
        required: false,
        enum: ['grid', 'striped', 'plain'],
        description:
          'Visual style. grid adds borders, striped alternates row colors, plain has no formatting. Default: grid.',
      },
      hasHeaderRow: {
        type: 'boolean',
        required: false,
        description: 'Whether the first row is a header row with distinct styling. Default: true.',
      },
    },
    execute: async (context, args) => {
      const {
        rows,
        columns,
        data,
        style = 'grid',
        hasHeaderRow = true,
      } = args as {
        rows: number;
        columns: number;
        data?: string[][];
        style?: 'grid' | 'striped' | 'plain';
        hasHeaderRow?: boolean;
      };

      const selection = context.document.getSelection();

      const cellValues: string[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: string[] = [];
        for (let c = 0; c < columns; c++) {
          row.push(data?.[r]?.[c] ?? '');
        }
        cellValues.push(row);
      }

      const table = selection.insertTable(rows, columns, Word.InsertLocation.after, cellValues);
      await context.sync();

      table.rows.load('items');
      await context.sync();

      if (style === 'grid') {
        table.style = 'Table Grid';
      }

      if (hasHeaderRow && rows > 0) {
        const headerRow = table.rows.items[0];
        headerRow.cells.load('items');
        await context.sync();

        for (const cell of headerRow.cells.items) {
          cell.shadingColor = '#4472C4';
          const paras = cell.body.paragraphs;
          paras.load('items');
          await context.sync();
          for (const para of paras.items) {
            para.font.bold = true;
            para.font.color = '#FFFFFF';
          }
        }
      }

      if (style === 'striped') {
        for (let r = hasHeaderRow ? 1 : 0; r < table.rows.items.length; r++) {
          if (r % 2 === 0) continue;
          const row = table.rows.items[r];
          row.cells.load('items');
          await context.sync();
          for (const cell of row.cells.items) {
            cell.shadingColor = '#E8E8E8';
          }
        }
      }

      await context.sync();
      return `Inserted a ${String(rows)}├ù${String(columns)} table with "${style}" style.`;
    },
  },

  {
    name: 'apply_style_to_selection',
    description: 'Apply font formatting to the currently selected text in the Word document.',
    params: {
      bold: { type: 'boolean', required: false, description: 'Make text bold.' },
      italic: { type: 'boolean', required: false, description: 'Make text italic.' },
      underline: {
        type: 'boolean',
        required: false,
        description: 'Underline text. true = single underline.',
      },
      strikeThrough: {
        type: 'boolean',
        required: false,
        description: 'Apply strikethrough.',
      },
      fontSize: { type: 'number', required: false, description: 'Font size in points.' },
      fontName: {
        type: 'string',
        required: false,
        description: 'Font family name (e.g. "Calibri", "Arial").',
      },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color as hex (e.g. "#FF0000") or named color.',
      },
      highlightColor: {
        type: 'string',
        required: false,
        description:
          'Highlight color. Allowed values: Yellow, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, Black, White, None.',
      },
    },
    execute: async (context, args) => {
      const {
        bold,
        italic,
        underline,
        strikeThrough,
        fontSize,
        fontName,
        fontColor,
        highlightColor,
      } = args as {
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        strikeThrough?: boolean;
        fontSize?: number;
        fontName?: string;
        fontColor?: string;
        highlightColor?: string;
      };

      const selection = context.document.getSelection();
      const font = selection.font;

      if (bold !== undefined) font.bold = bold;
      if (italic !== undefined) font.italic = italic;
      if (underline !== undefined)
        font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      if (strikeThrough !== undefined) font.strikeThrough = strikeThrough;
      if (fontSize !== undefined) font.size = fontSize;
      if (fontName !== undefined) font.name = fontName;
      if (fontColor !== undefined) font.color = fontColor;
      if (highlightColor !== undefined) font.highlightColor = highlightColor;

      await context.sync();

      const applied: string[] = [];
      if (bold !== undefined) applied.push(`bold=${String(bold)}`);
      if (italic !== undefined) applied.push(`italic=${String(italic)}`);
      if (underline !== undefined) applied.push(`underline=${String(underline)}`);
      if (strikeThrough !== undefined) applied.push(`strikeThrough=${String(strikeThrough)}`);
      if (fontSize !== undefined) applied.push(`fontSize=${String(fontSize)}`);
      if (fontName !== undefined) applied.push(`fontName="${fontName}"`);
      if (fontColor !== undefined) applied.push(`color="${fontColor}"`);
      if (highlightColor !== undefined) applied.push(`highlight="${highlightColor}"`);

      return applied.length > 0
        ? `Applied: ${applied.join(', ')}.`
        : 'No style properties specified.';
    },
  },

  // ─── Additional tools from PR #33 ────────────────────────────────────────

  {
    name: 'insert_paragraph',
    description:
      'Insert a paragraph at the end or beginning of the document body. Optionally apply a named style such as "Heading 1" or "Normal".',
    params: {
      text: { type: 'string', description: 'The paragraph text to insert.' },
      location: {
        type: 'string',
        required: false,
        enum: ['End', 'Start'],
        description: 'Where to insert the paragraph. Default: End.',
      },
      style: {
        type: 'string',
        required: false,
        description: 'Optional Word style name to apply (e.g. "Heading 1", "Normal", "Title").',
      },
    },
    execute: async (context, args) => {
      const { text, location = 'End', style } = args as {
        text: string;
        location?: string;
        style?: string;
      };
      const body = context.document.body;
      const insertLoc =
        location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end;
      const paragraph = body.insertParagraph(text, insertLoc);
      if (style) paragraph.style = style;
      await context.sync();
      return `Paragraph inserted at ${location}${style ? ` with style "${style}"` : ''}.`;
    },
  },

  {
    name: 'insert_break',
    description: 'Insert a page break or section break after the current selection.',
    params: {
      breakType: {
        type: 'string',
        required: false,
        enum: ['page', 'sectionNext', 'sectionContinuous'],
        description: 'Type of break to insert. Default: page.',
      },
    },
    execute: async (context, args) => {
      const { breakType = 'page' } = args as { breakType?: string };
      const breakMap: Record<string, Word.BreakType> = {
        page: Word.BreakType.page,
        sectionNext: Word.BreakType.sectionNext,
        sectionContinuous: Word.BreakType.sectionContinuous,
      };
      const wordBreakType = breakMap[breakType] ?? Word.BreakType.page;
      const selection = context.document.getSelection();
      selection.insertBreak(wordBreakType, Word.InsertLocation.after);
      await context.sync();
      return `Inserted ${breakType} break after selection.`;
    },
  },

  {
    name: 'apply_paragraph_style',
    description:
      'Apply a named style (e.g. "Heading 1", "Title", "Normal", "Quote") to every paragraph in the current selection.',
    params: {
      styleName: {
        type: 'string',
        description:
          'The Word style name to apply (e.g. "Heading 1", "Title", "Normal", "Quote").',
      },
    },
    execute: async (context, args) => {
      const { styleName } = args as { styleName: string };
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load('items');
      await context.sync();
      for (const para of paragraphs.items) {
        para.style = styleName;
      }
      await context.sync();
      return `Applied style "${styleName}" to ${String(paragraphs.items.length)} paragraph(s).`;
    },
  },

  {
    name: 'set_paragraph_format',
    description:
      'Set paragraph formatting on the current selection. Only provided properties are changed; others are preserved.',
    params: {
      alignment: {
        type: 'string',
        required: false,
        enum: ['left', 'center', 'right', 'justified'],
        description: 'Horizontal alignment.',
      },
      lineSpacing: {
        type: 'number',
        required: false,
        description: 'Line spacing in points.',
      },
      spaceBefore: {
        type: 'number',
        required: false,
        description: 'Space before paragraph in points.',
      },
      spaceAfter: {
        type: 'number',
        required: false,
        description: 'Space after paragraph in points.',
      },
      firstLineIndent: {
        type: 'number',
        required: false,
        description: 'First line indent in points.',
      },
    },
    execute: async (context, args) => {
      const { alignment, lineSpacing, spaceBefore, spaceAfter, firstLineIndent } = args as {
        alignment?: string;
        lineSpacing?: number;
        spaceBefore?: number;
        spaceAfter?: number;
        firstLineIndent?: number;
      };
      const alignmentMap: Record<string, Word.Alignment> = {
        left: Word.Alignment.left,
        center: Word.Alignment.centered,
        right: Word.Alignment.right,
        justified: Word.Alignment.justified,
      };
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load('items');
      await context.sync();
      for (const para of paragraphs.items) {
        if (alignment !== undefined && alignmentMap[alignment] !== undefined) {
          para.alignment = alignmentMap[alignment];
        }
        if (lineSpacing !== undefined) para.lineSpacing = lineSpacing;
        if (spaceBefore !== undefined) para.spaceBefore = spaceBefore;
        if (spaceAfter !== undefined) para.spaceAfter = spaceAfter;
        if (firstLineIndent !== undefined) para.firstLineIndent = firstLineIndent;
      }
      await context.sync();
      const applied: string[] = [];
      if (alignment !== undefined) applied.push(`alignment=${alignment}`);
      if (lineSpacing !== undefined) applied.push(`lineSpacing=${String(lineSpacing)}`);
      if (spaceBefore !== undefined) applied.push(`spaceBefore=${String(spaceBefore)}`);
      if (spaceAfter !== undefined) applied.push(`spaceAfter=${String(spaceAfter)}`);
      if (firstLineIndent !== undefined)
        applied.push(`firstLineIndent=${String(firstLineIndent)}`);
      return applied.length > 0
        ? `Applied to ${String(paragraphs.items.length)} paragraph(s): ${applied.join(', ')}.`
        : 'No formatting properties specified.';
    },
  },

  {
    name: 'get_document_properties',
    description:
      'Get document metadata including author, title, subject, keywords, creation date, last modified time, and revision number.',
    params: {},
    execute: async context => {
      const props = context.document.properties;
      props.load([
        'author',
        'title',
        'subject',
        'keywords',
        'creationDate',
        'lastSaveTime',
        'lastAuthor',
        'revisionNumber',
        'category',
        'comments',
        'company',
      ]);
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      return [
        'Document Properties',
        '='.repeat(40),
        `Title: ${props.title || '(none)'}`,
        `Author: ${props.author || '(none)'}`,
        `Last Author: ${props.lastAuthor || '(none)'}`,
        `Subject: ${props.subject || '(none)'}`,
        `Keywords: ${props.keywords || '(none)'}`,
        `Category: ${props.category || '(none)'}`,
        `Company: ${props.company || '(none)'}`,
        `Comments: ${props.comments || '(none)'}`,
        `Creation Date: ${String(props.creationDate)}`,
        `Last Save Time: ${String(props.lastSaveTime)}`,
        `Revision Number: ${props.revisionNumber || '(none)'}`,
        `Paragraph Count: ${String(paragraphs.items.length)}`,
      ].join('\n');
    },
  },

  {
    name: 'insert_image',
    description:
      'Insert a base64-encoded image at the current selection. Optionally set width and height in points.',
    params: {
      base64Image: {
        type: 'string',
        description: 'Base64-encoded image data (without data URI prefix).',
      },
      width: { type: 'number', required: false, description: 'Image width in points.' },
      height: { type: 'number', required: false, description: 'Image height in points.' },
    },
    execute: async (context, args) => {
      const { base64Image, width, height } = args as {
        base64Image: string;
        width?: number;
        height?: number;
      };
      const selection = context.document.getSelection();
      const picture = selection.insertInlinePictureFromBase64(
        base64Image,
        Word.InsertLocation.replace
      );
      if (width !== undefined) picture.width = width;
      if (height !== undefined) picture.height = height;
      await context.sync();
      return `Image inserted${width ? ` (width=${String(width)})` : ''}${height ? ` (height=${String(height)})` : ''}.`;
    },
  },

  {
    name: 'get_comments',
    description:
      'Get all comments from the document body, including author, content, creation date, and resolved status.',
    params: {},
    execute: async context => {
      const comments = context.document.body.getComments();
      comments.load('items');
      await context.sync();
      if (comments.items.length === 0) return '(no comments found)';
      for (const comment of comments.items) {
        comment.load(['authorName', 'authorEmail', 'content', 'creationDate', 'resolved', 'id']);
      }
      await context.sync();
      const lines = comments.items.map(
        (c, i) =>
          `${String(i + 1)}. [${c.resolved ? 'Resolved' : 'Open'}] ${c.authorName} (${c.authorEmail}): "${c.content}" — ${String(c.creationDate)}`
      );
      return `Comments (${String(comments.items.length)}):\n${lines.join('\n')}`;
    },
  },

  {
    name: 'insert_list',
    description:
      'Insert a bulleted or numbered list at the current selection using HTML. Provide the list items as a newline-separated string.',
    params: {
      text: {
        type: 'string',
        description: 'List items separated by newlines. Each line becomes a list item.',
      },
      listType: {
        type: 'string',
        required: false,
        enum: ['bullet', 'number'],
        description: 'Type of list to insert. Default: bullet.',
      },
    },
    execute: async (context, args) => {
      const { text, listType = 'bullet' } = args as { text: string; listType?: string };
      const items = text
        .split('\n')
        .map((line: string) => line.trim())
        .filter(Boolean);
      const tag = listType === 'number' ? 'ol' : 'ul';
      const html = `<${tag}>${items.map((item: string) => `<li>${item}</li>`).join('')}</${tag}>`;
      const selection = context.document.getSelection();
      selection.insertHtml(html, Word.InsertLocation.replace);
      await context.sync();
      return `Inserted ${listType === 'number' ? 'numbered' : 'bulleted'} list with ${String(items.length)} item(s).`;
    },
  },

  {
    name: 'get_content_controls',
    description:
      'List all content controls in the document, including their tag, title, text, and type.',
    params: {},
    execute: async context => {
      const controls = context.document.contentControls;
      controls.load('items');
      await context.sync();
      if (controls.items.length === 0) return '(no content controls found)';
      for (const cc of controls.items) {
        cc.load(['tag', 'title', 'text', 'type']);
      }
      await context.sync();
      const lines = controls.items.map(
        (cc, i) =>
          `${String(i + 1)}. [${String(cc.type)}] tag="${cc.tag}" title="${cc.title}" text="${cc.text.length > 100 ? cc.text.slice(0, 100) + '…' : cc.text}"`
      );
      return `Content Controls (${String(controls.items.length)}):\n${lines.join('\n')}`;
    },
  },

  {
    name: 'insert_text_at_bookmark',
    description:
      'Insert text at a named bookmark location in the document. Can replace the bookmark content, or insert before or after it.',
    params: {
      bookmarkName: {
        type: 'string',
        description: 'The name of the bookmark (case-insensitive).',
      },
      text: { type: 'string', description: 'Text to insert.' },
      insertLocation: {
        type: 'string',
        required: false,
        enum: ['Before', 'After', 'Replace'],
        description: 'Where to insert relative to the bookmark. Default: Replace.',
      },
    },
    execute: async (context, args) => {
      const { bookmarkName, text, insertLocation = 'Replace' } = args as {
        bookmarkName: string;
        text: string;
        insertLocation?: string;
      };
      const locationMap: Record<string, Word.InsertLocation> = {
        Before: Word.InsertLocation.before,
        After: Word.InsertLocation.after,
        Replace: Word.InsertLocation.replace,
      };
      const loc = locationMap[insertLocation] ?? Word.InsertLocation.replace;
      const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
      range.load('isNullObject');
      await context.sync();
      if (range.isNullObject) return `Bookmark "${bookmarkName}" not found.`;
      range.insertText(text, loc);
      await context.sync();
      return `Text inserted at bookmark "${bookmarkName}" (location: ${insertLocation}).`;
    },
  },

  {
    name: 'get_headers_footers',
    description:
      'Read headers and footers from all sections of the Word document. Returns section-by-section listing of primary header and footer text.',
    params: {},
    execute: async context => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      const results: string[] = ['Headers & Footers', '='.repeat(40)];
      for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];
        const header = section.getHeader('Primary');
        const footer = section.getFooter('Primary');
        header.load('text');
        footer.load('text');
        await context.sync();
        results.push(
          `\nSection ${String(i + 1)}:`,
          `  Header: ${header.text.trim() || '(empty)'}`,
          `  Footer: ${footer.text.trim() || '(empty)'}`
        );
      }
      return results.join('\n');
    },
  },

  {
    name: 'set_header_footer',
    description:
      'Set header or footer HTML content for a specific section of the Word document.',
    params: {
      sectionIndex: { type: 'number', description: '0-based section index.' },
      type: {
        type: 'string',
        enum: ['header', 'footer'],
        description: 'Whether to set the header or footer.',
      },
      headerFooterType: {
        type: 'string',
        required: false,
        enum: ['Primary', 'FirstPage', 'EvenPages'],
        description: 'Header/footer type. Default: Primary.',
      },
      html: { type: 'string', description: 'HTML content to set.' },
    },
    execute: async (context, args) => {
      const { sectionIndex, type, headerFooterType = 'Primary', html } = args as {
        sectionIndex: number;
        type: 'header' | 'footer';
        headerFooterType?: string;
        html: string;
      };
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      if (sectionIndex < 0 || sectionIndex >= sections.items.length) {
        return `Section index ${String(sectionIndex)} is out of range (0–${String(sections.items.length - 1)}).`;
      }
      const section = sections.items[sectionIndex];
      const hfType = headerFooterType as 'Primary' | 'FirstPage' | 'EvenPages';
      const target = type === 'header' ? section.getHeader(hfType) : section.getFooter(hfType);
      target.clear();
      target.insertHtml(html, Word.InsertLocation.start);
      await context.sync();
      return `Set ${type} (${hfType}) for section ${String(sectionIndex + 1)}.`;
    },
  },

  {
    name: 'get_table_data',
    description:
      'Read the contents of a table by index. Returns the table data as a formatted text grid.',
    params: {
      tableIndex: { type: 'number', description: '0-based table index.' },
    },
    execute: async (context, args) => {
      const { tableIndex } = args as { tableIndex: number };
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        return `Table index ${String(tableIndex)} is out of range (0–${String(tables.items.length - 1)}).`;
      }
      const table = tables.items[tableIndex];
      table.rows.load('items');
      await context.sync();
      const rows: string[][] = [];
      for (const row of table.rows.items) {
        row.cells.load('items');
        await context.sync();
        for (const cell of row.cells.items) {
          cell.body.load('text');
        }
        await context.sync();
        rows.push(row.cells.items.map((cell) => cell.body.text.trim()));
      }
      const lines = rows.map((r, i) => `Row ${String(i)}: ${r.join(' | ')}`);
      return `Table ${String(tableIndex)} (${String(rows.length)} rows):\n${lines.join('\n')}`;
    },
  },

  {
    name: 'add_table_rows',
    description: 'Add rows to an existing table in the Word document.',
    params: {
      tableIndex: { type: 'number', description: '0-based table index.' },
      rowCount: { type: 'number', description: 'Number of rows to add.' },
      insertLocation: {
        type: 'string',
        required: false,
        enum: ['Start', 'End'],
        description: 'Where to add rows. Default: End.',
      },
      values: {
        type: 'string[][]',
        required: false,
        description: 'Optional 2D array of cell values for the new rows.',
      },
    },
    execute: async (context, args) => {
      const { tableIndex, rowCount, insertLocation = 'End', values } = args as {
        tableIndex: number;
        rowCount: number;
        insertLocation?: string;
        values?: string[][];
      };
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        return `Table index ${String(tableIndex)} is out of range (0–${String(tables.items.length - 1)}).`;
      }
      const table = tables.items[tableIndex];
      const loc =
        insertLocation === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end;
      table.addRows(loc, rowCount, values);
      await context.sync();
      return `Added ${String(rowCount)} row(s) at ${insertLocation} of table ${String(tableIndex)}.`;
    },
  },

  {
    name: 'add_table_columns',
    description: 'Add columns to an existing table in the Word document.',
    params: {
      tableIndex: { type: 'number', description: '0-based table index.' },
      columnCount: { type: 'number', description: 'Number of columns to add.' },
      insertLocation: {
        type: 'string',
        required: false,
        enum: ['Start', 'End'],
        description: 'Where to add columns. Default: End.',
      },
      values: {
        type: 'string[][]',
        required: false,
        description: 'Optional 2D array of cell values for the new columns.',
      },
    },
    execute: async (context, args) => {
      const { tableIndex, columnCount, insertLocation = 'End', values } = args as {
        tableIndex: number;
        columnCount: number;
        insertLocation?: string;
        values?: string[][];
      };
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        return `Table index ${String(tableIndex)} is out of range (0–${String(tables.items.length - 1)}).`;
      }
      const table = tables.items[tableIndex];
      const loc =
        insertLocation === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end;
      table.addColumns(loc, columnCount, values);
      await context.sync();
      return `Added ${String(columnCount)} column(s) at ${insertLocation} of table ${String(tableIndex)}.`;
    },
  },

  {
    name: 'delete_table_row',
    description: 'Delete a row from a table in the Word document.',
    params: {
      tableIndex: { type: 'number', description: '0-based table index.' },
      rowIndex: { type: 'number', description: '0-based row index to delete.' },
    },
    execute: async (context, args) => {
      const { tableIndex, rowIndex } = args as { tableIndex: number; rowIndex: number };
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        return `Table index ${String(tableIndex)} is out of range (0–${String(tables.items.length - 1)}).`;
      }
      const table = tables.items[tableIndex];
      table.rows.load('items');
      await context.sync();
      if (rowIndex < 0 || rowIndex >= table.rows.items.length) {
        return `Row index ${String(rowIndex)} is out of range (0–${String(table.rows.items.length - 1)}).`;
      }
      table.rows.items[rowIndex].delete();
      await context.sync();
      return `Deleted row ${String(rowIndex)} from table ${String(tableIndex)}.`;
    },
  },

  {
    name: 'set_table_cell_value',
    description: 'Set the value and optional formatting for a specific cell in a table.',
    params: {
      tableIndex: { type: 'number', description: '0-based table index.' },
      rowIndex: { type: 'number', description: '0-based row index.' },
      cellIndex: { type: 'number', description: '0-based cell (column) index.' },
      text: { type: 'string', description: 'Text to set in the cell.' },
      shadingColor: {
        type: 'string',
        required: false,
        description: 'Optional cell background color as hex (e.g. "#FF0000").',
      },
      bold: {
        type: 'boolean',
        required: false,
        description: 'Optional: make cell text bold.',
      },
    },
    execute: async (context, args) => {
      const { tableIndex, rowIndex, cellIndex, text, shadingColor, bold } = args as {
        tableIndex: number;
        rowIndex: number;
        cellIndex: number;
        text: string;
        shadingColor?: string;
        bold?: boolean;
      };
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      if (tableIndex < 0 || tableIndex >= tables.items.length) {
        return `Table index ${String(tableIndex)} is out of range (0–${String(tables.items.length - 1)}).`;
      }
      const table = tables.items[tableIndex];
      const cell = table.getCell(rowIndex, cellIndex);
      const paras = cell.body.paragraphs;
      paras.load('items');
      await context.sync();
      if (paras.items.length > 0) {
        paras.items[0].insertText(text, Word.InsertLocation.replace);
        if (bold !== undefined) paras.items[0].font.bold = bold;
      }
      if (shadingColor !== undefined) cell.shadingColor = shadingColor;
      await context.sync();
      return `Set cell (${String(rowIndex)}, ${String(cellIndex)}) in table ${String(tableIndex)} to "${text}".`;
    },
  },

  {
    name: 'insert_hyperlink',
    description: 'Insert a hyperlink at the current selection in the Word document.',
    params: {
      url: { type: 'string', description: 'The URL for the hyperlink.' },
      displayText: {
        type: 'string',
        required: false,
        description: 'Optional display text. Defaults to the URL.',
      },
    },
    execute: async (context, args) => {
      const { url, displayText } = args as { url: string; displayText?: string };
      const selection = context.document.getSelection();
      const linkText = displayText ?? url;
      const html = `<a href="${url}">${linkText}</a>`;
      selection.insertHtml(html, Word.InsertLocation.replace);
      await context.sync();
      return `Hyperlink inserted: "${linkText}" → ${url}`;
    },
  },

  {
    name: 'insert_footnote',
    description:
      'Insert a footnote at the current selection. Requires WordApi 1.5; returns an error message if not supported.',
    params: {
      text: { type: 'string', description: 'Footnote text.' },
    },
    execute: async (context, args) => {
      const { text } = args as { text: string };
      const selection = context.document.getSelection();
      const footnote = selection.insertFootnote(text);
      footnote.load('id');
      await context.sync();
      return `Footnote inserted: "${text}"`;
    },
  },

  {
    name: 'insert_endnote',
    description:
      'Insert an endnote at the current selection. Requires WordApi 1.5; returns an error message if not supported.',
    params: {
      text: { type: 'string', description: 'Endnote text.' },
    },
    execute: async (context, args) => {
      const { text } = args as { text: string };
      const selection = context.document.getSelection();
      const endnote = selection.insertEndnote(text);
      endnote.load('id');
      await context.sync();
      return `Endnote inserted: "${text}"`;
    },
  },

  {
    name: 'get_footnotes_endnotes',
    description:
      'Get all footnotes and endnotes from the document body. Requires WordApi 1.5; returns an error message if not supported.',
    params: {},
    execute: async context => {
      const body = context.document.body;
      const footnotes = body.footnotes;
      const endnotes = body.endnotes;
      footnotes.load('items');
      endnotes.load('items');
      await context.sync();
      const lines: string[] = ['Footnotes & Endnotes', '='.repeat(40)];
      if (footnotes.items.length > 0) {
        lines.push(`\nFootnotes (${String(footnotes.items.length)}):`);
        for (const fn of footnotes.items) fn.body.load('text');
        await context.sync();
        for (let i = 0; i < footnotes.items.length; i++) {
          lines.push(`  ${String(i + 1)}. ${footnotes.items[i].body.text.trim()}`);
        }
      } else {
        lines.push('\nFootnotes: (none)');
      }
      if (endnotes.items.length > 0) {
        lines.push(`\nEndnotes (${String(endnotes.items.length)}):`);
        for (const en of endnotes.items) en.body.load('text');
        await context.sync();
        for (let i = 0; i < endnotes.items.length; i++) {
          lines.push(`  ${String(i + 1)}. ${endnotes.items[i].body.text.trim()}`);
        }
      } else {
        lines.push('\nEndnotes: (none)');
      }
      return lines.join('\n');
    },
  },

  {
    name: 'delete_content',
    description: 'Delete the currently selected content in the Word document.',
    params: {},
    execute: async context => {
      const selection = context.document.getSelection();
      selection.delete();
      await context.sync();
      return 'Selected content deleted.';
    },
  },

  {
    name: 'insert_content_control',
    description:
      'Wrap the current selection in a content control. Optionally set a title, tag, and type.',
    params: {
      title: {
        type: 'string',
        required: false,
        description: 'Optional title for the content control.',
      },
      tag: {
        type: 'string',
        required: false,
        description: 'Optional tag for the content control.',
      },
      type: {
        type: 'string',
        required: false,
        enum: ['RichText', 'PlainText'],
        description: 'Content control type. Default: RichText.',
      },
    },
    execute: async (context, args) => {
      const { title, tag, type = 'RichText' } = args as {
        title?: string;
        tag?: string;
        type?: string;
      };
      const selection = context.document.getSelection();
      const ccType =
        type === 'PlainText'
          ? Word.ContentControlType.plainText
          : Word.ContentControlType.richText;
      const cc = selection.insertContentControl(ccType);
      if (title !== undefined) cc.title = title;
      if (tag !== undefined) cc.tag = tag;
      await context.sync();
      return `Content control inserted (type=${type}${title ? `, title="${title}"` : ''}${tag ? `, tag="${tag}"` : ''}).`;
    },
  },

  {
    name: 'format_found_text',
    description:
      'Search for text in the document and apply font formatting to all matches.',
    params: {
      searchText: { type: 'string', description: 'Text to search for.' },
      bold: { type: 'boolean', required: false, description: 'Make matched text bold.' },
      italic: { type: 'boolean', required: false, description: 'Make matched text italic.' },
      fontSize: { type: 'number', required: false, description: 'Font size in points.' },
      fontColor: {
        type: 'string',
        required: false,
        description: 'Font color as hex (e.g. "#FF0000").',
      },
      fontName: {
        type: 'string',
        required: false,
        description: 'Font family name (e.g. "Calibri").',
      },
      underline: { type: 'boolean', required: false, description: 'Underline matched text.' },
      highlightColor: {
        type: 'string',
        required: false,
        description:
          'Highlight color. Allowed values: Yellow, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, Black, White, None.',
      },
    },
    execute: async (context, args) => {
      const { searchText, bold, italic, fontSize, fontColor, fontName, underline, highlightColor } =
        args as {
          searchText: string;
          bold?: boolean;
          italic?: boolean;
          fontSize?: number;
          fontColor?: string;
          fontName?: string;
          underline?: boolean;
          highlightColor?: string;
        };
      const body = context.document.body;
      const results = body.search(searchText, { matchCase: false, matchWholeWord: false });
      results.load('items');
      await context.sync();
      if (results.items.length === 0) return `No matches found for "${searchText}".`;
      for (const result of results.items) {
        const font = result.font;
        if (bold !== undefined) font.bold = bold;
        if (italic !== undefined) font.italic = italic;
        if (fontSize !== undefined) font.size = fontSize;
        if (fontColor !== undefined) font.color = fontColor;
        if (fontName !== undefined) font.name = fontName;
        if (underline !== undefined)
          font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        if (highlightColor !== undefined) font.highlightColor = highlightColor;
      }
      await context.sync();
      const applied: string[] = [];
      if (bold !== undefined) applied.push(`bold=${String(bold)}`);
      if (italic !== undefined) applied.push(`italic=${String(italic)}`);
      if (fontSize !== undefined) applied.push(`fontSize=${String(fontSize)}`);
      if (fontColor !== undefined) applied.push(`color="${fontColor}"`);
      if (fontName !== undefined) applied.push(`fontName="${fontName}"`);
      if (underline !== undefined) applied.push(`underline=${String(underline)}`);
      if (highlightColor !== undefined) applied.push(`highlight="${highlightColor}"`);
      return `Formatted ${String(results.items.length)} match(es) of "${searchText}": ${applied.join(', ')}.`;
    },
  },

  {
    name: 'get_sections',
    description: 'List all sections in the Word document with header and footer summaries.',
    params: {},
    execute: async context => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      const lines: string[] = [
        'Document Sections',
        '='.repeat(40),
        `Total sections: ${String(sections.items.length)}`,
      ];
      for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];
        const header = section.getHeader('Primary');
        const footer = section.getFooter('Primary');
        header.load('text');
        footer.load('text');
        await context.sync();
        lines.push(
          `\nSection ${String(i + 1)}:`,
          `  Header: ${header.text.trim() ? header.text.trim().slice(0, 80) : '(empty)'}`,
          `  Footer: ${footer.text.trim() ? footer.text.trim().slice(0, 80) : '(empty)'}`
        );
      }
      return lines.join('\n');
    },
  },
];

export const wordTools = createWordTools(wordConfigs);
