import type { Tool, ToolResultObject } from '@github/copilot-sdk';

const getDocumentOverview: Tool = {
  name: 'get_document_overview',
  description:
    'Get a structural overview of the Word document including heading hierarchy, paragraph count, tables, and content controls. Use this first to understand the document structure.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
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

        const summary = [
          `Document Overview`,
          `${'='.repeat(40)}`,
          `Paragraphs: ${String(paragraphCount)}`,
          `Tables: ${String(tableCount)}`,
          '',
          headings.length > 0 ? `Headings:\n${headings.join('\n')}` : '(no headings found)',
        ].join('\n');

        return summary;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getDocumentContent: Tool = {
  name: 'get_document_content',
  description:
    'Get the full HTML content of the Word document body. Returns rich formatted content that preserves structure such as headings, lists, and tables.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        const htmlResult = body.getHtml();
        await context.sync();
        return htmlResult.value;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getDocumentSection: Tool = {
  name: 'get_document_section',
  description:
    'Get the HTML content of a specific section identified by a heading. Finds the heading by partial text match and returns the content until the next heading of the same or higher level.',
  parameters: {
    type: 'object',
    properties: {
      headingText: {
        type: 'string',
        description: 'Partial or full text of the heading that starts the section.',
      },
    },
    required: ['headingText'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { headingText } = (args ?? {}) as { headingText: string };
    try {
      return await Word.run(async context => {
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
            p.style?.startsWith('Heading') &&
            p.text.toLowerCase().includes(headingText.toLowerCase())
        );

        if (!headingPara) {
          return `No heading found containing "${headingText}".`;
        }

        const headingLevel = parseInt(headingPara.style.replace('Heading ', ''), 10) || 1;
        const range = headingPara.getRange();

        // Expand range to include content until next heading of same or higher level
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const setDocumentContent: Tool = {
  name: 'set_document_content',
  description:
    'Replace the entire document body with new HTML content. WARNING: This clears all existing content.',
  parameters: {
    type: 'object',
    properties: {
      html: {
        type: 'string',
        description: 'HTML content to set as the full document body.',
      },
    },
    required: ['html'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { html } = (args ?? {}) as { html: string };
    try {
      return await Word.run(async context => {
        const body = context.document.body;
        body.clear();
        body.insertHtml(html, Word.InsertLocation.start);
        await context.sync();
        return 'Document content replaced successfully.';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelection: Tool = {
  name: 'get_selection',
  description:
    'Get the currently selected content in the Word document as OOXML. Useful for inspecting the structure of the selection.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        const ooxmlResult = selection.getOoxml();
        await context.sync();

        const ooxml = ooxmlResult.value;
        // Extract the <w:document> element for a cleaner view
        const docMatch = /<w:document[^>]*>[\s\S]*<\/w:document>/i.exec(ooxml);
        return docMatch ? docMatch[0] : ooxml;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const getSelectionText: Tool = {
  name: 'get_selection_text',
  description: 'Get the plain text of the currently selected content in the Word document.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        return selection.text.length > 0 ? selection.text : '(no text selected)';
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertContentAtSelection: Tool = {
  name: 'insert_content_at_selection',
  description: 'Insert HTML content at or relative to the current selection in the Word document.',
  parameters: {
    type: 'object',
    properties: {
      html: { type: 'string', description: 'HTML content to insert.' },
      location: {
        type: 'string',
        enum: ['Replace', 'Before', 'After', 'Start', 'End'],
        description:
          'Where to insert relative to the selection. Replace overwrites the selection. Default: Replace.',
      },
    },
    required: ['html'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const { html, location = 'Replace' } = (args ?? {}) as {
      html: string;
      location?: string;
    };

    const locationMap: Record<string, Word.InsertLocation> = {
      Replace: Word.InsertLocation.replace,
      Before: Word.InsertLocation.before,
      After: Word.InsertLocation.after,
      Start: Word.InsertLocation.start,
      End: Word.InsertLocation.end,
    };

    const insertLocation = locationMap[location] ?? Word.InsertLocation.replace;

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.insertHtml(html, insertLocation);
        await context.sync();
        return `Content inserted at selection (location: ${location}).`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const findAndReplace: Tool = {
  name: 'find_and_replace',
  description: 'Find text in the Word document and replace all occurrences with new text.',
  parameters: {
    type: 'object',
    properties: {
      find: { type: 'string', description: 'Text to search for.' },
      replace: { type: 'string', description: 'Replacement text.' },
      matchCase: {
        type: 'boolean',
        description: 'Whether the search is case-sensitive. Default: false.',
      },
      matchWholeWord: {
        type: 'boolean',
        description: 'Whether to match whole words only. Default: false.',
      },
    },
    required: ['find', 'replace'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      find,
      replace,
      matchCase = false,
      matchWholeWord = false,
    } = (args ?? {}) as {
      find: string;
      replace: string;
      matchCase?: boolean;
      matchWholeWord?: boolean;
    };
    try {
      return await Word.run(async context => {
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const insertTable: Tool = {
  name: 'insert_table',
  description:
    'Insert a table at the current selection in the Word document. Supports grid, striped, and plain styles.',
  parameters: {
    type: 'object',
    properties: {
      rows: { type: 'number', description: 'Number of rows.' },
      columns: { type: 'number', description: 'Number of columns.' },
      data: {
        type: 'array',
        items: { type: 'array', items: { type: 'string' } },
        description: 'Optional 2D array of cell values (row-major). Omit for an empty table.',
      },
      style: {
        type: 'string',
        enum: ['grid', 'striped', 'plain'],
        description:
          'Visual style. grid adds borders, striped alternates row colors, plain has no formatting. Default: grid.',
      },
      hasHeaderRow: {
        type: 'boolean',
        description: 'Whether the first row is a header row with distinct styling. Default: true.',
      },
    },
    required: ['rows', 'columns'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      rows,
      columns,
      data,
      style = 'grid',
      hasHeaderRow = true,
    } = (args ?? {}) as {
      rows: number;
      columns: number;
      data?: string[][];
      style?: 'grid' | 'striped' | 'plain';
      hasHeaderRow?: boolean;
    };

    try {
      return await Word.run(async context => {
        const selection = context.document.getSelection();

        // Build flat cell data array expected by insertTable
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

        table.load('rows');
        await context.sync();

        if (style === 'grid') {
          table.style = 'Table Grid';
        }

        if (hasHeaderRow && rows > 0) {
          const headerRow = table.rows.items[0];
          headerRow.load('cells');
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
            row.load('cells');
            await context.sync();
            for (const cell of row.cells.items) {
              cell.shadingColor = '#E8E8E8';
            }
          }
        }

        await context.sync();
        return `Inserted a ${String(rows)}×${String(columns)} table with "${style}" style.`;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

const applyStyleToSelection: Tool = {
  name: 'apply_style_to_selection',
  description: 'Apply font formatting to the currently selected text in the Word document.',
  parameters: {
    type: 'object',
    properties: {
      bold: { type: 'boolean', description: 'Make text bold.' },
      italic: { type: 'boolean', description: 'Make text italic.' },
      underline: { type: 'boolean', description: 'Underline text. true = single underline.' },
      strikeThrough: { type: 'boolean', description: 'Apply strikethrough.' },
      fontSize: { type: 'number', description: 'Font size in points.' },
      fontName: { type: 'string', description: 'Font family name (e.g. "Calibri", "Arial").' },
      fontColor: {
        type: 'string',
        description: 'Font color as hex (e.g. "#FF0000") or named color.',
      },
      highlightColor: {
        type: 'string',
        description:
          'Highlight color. Allowed values: Yellow, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, Black, White, None.',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    const {
      bold,
      italic,
      underline,
      strikeThrough,
      fontSize,
      fontName,
      fontColor,
      highlightColor,
    } = (args ?? {}) as {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strikeThrough?: boolean;
      fontSize?: number;
      fontName?: string;
      fontColor?: string;
      highlightColor?: string;
    };

    try {
      return await Word.run(async context => {
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
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
    }
  },
};

export const wordTools: Tool[] = [
  getDocumentOverview,
  getDocumentContent,
  getDocumentSection,
  setDocumentContent,
  getSelection,
  getSelectionText,
  insertContentAtSelection,
  findAndReplace,
  insertTable,
  applyStyleToSelection,
];
