import type { WordToolConfig } from '../codegen';

function requireWordApi(version: string): void {
  if (!Office.context.requirements.isSetSupported('WordApi', version)) {
    throw new Error(
      `This operation requires WordApi ${version}, which this Word version does not support.`
    );
  }
}

function requireIndex(value: unknown, name: string, count: number): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value >= count) {
    throw new Error(
      `${name} must be an integer between 0 and ${count - 1}. Read the current document again.`
    );
  }
  return value;
}

function requireString(value: unknown, name: string, allowEmpty = false): string {
  if (typeof value !== 'string' || (!allowEmpty && !value.trim())) {
    throw new Error(`${name} must be ${allowEmpty ? 'a string' : 'a non-empty string'}.`);
  }
  return value;
}

async function headingSection(
  context: Word.RequestContext,
  headingText: string,
  partial = false,
  includeHeading = false
) {
  requireWordApi('1.3');
  requireString(headingText, 'headingText');
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text,items/styleBuiltIn');
  await context.sync();
  const level = (p: Word.Paragraph): number => {
    const match = /^Heading([1-9])$/.exec(p.styleBuiltIn);
    return match ? Number(match[1]) : 0;
  };
  const wanted = headingText.trim().toLowerCase();
  const matches = paragraphs.items.filter(
    p =>
      level(p) > 0 &&
      (partial
        ? p.text.trim().toLowerCase().includes(wanted)
        : p.text.trim().toLowerCase() === wanted)
  );
  if (matches.length !== 1) {
    throw new Error(
      matches.length === 0
        ? `No heading found matching "${headingText}".`
        : `Heading "${headingText}" is ambiguous (${matches.length} matches). Use a unique heading or a physical sectionIndex.`
    );
  }
  const heading = matches[0];
  const following = paragraphs.items.slice(paragraphs.items.indexOf(heading) + 1);
  const next = following.find(p => level(p) > 0 && level(p) <= level(heading));
  const start = heading.getRange(
    includeHeading ? Word.RangeLocation.start : Word.RangeLocation.after
  );
  const end = next
    ? next.getRange(Word.RangeLocation.start)
    : context.document.body.getRange(Word.RangeLocation.end);
  const range =
    !includeHeading && following.length === 0
      ? heading.getRange(Word.RangeLocation.end)
      : start.expandTo(end);
  return { range, heading, next, body: undefined };
}

export async function headingSectionRange(
  context: Word.RequestContext,
  headingText: string,
  partial = false,
  includeHeading = false
): Promise<Word.Range> {
  return (await headingSection(context, headingText, partial, includeHeading)).range;
}

async function sectionRange(context: Word.RequestContext, args: Record<string, unknown>) {
  if ((args.headingText !== undefined) === (args.sectionIndex !== undefined)) {
    throw new Error('Provide exactly one of headingText or sectionIndex.');
  }
  if (args.headingText !== undefined) {
    return headingSection(context, requireString(args.headingText, 'headingText'));
  }
  requireWordApi('1.3');
  const sections = context.document.sections;
  sections.load('items');
  await context.sync();
  const index = requireIndex(args.sectionIndex, 'sectionIndex', sections.items.length);
  // Content excludes the section's terminating break, preserving its layout and neighbours.
  const body = sections.items[index].body;
  return {
    range: body.getRange(Word.RangeLocation.content),
    heading: undefined,
    next: undefined,
    body,
  };
}

const targetParams: WordToolConfig['params'] = {
  headingText: {
    type: 'string',
    required: false,
    description:
      'Exact, case-insensitive, unique built-in heading text. Content excludes this heading and ends at the next same/higher-level heading. Supply this OR sectionIndex.',
  },
  sectionIndex: {
    type: 'number',
    required: false,
    description:
      'Zero-based physical section index from get_sections (not a heading section). Supply this OR headingText. Headers, footers, and terminating section break are excluded.',
  },
};

async function checkSectionText(
  context: Word.RequestContext,
  range: Word.Range,
  expectedText: unknown
): Promise<void> {
  const expected = requireString(expectedText, 'expectedText', true);
  range.load('text');
  await context.sync();
  if (range.text !== expected) {
    throw new Error(
      'Section content has changed or the target is incorrect. Read it again before editing.'
    );
  }
}

function insertSectionHtml(
  context: Word.RequestContext,
  target: Awaited<ReturnType<typeof sectionRange>>,
  html: string,
  location: 'Start' | 'End'
): void {
  let anchor: Word.Paragraph;
  if (target.body) anchor = target.body.insertParagraph('', location);
  else if (target.heading) {
    if (location === 'Start')
      anchor = target.heading.insertParagraph('', Word.InsertLocation.after);
    else if (target.next) anchor = target.next.insertParagraph('', Word.InsertLocation.before);
    else anchor = context.document.body.insertParagraph('', Word.InsertLocation.end);
  } else throw new Error('The section target could not be resolved.');
  // An independent paragraph avoids merging HTML into a heading or terminal table cell.
  anchor.styleBuiltIn = Word.BuiltInStyleName.normal;
  anchor.insertHtml(html, Word.InsertLocation.replace);
}

async function trackedSnapshot(context: Word.RequestContext) {
  requireWordApi('1.6');
  const collection = context.document.body.getTrackedChanges();
  collection.load('items/author,items/date,items/text,items/type');
  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const changes = collection.items.map((change, index) => ({
    index,
    author: change.author,
    date: change.date.toISOString(),
    text: change.text,
    type: change.type,
  }));
  // Word regenerates revision-session and export paragraph IDs even without edits.
  const documentXml = /<w:document\b[\s\S]*?<\/w:document>/.exec(ooxml.value)?.[0];
  if (!documentXml)
    throw new Error('Word did not return document XML for tracked-change verification.');
  const stableOoxml = documentXml
    .replace(/\s+w:rsid[A-Za-z]*="[^"]*"/g, '')
    .replace(/\s+w14:(?:paraId|textId)="[^"]*"/g, '')
    .replace(/<w:rsids>[\s\S]*?<\/w:rsids>/g, '');
  const bytes = new TextEncoder().encode(JSON.stringify({ ooxml: stableOoxml, changes }));
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  const snapshot = Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join(
    ''
  );
  return { collection, changes, snapshot };
}

export const targetedWordConfigs: readonly WordToolConfig[] = [
  {
    name: 'get_section_content',
    description:
      'Read HTML and exact text of a heading or physical section without changing the selection. Use the returned text as expectedText for a targeted edit.',
    params: targetParams,
    execute: async (context, args) => {
      const { range } = await sectionRange(context, args);
      range.load('text,isEmpty');
      await context.sync();
      if (range.isEmpty) return { text: range.text, html: '' };
      const html = range.getHtml();
      await context.sync();
      return { text: range.text, html: html.value };
    },
  },
  {
    name: 'insert_content_in_section',
    description:
      'Insert HTML at the start/end of an explicitly targeted section, preserving its heading, neighbouring content and selection outside the edited range. Read get_section_content first.',
    params: {
      ...targetParams,
      html: { type: 'string', description: 'Non-empty HTML to insert.' },
      location: {
        type: 'string',
        enum: ['Start', 'End'],
        description: 'Insert at the start or end of the section content.',
      },
      expectedText: {
        type: 'string',
        description:
          'Exact text returned by get_section_content. Editing is refused if it no longer matches.',
      },
    },
    execute: async (context, args) => {
      const html = requireString(args.html, 'html');
      if (args.location !== 'Start' && args.location !== 'End') {
        throw new Error('location must be Start or End.');
      }
      const target = await sectionRange(context, args);
      await checkSectionText(context, target.range, args.expectedText);
      insertSectionHtml(context, target, html, args.location);
      await context.sync();
      return 'Content inserted in the targeted section.';
    },
  },
  {
    name: 'replace_section_content',
    description:
      'Replace only the content of an explicitly targeted section with HTML. Preserves the starting heading, following sections, headers/footers and terminating section break. Read get_section_content first.',
    params: {
      ...targetParams,
      html: {
        type: 'string',
        description: 'Replacement HTML; an empty string clears the section content.',
      },
      expectedText: {
        type: 'string',
        description:
          'Exact text returned by get_section_content. Editing is refused if it no longer matches.',
      },
    },
    execute: async (context, args) => {
      const html = requireString(args.html, 'html', true);
      const target = await sectionRange(context, args);
      await checkSectionText(context, target.range, args.expectedText);
      target.range.clear();
      if (html) insertSectionHtml(context, target, html, 'Start');
      await context.sync();
      return 'Targeted section content replaced.';
    },
  },
  {
    name: 'get_tracked_changes',
    description:
      'Inspect tracked changes in the main document body (not headers/footers or other stories). Returns explicit zero-based indices, author/date/text/type and a snapshot required for accept/reject. Requires WordApi 1.6.',
    params: {},
    execute: async context => {
      const { changes, snapshot } = await trackedSnapshot(context);
      return { scope: 'document body', changes, snapshot };
    },
  },
  {
    name: 'manage_tracked_changes',
    description:
      'Accept or reject only explicitly listed tracked-change indices from get_tracked_changes. No implicit all/selection target. Refuses a changed snapshot; reread after each operation. Requires WordApi 1.6.',
    params: {
      action: {
        type: 'string',
        enum: ['Accept', 'Reject'],
        description: 'Deliberate action for the listed changes.',
      },
      changeIndices: {
        type: 'number[]',
        description: 'Non-empty list of unique zero-based indices from get_tracked_changes.',
      },
      snapshot: {
        type: 'string',
        description: 'Exact snapshot from the most recent get_tracked_changes result.',
      },
    },
    execute: async (context, args) => {
      if (args.action !== 'Accept' && args.action !== 'Reject') {
        throw new Error('action must be Accept or Reject.');
      }
      const snapshot = requireString(args.snapshot, 'snapshot');
      if (!Array.isArray(args.changeIndices) || args.changeIndices.length === 0) {
        throw new Error('changeIndices must be a non-empty list of unique integer indices.');
      }
      const current = await trackedSnapshot(context);
      if (current.snapshot !== snapshot) {
        throw new Error(
          'Document or tracked changes have changed. Call get_tracked_changes again before accepting or rejecting.'
        );
      }
      const indices = args.changeIndices.map(value =>
        requireIndex(value, 'changeIndex', current.changes.length)
      );
      if (new Set(indices).size !== indices.length) {
        throw new Error('changeIndices must not contain duplicates.');
      }
      for (const index of [...indices].sort((a, b) => b - a)) {
        const change = current.collection.items[index];
        if (args.action === 'Accept') change.accept();
        else change.reject();
      }
      await context.sync();
      return { action: args.action, changeIndices: indices, count: indices.length };
    },
  },
  {
    name: 'get_change_tracking_mode',
    description: 'Read document-level change tracking mode. Requires WordApi 1.4.',
    params: {},
    execute: async context => {
      requireWordApi('1.4');
      context.document.load('changeTrackingMode');
      await context.sync();
      return { mode: context.document.changeTrackingMode };
    },
  },
  {
    name: 'set_change_tracking_mode',
    description:
      'Explicitly set document-level change tracking to Off, TrackAll or TrackMineOnly. Does not accept/reject existing changes. Requires WordApi 1.4.',
    params: {
      mode: {
        type: 'string',
        enum: ['Off', 'TrackAll', 'TrackMineOnly'],
        description: 'The requested document change tracking mode.',
      },
    },
    execute: async (context, args) => {
      requireWordApi('1.4');
      if (args.mode !== 'Off' && args.mode !== 'TrackAll' && args.mode !== 'TrackMineOnly') {
        throw new Error('mode must be Off, TrackAll or TrackMineOnly.');
      }
      context.document.changeTrackingMode = args.mode;
      await context.sync();
      context.document.load('changeTrackingMode');
      await context.sync();
      if (context.document.changeTrackingMode !== args.mode) {
        throw new Error('Word did not apply the requested change tracking mode.');
      }
      return { mode: context.document.changeTrackingMode };
    },
  },
];
