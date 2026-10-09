import { wordConfigs } from '@/tools/word';

function assert(condition: boolean, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

async function tool<T>(name: string, args: Record<string, unknown> = {}): Promise<T> {
  const config = wordConfigs.find(c => c.name === name);
  if (!config) throw new Error(`Missing tool ${name}`);
  return Word.run(async context => (await config.execute(context, args)) as T);
}

interface SectionContent {
  text: string;
  html: string;
}
interface Change {
  index: number;
  author: string;
  date: string;
  text: string;
  type: string;
}
interface Changes {
  changes: Change[];
  snapshot: string;
}

async function expectError(
  name: string,
  args: Record<string, unknown>,
  message: string
): Promise<void> {
  try {
    await tool(name, args);
  } catch (error) {
    assert(String(error).includes(message), `Expected "${message}", got ${String(error)}`);
    return;
  }
  throw new Error(`Expected ${name} to refuse the operation`);
}

async function bodyText(): Promise<string> {
  return Word.run(async context => {
    context.document.body.load('text');
    await context.sync();
    return context.document.body.text;
  });
}

async function selectedText(): Promise<string> {
  return Word.run(async context => {
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    return selection.text;
  });
}

export async function testTargetedEditing(
  run: (name: string, test: () => Promise<void>) => Promise<void>
): Promise<void> {
  await Word.run(async context => {
    context.document.body.insertHtml(
      '<h1>Target</h1><p>Original content</p><h2>Nested</h2><p>Nested content</p>' +
        '<table><tr><td>Table sentinel</td></tr></table>' +
        '<h1>Neighbour</h1><p>Selection sentinel</p><h1>Empty last</h1><p></p>',
      Word.InsertLocation.replace
    );
    await context.sync();
    const seededParagraphs = context.document.body.paragraphs;
    seededParagraphs.load('items/text');
    await context.sync();
    const emptyHeading = seededParagraphs.items.find(p => p.text === 'Empty last');
    assert(emptyHeading !== undefined, 'Empty heading fixture missing');
    emptyHeading.styleBuiltIn = Word.BuiltInStyleName.heading1;
    const blank = context.document.body.insertParagraph('', Word.InsertLocation.end);
    blank.styleBuiltIn = Word.BuiltInStyleName.normal;
    await context.sync();
    const found = context.document.body.search('Selection sentinel');
    found.load('items');
    await context.sync();
    found.items[0].select();
    await context.sync();
  });

  await run('section:read-boundaries', async () => {
    const section = await tool<SectionContent>('get_section_content', { headingText: 'target' });
    assert(
      section.text.includes('Original content') &&
        section.text.includes('Nested content') &&
        section.text.includes('Table sentinel'),
      'Nested content/table missing'
    );
    assert(
      !section.text.includes('Target') && !section.text.includes('Neighbour'),
      'Section includes boundary heading'
    );
    const legacy = await tool<string>('get_document_section', { headingText: 'Target' });
    assert(
      legacy.includes('Target') && !legacy.includes('Neighbour'),
      'Legacy read boundaries incorrect'
    );
    assert((await selectedText()) === 'Selection sentinel', 'Reading moved the selection');
  });
  await run('section:insert-start-end', async () => {
    for (const location of ['Start', 'End']) {
      const section = await tool<SectionContent>('get_section_content', { headingText: 'Target' });
      await tool('insert_content_in_section', {
        headingText: 'Target',
        expectedText: section.text,
        html: `<p>Inserted ${location}</p>`,
        location,
      });
    }
    const section = await tool<SectionContent>('get_section_content', { headingText: 'Target' });
    assert(
      section.text.startsWith('Inserted Start') && section.text.trimEnd().endsWith('Inserted End'),
      'Insert escaped its requested boundary'
    );
    assert((await selectedText()) === 'Selection sentinel', 'Insertion moved unrelated selection');
  });
  await run('section:replace-preserves-neighbours', async () => {
    const section = await tool<SectionContent>('get_section_content', { headingText: 'Target' });
    await tool('replace_section_content', {
      headingText: 'Target',
      expectedText: section.text,
      html: '<p>Replacement only</p>',
    });
    const text = await bodyText();
    assert(
      text.includes('Target') &&
        text.includes('Neighbour') &&
        text.includes('Selection sentinel') &&
        text.includes('Empty last'),
      'Replacement removed neighbouring content/headings'
    );
    assert(
      text.includes('Replacement only') &&
        !text.includes('Nested content') &&
        !text.includes('Table sentinel'),
      'Replacement did not cover all section content'
    );
    assert(
      (await selectedText()) === 'Selection sentinel',
      'Replacement moved unrelated selection'
    );
    const neighbour = await tool<SectionContent>('get_section_content', {
      headingText: 'Neighbour',
    });
    assert(
      neighbour.text.includes('Selection sentinel'),
      'Replacement merged into the next heading'
    );
    await Word.run(async context => {
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      assert(tables.items.length === 0, 'Replacement left an empty table behind');
    });
  });
  await run('section:empty-last-section', async () => {
    const section = await tool<SectionContent>('get_section_content', {
      headingText: 'Empty last',
    });
    await tool('insert_content_in_section', {
      headingText: 'Empty last',
      expectedText: section.text,
      html: '<p>Last section text</p>',
      location: 'End',
    });
    const updated = await tool<SectionContent>('get_section_content', {
      headingText: 'Empty last',
    });
    assert(updated.text.includes('Last section text'), 'Empty final section insertion failed');
    await tool('insert_content_in_section', {
      headingText: 'Empty last',
      expectedText: updated.text,
      html: '<p>Second ending</p>',
      location: 'End',
    });
    const ending = await tool<SectionContent>('get_section_content', { headingText: 'Empty last' });
    assert(
      ending.text.indexOf('Second ending') > ending.text.indexOf('Last section text'),
      'End insertion was placed before existing section content'
    );
    await tool('replace_section_content', {
      headingText: 'Empty last',
      expectedText: ending.text,
      html: '',
    });
    assert((await bodyText()).includes('Empty last'), 'Clearing final section removed heading');
  });
  await run('section:invalid-targets-no-mutation', async () => {
    await Word.run(async context => {
      for (const text of ['Duplicate', 'one', 'Duplicate', 'two']) {
        const p = context.document.body.insertParagraph(text, Word.InsertLocation.end);
        p.styleBuiltIn =
          text === 'Duplicate' ? Word.BuiltInStyleName.heading1 : Word.BuiltInStyleName.normal;
      }
      await context.sync();
    });
    const before = await bodyText();
    const args = { html: '<p>wrong</p>', expectedText: '' };
    await expectError(
      'replace_section_content',
      { ...args, headingText: 'Duplicate' },
      'ambiguous'
    );
    await expectError('replace_section_content', { ...args, headingText: 'Missing' }, 'No heading');
    await expectError('replace_section_content', args, 'exactly one');
    await expectError(
      'replace_section_content',
      { ...args, headingText: 'Target', sectionIndex: 0 },
      'exactly one'
    );
    for (const sectionIndex of [-1, 0.5, 999]) {
      await expectError('replace_section_content', { ...args, sectionIndex }, 'integer');
    }
    assert((await bodyText()) === before, 'Invalid target mutated the document');
  });
  await run('section:stale-text-no-mutation', async () => {
    const before = await bodyText();
    await expectError(
      'replace_section_content',
      { headingText: 'Target', expectedText: 'stale', html: 'wrong' },
      'has changed'
    );
    await expectError(
      'insert_content_in_section',
      { headingText: 'Target', expectedText: 'stale', html: 'wrong', location: 'Start' },
      'has changed'
    );
    assert((await bodyText()) === before, 'Stale text mutated the document');
  });
  await run('section:heading-only-document', async () => {
    await Word.run(async context => {
      context.document.body.insertHtml('<p>Only heading</p>', Word.InsertLocation.replace);
      await context.sync();
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      assert(paragraphs.items.length === 1, 'Fixture must have exactly one paragraph');
      paragraphs.items[0].styleBuiltIn = Word.BuiltInStyleName.heading1;
      await context.sync();
    });
    const empty = await tool<SectionContent>('get_section_content', {
      headingText: 'Only heading',
    });
    assert(empty.text === '' && empty.html === '', 'Heading-only section should be empty');
    await tool('replace_section_content', {
      headingText: 'Only heading',
      expectedText: empty.text,
      html: '<p>Now populated</p>',
    });
    const populated = await tool<SectionContent>('get_section_content', {
      headingText: 'Only heading',
    });
    assert(populated.text.includes('Now populated'), 'Heading-only section could not be replaced');
  });
  await run('section:physical-boundaries', async () => {
    await Word.run(async context => {
      const body = context.document.body;
      body.insertHtml('<p>Physical first</p>', Word.InsertLocation.replace);
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
      body.insertParagraph('Physical neighbour', Word.InsertLocation.end);
      await context.sync();
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      assert(sections.items.length === 2, 'Fixture must have two physical sections');
      sections.items[0]
        .getHeader('Primary')
        .insertText('Header sentinel', Word.InsertLocation.replace);
      sections.items[0]
        .getFooter('Primary')
        .insertText('Footer sentinel', Word.InsertLocation.replace);
      await context.sync();
    });
    const section = await tool<SectionContent>('get_section_content', { sectionIndex: 0 });
    assert(
      section.text.includes('Physical first') && !section.text.includes('Physical neighbour'),
      'Physical read crossed section boundary'
    );
    await tool('replace_section_content', {
      sectionIndex: 0,
      expectedText: section.text,
      html: '<p>Physical replacement</p>',
    });
    const updated = await tool<SectionContent>('get_section_content', { sectionIndex: 0 });
    await tool('insert_content_in_section', {
      sectionIndex: 0,
      expectedText: updated.text,
      html: '<p>Physical insertion</p>',
      location: 'End',
    });
    await Word.run(async context => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      assert(sections.items.length === 2, 'Mutation removed section break');
      const header = sections.items[0].getHeader('Primary');
      const footer = sections.items[0].getFooter('Primary');
      header.load('text');
      footer.load('text');
      sections.items[1].body.load('text');
      await context.sync();
      assert(
        header.text.includes('Header sentinel') && footer.text.includes('Footer sentinel'),
        'Mutation changed header/footer'
      );
      assert(
        sections.items[1].body.text.includes('Physical neighbour'),
        'Mutation changed next physical section'
      );
    });
  });

  let originalMode: string | undefined;
  try {
    await run('tracking:mode', async () => {
      originalMode = (await tool<{ mode: string }>('get_change_tracking_mode')).mode;
      for (const mode of ['Off', 'TrackAll', 'TrackMineOnly']) {
        assert(
          (await tool<{ mode: string }>('set_change_tracking_mode', { mode })).mode === mode,
          `Could not set ${mode}`
        );
        assert(
          (await tool<{ mode: string }>('get_change_tracking_mode')).mode === mode,
          `Could not read ${mode}`
        );
      }
      await expectError('set_change_tracking_mode', { mode: 'On' }, 'mode must');
      await tool('set_change_tracking_mode', { mode: 'Off' });
      await Word.run(async context => {
        context.document.body.insertHtml(
          '<p>First base</p><p>Second base</p>',
          Word.InsertLocation.replace
        );
        await context.sync();
      });
      await tool('set_change_tracking_mode', { mode: 'TrackAll' });
      await Word.run(async context => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        paragraphs.items[0].insertText(' Added alpha', Word.InsertLocation.end);
        await context.sync();
        paragraphs.items[1].insertText(' Added beta', Word.InsertLocation.end);
        await context.sync();
      });
      await tool('set_change_tracking_mode', { mode: 'Off' });
    });
    await run('tracking:inspect', async () => {
      const first = await tool<Changes>('get_tracked_changes');
      const second = await tool<Changes>('get_tracked_changes');
      assert(first.changes.length >= 2, 'Expected independent tracked changes');
      assert(first.snapshot === second.snapshot, 'Snapshot is not stable across consecutive reads');
      assert(
        first.changes.some(c => c.text.includes('Added alpha')) &&
          first.changes.some(c => c.text.includes('Added beta')),
        'Change text missing'
      );
      assert(
        first.changes.every(c => c.author.length > 0 && c.date.length > 0 && c.type === 'Added'),
        'Change metadata missing'
      );
    });
    await run('tracking:invalid-targets-no-mutation', async () => {
      const before = await tool<Changes>('get_tracked_changes');
      for (const changeIndices of [[], [0, 0], [-1], [0.5], [999], [0, 999]]) {
        await expectError(
          'manage_tracked_changes',
          { action: 'Accept', snapshot: before.snapshot, changeIndices },
          changeIndices.length === 0
            ? 'non-empty'
            : changeIndices.length === 2 && changeIndices[1] === 0
              ? 'duplicates'
              : 'integer'
        );
      }
      await expectError(
        'manage_tracked_changes',
        { action: 'All', snapshot: before.snapshot, changeIndices: [0] },
        'action must'
      );
      assert(
        (await tool<Changes>('get_tracked_changes')).snapshot === before.snapshot,
        'Invalid indices mutated changes'
      );
    });
    await run('tracking:stale-snapshot-no-mutation', async () => {
      const before = await tool<Changes>('get_tracked_changes');
      await Word.run(async context => {
        context.document.body.insertParagraph('Untracked edit', Word.InsertLocation.end);
        await context.sync();
      });
      const current = await tool<Changes>('get_tracked_changes');
      await expectError(
        'manage_tracked_changes',
        { action: 'Reject', changeIndices: [0], snapshot: before.snapshot },
        'have changed'
      );
      assert(
        (await tool<Changes>('get_tracked_changes')).snapshot === current.snapshot,
        'Stale snapshot mutated changes'
      );
    });
    await run('tracking:accept-explicit-target', async () => {
      const before = await tool<Changes>('get_tracked_changes');
      const alpha = before.changes.find(c => c.text.includes('Added alpha'));
      assert(alpha !== undefined, 'Alpha target missing');
      await tool('manage_tracked_changes', {
        action: 'Accept',
        changeIndices: [alpha.index],
        snapshot: before.snapshot,
      });
      const after = await tool<Changes>('get_tracked_changes');
      assert(
        after.changes.length === before.changes.length - 1 &&
          after.changes.some(c => c.text.includes('Added beta')),
        'Accept affected untargeted changes'
      );
      assert((await bodyText()).includes('Added alpha'), 'Accept removed inserted text');
      await expectError(
        'manage_tracked_changes',
        { action: 'Reject', changeIndices: [0], snapshot: before.snapshot },
        'have changed'
      );
    });
    await run('tracking:reject-explicit-target', async () => {
      const before = await tool<Changes>('get_tracked_changes');
      const beta = before.changes.find(c => c.text.includes('Added beta'));
      assert(beta !== undefined, 'Beta target missing');
      await tool('manage_tracked_changes', {
        action: 'Reject',
        changeIndices: [beta.index],
        snapshot: before.snapshot,
      });
      await run('tracking:batch-deletions', async () => {
        await tool('set_change_tracking_mode', { mode: 'TrackAll' });
        await Word.run(async context => {
          for (const text of ['First base', 'Second base']) {
            const found = context.document.body.search(text, { matchCase: true });
            found.load('items');
            await context.sync();
            assert(found.items.length === 1, `Deletion fixture target missing: ${text}`);
            found.items[0].delete();
            await context.sync();
          }
        });
        await tool('set_change_tracking_mode', { mode: 'Off' });
        const deletions = await tool<Changes>('get_tracked_changes');
        const targets = deletions.changes.filter(c => c.type === 'Deleted');
        assert(targets.length === 2, 'Expected two independent tracked deletions');
        await tool('manage_tracked_changes', {
          action: 'Reject',
          changeIndices: targets.map(c => c.index),
          snapshot: deletions.snapshot,
        });
        const text = await bodyText();
        assert(
          text.includes('First base') &&
            text.includes('Second base') &&
            text.includes('Added alpha'),
          'Batch rejection did not restore only the deleted text'
        );
        assert(
          (await tool<Changes>('get_tracked_changes')).changes.length === 0,
          'Batch targets were not resolved'
        );
      });
      assert(
        (await tool<Changes>('get_tracked_changes')).changes.length === before.changes.length - 1,
        'Reject affected wrong number of changes'
      );
      assert(
        !(await bodyText()).includes('Added beta') && (await bodyText()).includes('Added alpha'),
        'Reject affected untargeted text'
      );
    });
  } finally {
    if (originalMode !== undefined) await tool('set_change_tracking_mode', { mode: originalMode });
  }
}
