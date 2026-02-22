/**
 * Integration test: Word tool configs and factory output.
 *
 * Validates that:
 * - wordConfigs are well-formed (names, descriptions, execute fns)
 * - createWordTools() produces Tool objects with correct JSON Schema parameters
 * - each tool's schema accepts valid inputs and rejects invalid ones
 *
 * NOTE: Testing the execute() logic requires a real Word.RequestContext and
 * MUST be covered by E2E tests running in Word Desktop — not here.
 */
import { describe, it, expect } from 'vitest';
import Ajv from 'ajv';
import { wordConfigs, wordTools } from '@/tools/word';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return Boolean(ajv.compile(schema as object)(data));
}

const toolsByName = Object.fromEntries(wordTools.map(t => [t.name, t]));

const EXPECTED_TOOL_NAMES = [
  'get_document_overview',
  'get_document_content',
  'get_document_section',
  'set_document_content',
  'get_selection',
  'get_selection_text',
  'insert_content_at_selection',
  'find_and_replace',
  'insert_table',
  'apply_style_to_selection',
] as const;

// ─── Structural ───────────────────────────────────────────────────────────────

describe('Integration: Word tool configs — structural', () => {
  it('wordConfigs contains exactly the expected tools', () => {
    const actual = wordConfigs.map(c => c.name).sort();
    const expected = [...EXPECTED_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every config has a non-empty name, description, and execute function', () => {
    for (const config of wordConfigs) {
      expect(config.name.length).toBeGreaterThan(0);
      expect(config.description.length).toBeGreaterThan(0);
      expect(typeof config.execute).toBe('function');
    }
  });

  it('createWordTools() produces the same tool names as configs', () => {
    const configNames = wordConfigs.map(c => c.name).sort();
    const toolNames = wordTools.map(t => t.name).sort();
    expect(toolNames).toEqual(configNames);
  });

  it('every generated tool has a parameters schema and a handler', () => {
    for (const tool of wordTools) {
      expect(tool.parameters).toBeDefined();
      expect(typeof tool.handler).toBe('function');
    }
  });
});

// ─── No-param tools ───────────────────────────────────────────────────────────

describe('Integration: Word schema — no-param tools', () => {
  it('get_document_overview accepts empty args', () => {
    expect(validate(toolsByName.get_document_overview.parameters, {})).toBe(true);
  });

  it('get_document_content accepts empty args', () => {
    expect(validate(toolsByName.get_document_content.parameters, {})).toBe(true);
  });

  it('get_selection accepts empty args', () => {
    expect(validate(toolsByName.get_selection.parameters, {})).toBe(true);
  });

  it('get_selection_text accepts empty args', () => {
    expect(validate(toolsByName.get_selection_text.parameters, {})).toBe(true);
  });
});

// ─── get_document_section ─────────────────────────────────────────────────────

describe('Integration: Word schema — get_document_section', () => {
  it('requires headingText', () => {
    const schema = toolsByName.get_document_section.parameters;
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { headingText: 'Introduction' })).toBe(true);
  });

  it('rejects non-string headingText', () => {
    expect(validate(toolsByName.get_document_section.parameters, { headingText: 42 })).toBe(false);
  });
});

// ─── set_document_content ─────────────────────────────────────────────────────

describe('Integration: Word schema — set_document_content', () => {
  it('requires html', () => {
    const schema = toolsByName.set_document_content.parameters;
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { html: '<p>Hello</p>' })).toBe(true);
  });
});

// ─── insert_content_at_selection ─────────────────────────────────────────────

describe('Integration: Word schema — insert_content_at_selection', () => {
  it('requires html', () => {
    const schema = toolsByName.insert_content_at_selection.parameters;
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { html: '<p>Inserted</p>' })).toBe(true);
  });

  it('accepts optional location enum values', () => {
    const schema = toolsByName.insert_content_at_selection.parameters;
    for (const loc of ['Replace', 'Before', 'After', 'Start', 'End']) {
      expect(validate(schema, { html: '<p>x</p>', location: loc })).toBe(true);
    }
  });

  it('rejects invalid location value', () => {
    expect(
      validate(toolsByName.insert_content_at_selection.parameters, {
        html: '<p>x</p>',
        location: 'Invalid',
      })
    ).toBe(false);
  });
});

// ─── find_and_replace ─────────────────────────────────────────────────────────

describe('Integration: Word schema — find_and_replace', () => {
  it('requires find and replace', () => {
    const schema = toolsByName.find_and_replace.parameters;
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { find: 'foo' })).toBe(false);
    expect(validate(schema, { find: 'foo', replace: 'bar' })).toBe(true);
  });
});

// ─── insert_table ─────────────────────────────────────────────────────────────

describe('Integration: Word schema — insert_table', () => {
  it('requires rows and columns', () => {
    const schema = toolsByName.insert_table.parameters;
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { rows: 3 })).toBe(false);
    expect(validate(schema, { rows: 3, columns: 4 })).toBe(true);
  });

  it('rejects non-integer rows/columns', () => {
    const schema = toolsByName.insert_table.parameters;
    expect(validate(schema, { rows: 'three', columns: 4 })).toBe(false);
  });
});

// ─── apply_style_to_selection ─────────────────────────────────────────────────

describe('Integration: Word schema — apply_style_to_selection', () => {
  it('accepts empty args (all optional)', () => {
    expect(validate(toolsByName.apply_style_to_selection.parameters, {})).toBe(true);
  });

  it('accepts style name', () => {
    expect(
      validate(toolsByName.apply_style_to_selection.parameters, { styleName: 'Heading 1' })
    ).toBe(true);
  });
});
