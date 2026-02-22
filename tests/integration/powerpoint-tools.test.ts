/**
 * Integration test: PowerPoint tool configs and factory output.
 *
 * Validates that:
 * - powerPointConfigs are well-formed (names, descriptions, execute fns)
 * - createPptTools() produces Tool objects with correct JSON Schema parameters
 * - get_presentation_overview exposes the thumbnailWidth param with the right shape
 *
 * NOTE: Testing the execute() logic (getImageAsBase64, slide text extraction, etc.)
 * requires a real PowerPoint.RequestContext and MUST be covered by E2E tests running
 * in PowerPoint Desktop — not here.
 */
import { describe, it, expect } from 'vitest';
import Ajv from 'ajv';
import { powerPointConfigs, powerPointTools } from '@/tools/powerpoint';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return Boolean(ajv.compile(schema as object)(data));
}

const toolsByName = Object.fromEntries(powerPointTools.map(t => [t.name, t]));
const configsByName = Object.fromEntries(powerPointConfigs.map(c => [c.name, c]));

const EXPECTED_TOOL_NAMES = [
  'get_presentation_overview',
  'get_presentation_content',
  'get_slide_image',
  'get_slide_notes',
  'set_presentation_content',
  'add_slide_from_code',
  'clear_slide',
  'update_slide_shape',
  'set_slide_notes',
  'duplicate_slide',
] as const;

// ─── Structural ───────────────────────────────────────────────────────────────

describe('Integration: PowerPoint tool configs — structural', () => {
  it('powerPointConfigs contains exactly the expected tools', () => {
    const actual = powerPointConfigs.map(c => c.name).sort();
    const expected = [...EXPECTED_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every config has a non-empty name, description, and execute function', () => {
    for (const config of powerPointConfigs) {
      expect(config.name.length).toBeGreaterThan(0);
      expect(config.description.length).toBeGreaterThan(0);
      expect(typeof config.execute).toBe('function');
    }
  });

  it('createPptTools() produces the same tool names as configs', () => {
    const configNames = powerPointConfigs.map(c => c.name).sort();
    const toolNames = powerPointTools.map(t => t.name).sort();
    expect(toolNames).toEqual(configNames);
  });

  it('every generated tool has a parameters schema and a handler', () => {
    for (const tool of powerPointTools) {
      expect(tool.parameters).toBeDefined();
      expect(typeof tool.handler).toBe('function');
    }
  });
});

// ─── get_presentation_overview ────────────────────────────────────────────────

describe('Integration: get_presentation_overview schema', () => {
  const schema = toolsByName.get_presentation_overview.parameters;

  it('accepts empty args (all params optional)', () => {
    expect(validate(schema, {})).toBe(true);
  });

  it('accepts thumbnailWidth as a number', () => {
    expect(validate(schema, { thumbnailWidth: 600 })).toBe(true);
    expect(validate(schema, { thumbnailWidth: 1200 })).toBe(true);
  });

  it('rejects thumbnailWidth as a string', () => {
    expect(validate(schema, { thumbnailWidth: 'wide' })).toBe(false);
  });

  it('config exposes thumbnailWidth param as optional number', () => {
    const paramDef = configsByName.get_presentation_overview.params.thumbnailWidth;
    expect(paramDef).toBeDefined();
    expect(paramDef.type).toBe('number');
    expect(paramDef.required).toBeFalsy();
  });
});

// ─── Slide-index tools ────────────────────────────────────────────────────────

describe('Integration: slide-index tool schemas', () => {
  it('get_slide_image requires slideIndex (number)', () => {
    const schema = toolsByName.get_slide_image.parameters;
    expect(validate(schema, { slideIndex: 0 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('get_slide_image accepts optional width', () => {
    const schema = toolsByName.get_slide_image.parameters;
    expect(validate(schema, { slideIndex: 0, width: 1024 })).toBe(true);
  });

  it('clear_slide requires slideIndex', () => {
    const schema = toolsByName.clear_slide.parameters;
    expect(validate(schema, { slideIndex: 2 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('update_slide_shape requires slideIndex, shapeIndex, and text', () => {
    const schema = toolsByName.update_slide_shape.parameters;
    expect(validate(schema, { slideIndex: 0, shapeIndex: 1, text: 'Hello' })).toBe(true);
    expect(validate(schema, { slideIndex: 0 })).toBe(false);
  });

  it('duplicate_slide requires sourceIndex', () => {
    const schema = toolsByName.duplicate_slide.parameters;
    expect(validate(schema, { sourceIndex: 0 })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });
});

// ─── Content tools ────────────────────────────────────────────────────────────

describe('Integration: content tool schemas', () => {
  it('set_presentation_content requires slideIndex and text', () => {
    const schema = toolsByName.set_presentation_content.parameters;
    expect(validate(schema, { slideIndex: 0, text: 'Hello' })).toBe(true);
    expect(validate(schema, { slideIndex: 0 })).toBe(false);
  });

  it('add_slide_from_code requires code string', () => {
    const schema = toolsByName.add_slide_from_code.parameters;
    expect(validate(schema, { code: 'slide.addText("Hi", {x:1,y:1,w:8,h:1})' })).toBe(true);
    expect(validate(schema, {})).toBe(false);
  });

  it('add_slide_from_code accepts optional replaceSlideIndex', () => {
    const schema = toolsByName.add_slide_from_code.parameters;
    expect(validate(schema, { code: 'slide.addText("X")', replaceSlideIndex: 0 })).toBe(true);
  });

  it('get_presentation_content accepts no args (read all slides)', () => {
    const schema = toolsByName.get_presentation_content.parameters;
    expect(validate(schema, {})).toBe(true);
  });

  it('get_presentation_content accepts slideIndex or startIndex/endIndex', () => {
    const schema = toolsByName.get_presentation_content.parameters;
    expect(validate(schema, { slideIndex: 2 })).toBe(true);
    expect(validate(schema, { startIndex: 0, endIndex: 3 })).toBe(true);
  });
});
