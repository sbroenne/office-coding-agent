import { describe, it, expect } from 'vitest';
import { getToolsForHost, MAX_TOOLS_PER_REQUEST } from '@/tools';
import { managementTools } from '@/tools/management';
import { excelTools } from '@/tools';
import { powerPointTools } from '@/tools/powerpoint';
import { wordTools } from '@/tools/word';
import { outlookTools } from '@/tools/outlook';

const EXCEL_TOOL_NAMES = new Set(excelTools.map(t => t.name));
const PPT_TOOL_NAMES = new Set(powerPointTools.map(t => t.name));
const WORD_TOOL_NAMES = new Set(wordTools.map(t => t.name));
const OUTLOOK_TOOL_NAMES = new Set(outlookTools.map(t => t.name));
const MANAGEMENT_TOOL_NAMES = new Set(managementTools.map(t => t.name));

describe('host tools limit', () => {
  it('getToolsForHost("excel") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('excel');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("powerpoint") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('powerpoint');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("word") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('word');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("unknown") returns an empty array', () => {
    const tools = getToolsForHost('unknown' as never);
    expect(tools).toEqual([]);
  });
});

// ─── Host isolation ───────────────────────────────────────────────────────────
// Excel tools must never appear in non-Excel sessions and vice versa.
// Office JS APIs are document-context-specific and only operate on the currently
// active document — cross-host tool calls are not supported by the platform.

describe('host tool isolation', () => {
  it('PowerPoint session contains no Excel-specific tools', () => {
    const pptSessionTools = new Set(getToolsForHost('powerpoint').map(t => t.name));
    for (const name of EXCEL_TOOL_NAMES) {
      expect(pptSessionTools.has(name), `Excel tool "${name}" must not appear in PowerPoint`).toBe(
        false
      );
    }
  });

  it('Word session contains no Excel-specific tools', () => {
    const wordSessionTools = new Set(getToolsForHost('word').map(t => t.name));
    for (const name of EXCEL_TOOL_NAMES) {
      expect(wordSessionTools.has(name), `Excel tool "${name}" must not appear in Word`).toBe(
        false
      );
    }
  });

  it('Outlook session contains no Excel-specific tools', () => {
    const outlookSessionTools = new Set(getToolsForHost('outlook').map(t => t.name));
    for (const name of EXCEL_TOOL_NAMES) {
      expect(outlookSessionTools.has(name), `Excel tool "${name}" must not appear in Outlook`).toBe(
        false
      );
    }
  });

  it('Excel session contains no PowerPoint-specific tools', () => {
    const excelSessionTools = new Set(getToolsForHost('excel').map(t => t.name));
    for (const name of PPT_TOOL_NAMES) {
      expect(excelSessionTools.has(name), `PowerPoint tool "${name}" must not appear in Excel`).toBe(
        false
      );
    }
  });

  it('Excel session contains no Word-specific tools', () => {
    const excelSessionTools = new Set(getToolsForHost('excel').map(t => t.name));
    for (const name of WORD_TOOL_NAMES) {
      expect(excelSessionTools.has(name), `Word tool "${name}" must not appear in Excel`).toBe(
        false
      );
    }
  });

  it('Excel session contains no Outlook-specific tools', () => {
    const excelSessionTools = new Set(getToolsForHost('excel').map(t => t.name));
    for (const name of OUTLOOK_TOOL_NAMES) {
      expect(
        excelSessionTools.has(name),
        `Outlook tool "${name}" must not appear in Excel`
      ).toBe(false);
    }
  });

  it('management tools are included in every host session', () => {
    for (const host of ['excel', 'powerpoint', 'word', 'outlook'] as const) {
      const sessionToolNames = new Set(getToolsForHost(host).map(t => t.name));
      for (const name of MANAGEMENT_TOOL_NAMES) {
        expect(
          sessionToolNames.has(name),
          `Management tool "${name}" must appear in ${host} session`
        ).toBe(true);
      }
    }
  });

  it('each host session contains only its own host tools plus management tools', () => {
    const hostToolSets: Record<string, Set<string>> = {
      excel: EXCEL_TOOL_NAMES,
      powerpoint: PPT_TOOL_NAMES,
      word: WORD_TOOL_NAMES,
      outlook: OUTLOOK_TOOL_NAMES,
    };

    for (const [host, ownTools] of Object.entries(hostToolSets)) {
      const sessionToolNames = getToolsForHost(host as Parameters<typeof getToolsForHost>[0]).map(
        t => t.name
      );
      for (const name of sessionToolNames) {
        const isOwnTool = ownTools.has(name);
        const isManagementTool = MANAGEMENT_TOOL_NAMES.has(name);
        expect(
          isOwnTool || isManagementTool,
          `Unexpected tool "${name}" in ${host} session (not a ${host} tool or management tool)`
        ).toBe(true);
      }
    }
  });
});
