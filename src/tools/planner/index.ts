/**
 * Planner tool: submit_plan
 *
 * The planner agent calls this tool to submit its structured slide plan.
 * The orchestrator intercepts the tool result to extract the plan.
 */

import type { Tool } from '@github/copilot-sdk';

export interface SlidePlan {
  index: number;
  title: string;
  layout: string;
  content: string;
}

export interface DeckPlan {
  slides: SlidePlan[];
}

/** Sentinel value to identify planner tool results */
export const PLAN_TOOL_NAME = 'submit_plan';

/** Stores the last plan received by the tool handler */
let lastPlan: DeckPlan | null = null;

/** Get and clear the last captured plan */
export function getLastPlan(): DeckPlan | null {
  const plan = lastPlan;
  lastPlan = null;
  return plan;
}

export const submitPlanTool: Tool = {
  name: PLAN_TOOL_NAME,
  description:
    'Submit the structured slide plan. Call this exactly once with the complete plan for all slides.',
  parameters: {
    type: 'object',
    properties: {
      slides: {
        type: 'array',
        description: 'Array of slide plans, one per slide.',
        items: {
          type: 'object',
          properties: {
            index: { type: 'number', description: '0-based slide index.' },
            title: { type: 'string', description: 'Slide title.' },
            layout: {
              type: 'string',
              description:
                'Layout type: title-dark, title-light, agenda, stat-cards, bullet-list, two-column, three-column-cards, card-grid, table, timeline, quote, case-study, image-text.',
            },
            content: {
              type: 'string',
              description:
                'Detailed content description. Specific enough for a slide creator to execute without seeing the original request.',
            },
          },
          required: ['index', 'title', 'layout', 'content'],
        },
      },
    },
    required: ['slides'],
  },
  handler: (args: unknown) => {
    const plan = args as DeckPlan;
    // Capture the plan so the orchestrator can read it
    if (Array.isArray(plan.slides) && plan.slides.length > 0) {
      lastPlan = plan;
    }
    return `Plan received: ${String(plan.slides?.length ?? 0)} slides.`;
  },
};

/**
 * Extract plan from tool call events.
 */
export function extractPlanFromEvents(
  events: { type: string; data: Record<string, unknown> }[]
): DeckPlan | null {
  for (const event of events) {
    if (event.type === 'tool.execution_start') {
      const data = event.data as { toolName?: string; arguments?: unknown };
      if (data.toolName === PLAN_TOOL_NAME && data.arguments) {
        const plan = data.arguments as DeckPlan;
        if (Array.isArray(plan.slides) && plan.slides.length > 0) {
          return plan;
        }
      }
    }
  }
  return null;
}

/**
 * Try to coerce a parsed JSON value into a DeckPlan.
 * Accepts `{ slides: [...] }` or a bare array of slide objects.
 */
function coerceToPlan(value: unknown): DeckPlan | null {
  if (value && typeof value === 'object') {
    // { slides: [...] }
    const obj = value as Record<string, unknown>;
    if (Array.isArray(obj.slides) && obj.slides.length > 0) {
      return obj as unknown as DeckPlan;
    }
    // { plan: { slides: [...] } }
    if (obj.plan && typeof obj.plan === 'object') {
      const inner = obj.plan as Record<string, unknown>;
      if (Array.isArray(inner.slides) && inner.slides.length > 0) {
        return inner as unknown as DeckPlan;
      }
    }
  }
  // Bare array of slide objects
  if (Array.isArray(value) && value.length > 0) {
    const first = value[0] as Record<string, unknown>;
    if (typeof first.title === 'string' && typeof first.layout === 'string') {
      const slides = (value as Record<string, unknown>[]).map((s, i) => ({
        index: typeof s.index === 'number' ? s.index : i,
        title: typeof s.title === 'string' ? s.title : '',
        layout: typeof s.layout === 'string' ? s.layout : '',
        content: typeof s.content === 'string' ? s.content : '',
      }));
      return { slides };
    }
  }
  return null;
}

/**
 * Extract the outermost JSON value (object or array) from a string.
 * Uses bracket counting instead of lazy regex to correctly handle nested braces.
 */
function extractOutermostJson(text: string): string | null {
  const startIdx = text.search(/[{[]/);
  if (startIdx === -1) return null;

  const openChar = text[startIdx];
  const closeChar = openChar === '{' ? '}' : ']';
  let depth = 0;
  let inString = false;
  let escape = false;

  for (let i = startIdx; i < text.length; i++) {
    const ch = text[i];
    if (escape) {
      escape = false;
      continue;
    }
    if (ch === '\\' && inString) {
      escape = true;
      continue;
    }
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === openChar) depth++;
    if (ch === closeChar) depth--;
    if (depth === 0) {
      return text.slice(startIdx, i + 1);
    }
  }
  return null;
}

/**
 * Parse a structured markdown plan into a DeckPlan.
 * Handles numbered lists like:
 *   1. **Title** — Layout: title-dark. Content description...
 *   ### Slide 1: Title\n- Layout: title-dark\n- Content: ...
 */
function parseMarkdownPlan(text: string): DeckPlan | null {
  const slides: SlidePlan[] = [];

  // Pattern 1: "### Slide N: Title" blocks with "- Layout:" and "- Content:" lines
  const blockPattern =
    /###\s*Slide\s+(\d+)[:\s]*([^\n]+)\n(?:[\s\S]*?-\s*\*{0,2}Layout\*{0,2}[:\s]+([^\n]+)\n)?(?:[\s\S]*?-\s*\*{0,2}Content\*{0,2}[:\s]+([^\n]+))?/gi;
  let match: RegExpExecArray | null;
  while ((match = blockPattern.exec(text)) !== null) {
    slides.push({
      index: slides.length,
      title: match[2].trim().replace(/\*{1,2}/g, ''),
      layout: (match[3] ?? 'bullet-list').trim().replace(/\*{1,2}/g, ''),
      content: (match[4] ?? match[2]).trim().replace(/\*{1,2}/g, ''),
    });
  }
  if (slides.length > 0) return { slides };

  // Pattern 2: Numbered list "N. **Title** — Layout: X. Content..."
  const linePattern =
    /^\s*\d+\.\s*\*{0,2}([^*\n]+?)\*{0,2}\s*[—–-]\s*(?:Layout:\s*(\S+))?[.,]?\s*(.*)/gm;
  while ((match = linePattern.exec(text)) !== null) {
    slides.push({
      index: slides.length,
      title: match[1].trim(),
      layout: (match[2] ?? 'bullet-list').trim(),
      content: (match[3] ?? match[1]).trim(),
    });
  }
  if (slides.length >= 2) return { slides };

  return null;
}

/**
 * Last-resort: parse a slide plan from the planner's text output.
 * The model sometimes outputs the plan as text instead of calling the tool.
 *
 * Tries in order:
 * 1. JSON object/array in a code block (```json ... ```)
 * 2. Raw JSON object/array in the text
 * 3. Structured markdown (### Slide N or numbered lists)
 */
export function parsePlanFromText(text: string): DeckPlan | null {
  const candidates: string[] = [];

  // 1. Extract JSON from fenced code blocks (all occurrences)
  const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)```/g;
  let cbMatch: RegExpExecArray | null;
  while ((cbMatch = codeBlockRegex.exec(text)) !== null) {
    const inner = cbMatch[1].trim();
    if (inner.startsWith('{') || inner.startsWith('[')) {
      candidates.push(inner);
    }
  }

  // 2. Try extracting outermost JSON from the full text
  const outerJson = extractOutermostJson(text);
  if (outerJson) candidates.push(outerJson);

  // 3. Try the entire text as-is
  candidates.push(text.trim());

  for (const candidate of candidates) {
    try {
      const parsed: unknown = JSON.parse(candidate);
      const plan = coerceToPlan(parsed);
      if (plan) return plan;
    } catch {
      // Not valid JSON — continue
    }
  }

  // 4. Try structured markdown parsing
  return parseMarkdownPlan(text);
}
