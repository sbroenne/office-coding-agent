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
  handler: async (args: unknown) => {
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
  events: Array<{ type: string; data: Record<string, unknown> }>,
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
