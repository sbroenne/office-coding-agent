/**
 * Word Document Planner tool: submit_document_plan
 *
 * The planner agent calls this tool to submit its structured document plan.
 * The orchestrator intercepts the tool result to extract the plan.
 */

import type { Tool } from '@github/copilot-sdk';

export interface SectionPlan {
  index: number;
  title: string;
  type: string;
  headingLevel: number;
  content: string;
}

export interface DocumentPlan {
  sections: SectionPlan[];
}

export const DOCUMENT_PLAN_TOOL_NAME = 'submit_document_plan';

let lastDocumentPlan: DocumentPlan | null = null;

export function getLastDocumentPlan(): DocumentPlan | null {
  const plan = lastDocumentPlan;
  lastDocumentPlan = null;
  return plan;
}

export const submitDocumentPlanTool: Tool = {
  name: DOCUMENT_PLAN_TOOL_NAME,
  description:
    'Submit the structured document section plan. Call this exactly once with the complete plan for all sections.',
  parameters: {
    type: 'object',
    properties: {
      sections: {
        type: 'array',
        description: 'Array of section plans, one per section.',
        items: {
          type: 'object',
          properties: {
            index: { type: 'number', description: '0-based section index.' },
            title: { type: 'string', description: 'Section title or heading.' },
            type: {
              type: 'string',
              description:
                'Content type: heading, paragraph, bullet-list, numbered-list, table, quote, summary.',
            },
            headingLevel: {
              type: 'number',
              description:
                'Heading level (1, 2, or 3). Use 1 for main sections, 2 for subsections.',
            },
            content: {
              type: 'string',
              description:
                'Detailed content description. Specific enough for a section writer to execute without seeing the original request.',
            },
          },
          required: ['index', 'title', 'type', 'headingLevel', 'content'],
        },
      },
    },
    required: ['sections'],
  },
  handler: (args: unknown) => {
    const plan = args as DocumentPlan;
    if (Array.isArray(plan.sections) && plan.sections.length > 0) {
      lastDocumentPlan = plan;
    }
    return `Document plan received: ${String(plan.sections?.length ?? 0)} sections.`;
  },
};
