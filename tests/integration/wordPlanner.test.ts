import { describe, it, expect, beforeEach } from 'vitest';
import {
  submitDocumentPlanTool,
  getLastDocumentPlan,
  DOCUMENT_PLAN_TOOL_NAME,
} from '@/tools/planner/wordPlanner';
import type { DocumentPlan } from '@/tools/planner/wordPlanner';

const mockInvocation = {} as Parameters<typeof submitDocumentPlanTool.handler>[1];

describe('wordPlanner', () => {
  beforeEach(() => {
    // Clear any leftover plan
    getLastDocumentPlan();
  });

  describe('DOCUMENT_PLAN_TOOL_NAME', () => {
    it('equals "submit_document_plan"', () => {
      expect(DOCUMENT_PLAN_TOOL_NAME).toBe('submit_document_plan');
    });
  });

  describe('submitDocumentPlanTool', () => {
    it('has correct name and description', () => {
      expect(submitDocumentPlanTool.name).toBe('submit_document_plan');
      expect(submitDocumentPlanTool.description).toBeTruthy();
    });

    it('has a valid JSON Schema for parameters', () => {
      const params = submitDocumentPlanTool.parameters as Record<string, unknown>;
      expect(params.type).toBe('object');
      expect(params.required).toContain('sections');
    });

    it('handler captures a valid plan', () => {
      const plan: DocumentPlan = {
        sections: [
          { index: 0, title: 'Introduction', type: 'paragraph', headingLevel: 1, content: 'Overview of the topic' },
          { index: 1, title: 'Details', type: 'bullet-list', headingLevel: 2, content: 'Key points' },
        ],
      };

      const result = submitDocumentPlanTool.handler(plan, mockInvocation);
      expect(result).toContain('2 sections');

      const captured = getLastDocumentPlan();
      expect(captured).not.toBeNull();
      expect(captured!.sections).toHaveLength(2);
      expect(captured!.sections[0].title).toBe('Introduction');
      expect(captured!.sections[1].type).toBe('bullet-list');
    });

    it('getLastDocumentPlan clears after reading', () => {
      const plan: DocumentPlan = {
        sections: [{ index: 0, title: 'Test', type: 'paragraph', headingLevel: 1, content: 'Content' }],
      };
      submitDocumentPlanTool.handler(plan, mockInvocation);

      const first = getLastDocumentPlan();
      expect(first).not.toBeNull();

      const second = getLastDocumentPlan();
      expect(second).toBeNull();
    });

    it('handler ignores empty sections array', () => {
      const plan: DocumentPlan = { sections: [] };
      submitDocumentPlanTool.handler(plan, mockInvocation);

      const captured = getLastDocumentPlan();
      expect(captured).toBeNull();
    });

    it('handler ignores invalid input without sections', () => {
      submitDocumentPlanTool.handler({}, mockInvocation);

      const captured = getLastDocumentPlan();
      expect(captured).toBeNull();
    });
  });
});
