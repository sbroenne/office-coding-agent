/**
 * Word Document Orchestrator
 *
 * Manages a planner → worker multi-session flow for creating/restructuring Word documents.
 * Each worker gets a fresh context window to avoid context exhaustion.
 */

import type { WebSocketCopilotClient } from '@/lib/websocket-client';
import type { SessionEvent } from '@github/copilot-sdk';
import { runSubSession } from '@/lib/session-factory';
import { submitDocumentPlanTool, getLastDocumentPlan } from '@/tools/planner/wordPlanner';
import type { DocumentPlan, SectionPlan } from '@/tools/planner/wordPlanner';
import { getToolsForHost } from '@/tools';
import wordPlannerPromptRaw from '@/services/ai/prompts/WORD_PLANNER_PROMPT.md?raw';
import wordWorkerPromptRaw from '@/services/ai/prompts/WORD_WORKER_PROMPT.md?raw';

export type SectionStatus = 'pending' | 'running' | 'done' | 'failed';

export interface SectionProgress {
  plan: SectionPlan;
  status: SectionStatus;
  error?: string;
}

export interface DocumentOrchestratorCallbacks {
  onPlan: (plan: DocumentPlan) => void;
  onSectionProgress: (index: number, progress: SectionProgress) => void;
  onText: (text: string) => void;
  onWorkerEvent?: (sectionIndex: number, event: SessionEvent) => void;
  onComplete: (results: SectionProgress[]) => void;
  onError: (error: string) => void;
}

export type DocumentMode = 'fast' | 'deep';

function buildWorkerPrompt(sections: SectionPlan[], totalSections: number): string {
  const tasks = sections
    .map(
      section =>
        `### Section ${String(section.index + 1)} of ${String(totalSections)}
- **Title**: ${section.title}
- **Type**: ${section.type}
- **Heading Level**: ${String(section.headingLevel)}
- **Content**: ${section.content}`
    )
    .join('\n\n');

  return `${wordWorkerPromptRaw}

## Section Tasks

Create these sections in order. For EACH section: create → verify → fix if needed → next.

${tasks}

Narrate your progress in the user's language.`;
}

/**
 * Run the full planner → worker orchestration for Word documents.
 */
export async function orchestrateDocument(
  client: WebSocketCopilotClient,
  model: string,
  userPrompt: string,
  callbacks: DocumentOrchestratorCallbacks,
  signal?: AbortSignal,
  mode: DocumentMode = 'fast'
): Promise<void> {
  const BATCH_SIZE = mode === 'deep' ? 1 : Infinity;

  callbacks.onText(
    `📋 Erstelle Plan… (${mode === 'deep' ? 'Deep — 1 Section/Worker' : 'Fast — alle Sections in 1 Session'})\n`
  );

  // --- Phase 1: Planner ---
  const plannerResult = await runSubSession(
    client,
    {
      model,
      systemPrompt: wordPlannerPromptRaw,
      tools: [submitDocumentPlanTool],
    },
    userPrompt,
    event => {
      if (event.type === 'assistant.message_delta') {
        callbacks.onText(event.data.deltaContent);
      }
    }
  );

  if (!plannerResult.success) {
    callbacks.onError(`Planner failed: ${plannerResult.error ?? 'unknown error'}`);
    return;
  }

  const plan = getLastDocumentPlan();

  if (!plan || plan.sections.length === 0) {
    callbacks.onError('Planner did not produce a valid document plan.');
    return;
  }

  callbacks.onPlan(plan);
  callbacks.onText(`\n\n📋 Plan: ${String(plan.sections.length)} Sections\n`);

  // --- Phase 2: Workers (batched) ---
  const results: SectionProgress[] = plan.sections.map(section => ({
    plan: section,
    status: 'pending' as SectionStatus,
  }));

  const wordTools = getToolsForHost('word');

  for (let batchStart = 0; batchStart < plan.sections.length; batchStart += BATCH_SIZE) {
    if (signal?.aborted) {
      callbacks.onText('\n⚠️ Abgebrochen.\n');
      break;
    }

    const batchEnd = Math.min(batchStart + BATCH_SIZE, plan.sections.length);
    const batchSections = plan.sections.slice(batchStart, batchEnd);
    const batchLabel = batchSections.map(s => `${String(s.index + 1)}`).join(', ');

    for (let i = batchStart; i < batchEnd; i++) {
      results[i].status = 'running';
      callbacks.onSectionProgress(i, results[i]);
    }
    callbacks.onText(`\n🔄 Sections ${batchLabel}/${String(plan.sections.length)}…\n`);

    const workerPrompt = buildWorkerPrompt(batchSections, plan.sections.length);
    const userMsg = batchSections
      .map(
        s =>
          `Section ${String(s.index + 1)}: "${s.title}" — Type: ${s.type}, Level: ${String(s.headingLevel)}. ${s.content}`
      )
      .join('\n');

    const workerResult = await runSubSession(
      client,
      { model, systemPrompt: workerPrompt, tools: wordTools },
      userMsg,
      event => {
        callbacks.onWorkerEvent?.(batchStart, event);
        if (event.type === 'assistant.message_delta') {
          callbacks.onText(event.data.deltaContent);
        }
      }
    );

    for (let i = batchStart; i < batchEnd; i++) {
      results[i].status = workerResult.success ? 'done' : 'failed';
      if (!workerResult.success) results[i].error = workerResult.error;
      callbacks.onSectionProgress(i, results[i]);
    }

    if (workerResult.success) {
      callbacks.onText(`\n✅ Sections ${batchLabel} fertig\n`);
    } else {
      callbacks.onText(
        `\n❌ Sections ${batchLabel} fehlgeschlagen: ${workerResult.error ?? 'unknown'}\n`
      );
    }
  }

  // --- Phase 3: Summary ---
  const done = results.filter(r => r.status === 'done').length;
  const failed = results.filter(r => r.status === 'failed').length;
  callbacks.onText(
    `\n\n📊 Fertig: ${String(done)}/${String(plan.sections.length)} Sections erstellt`
  );
  if (failed > 0) {
    callbacks.onText(` (${String(failed)} fehlgeschlagen)`);
  }
  callbacks.onText('\n');
  callbacks.onComplete(results);
}
