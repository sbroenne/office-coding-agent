/**
 * Deck Orchestrator
 *
 * Manages a planner → worker multi-session flow for creating slide decks.
 * Each worker gets a fresh context window to avoid context exhaustion.
 */

import type { WebSocketCopilotClient } from '@/lib/websocket-client';
import type { SessionEvent } from '@github/copilot-sdk';
import { runSubSession } from '@/lib/session-factory';
import {
  submitPlanTool,
  getLastPlan,
  extractPlanFromEvents,
  parsePlanFromText,
} from '@/tools/planner';
import type { DeckPlan, SlidePlan } from '@/tools/planner';
import { getToolsForHost } from '@/tools';
import plannerPromptRaw from '@/services/ai/prompts/PLANNER_PROMPT.md?raw';
import workerPromptRaw from '@/services/ai/prompts/WORKER_PROMPT.md?raw';

export type SlideStatus = 'pending' | 'running' | 'done' | 'failed';

export interface SlideProgress {
  plan: SlidePlan;
  status: SlideStatus;
  error?: string;
}

export interface DeckOrchestratorCallbacks {
  /** Called when the planner finishes with a plan */
  onPlan: (plan: DeckPlan) => void;
  /** Called when a slide's status changes */
  onSlideProgress: (index: number, progress: SlideProgress) => void;
  /** Called with streaming text from planner or workers */
  onText: (text: string) => void;
  /** Called with tool events from workers (for UI progress indicators) */
  onWorkerEvent?: (slideIndex: number, event: SessionEvent) => void;
  /** Called when the entire deck is complete */
  onComplete: (results: SlideProgress[]) => void;
  /** Called on fatal error */
  onError: (error: string) => void;
}

/**
 * Build the worker prompt for a batch of slides.
 */
function buildWorkerPrompt(slides: SlidePlan[], totalSlides: number): string {
  const tasks = slides
    .map(
      slide =>
        `### Slide ${String(slide.index + 1)} of ${String(totalSlides)}
- **Title**: ${slide.title}
- **Layout**: ${slide.layout}
- **Content**: ${slide.content}`
    )
    .join('\n\n');

  return `${workerPromptRaw}

## Slide Tasks

Create these slides in order. For EACH slide: create → verify (full + bottom quadrants) → fix if needed → next.

${tasks}

Narrate your progress in the user's language.`;
}

export type DeckMode = 'fast' | 'deep';

/**
 * Run the full planner → worker orchestration.
 *
 * @param client - Shared WebSocket Copilot client
 * @param model - Model ID to use for all sessions
 * @param userPrompt - The user's original deck request
 * @param callbacks - Progress callbacks for UI updates
 * @param signal - Optional AbortSignal for cancellation
 * @param mode - 'fast' (3 slides/worker) or 'deep' (1 slide/worker, max quality)
 */
export async function orchestrateDeck(
  client: WebSocketCopilotClient,
  model: string,
  userPrompt: string,
  callbacks: DeckOrchestratorCallbacks,
  signal?: AbortSignal,
  mode: DeckMode = 'fast'
): Promise<void> {
  // fast = all slides in one session, deep = 1 slide per session
  const BATCH_SIZE = mode === 'deep' ? 1 : Infinity;
  // --- Phase 1: Planner ---
  callbacks.onText(
    `📋 Erstelle Plan… (${mode === 'deep' ? 'Deep — 1 Slide/Worker' : 'Fast — alle Slides in 1 Session'})\n`
  );

  const plannerResult = await runSubSession(
    client,
    {
      model,
      systemPrompt: plannerPromptRaw,
      tools: [submitPlanTool],
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

  // Get plan: try handler capture → event extraction → text parsing (in order)
  let plan =
    getLastPlan() ??
    extractPlanFromEvents(
      plannerResult.events as { type: string; data: Record<string, unknown> }[]
    );

  if (!plan || plan.slides.length === 0) {
    // Last resort: try to parse a JSON plan from the streamed text
    plan = parsePlanFromText(plannerResult.text);
  }

  if (!plan || plan.slides.length === 0) {
    console.error(
      '[deck-orchestrator] No plan found. Events:',
      plannerResult.events.map(e => e.type)
    );
    console.error('[deck-orchestrator] Planner text output:', plannerResult.text.slice(0, 500));
    callbacks.onError('Planner did not produce a valid slide plan.');
    return;
  }

  callbacks.onPlan(plan);
  callbacks.onText(`\n\n📋 Plan: ${String(plan.slides.length)} Slides\n`);

  // --- Phase 2: Workers (batched) ---
  const results: SlideProgress[] = plan.slides.map(slide => ({
    plan: slide,
    status: 'pending',
  }));

  const pptTools = getToolsForHost('powerpoint');

  for (let batchStart = 0; batchStart < plan.slides.length; batchStart += BATCH_SIZE) {
    if (signal?.aborted) {
      callbacks.onText('\n⚠️ Abgebrochen.\n');
      break;
    }

    const batchEnd = Math.min(batchStart + BATCH_SIZE, plan.slides.length);
    const batchSlides = plan.slides.slice(batchStart, batchEnd);
    const batchLabel = batchSlides.map(s => `${String(s.index + 1)}`).join(', ');

    // Mark batch as running
    for (let i = batchStart; i < batchEnd; i++) {
      results[i].status = 'running';
      callbacks.onSlideProgress(i, results[i]);
    }
    callbacks.onText(`\n🔄 Slides ${batchLabel}/${String(plan.slides.length)}…\n`);

    const workerPrompt = buildWorkerPrompt(batchSlides, plan.slides.length);
    const userMsg = batchSlides
      .map(s => `Slide ${String(s.index + 1)}: "${s.title}" — Layout: ${s.layout}. ${s.content}`)
      .join('\n');

    const workerResult = await runSubSession(
      client,
      { model, systemPrompt: workerPrompt, tools: pptTools },
      userMsg,
      event => {
        callbacks.onWorkerEvent?.(batchStart, event);
        if (event.type === 'assistant.message_delta') {
          callbacks.onText(event.data.deltaContent);
        }
      }
    );

    // Mark batch results
    for (let i = batchStart; i < batchEnd; i++) {
      results[i].status = workerResult.success ? 'done' : 'failed';
      if (!workerResult.success) results[i].error = workerResult.error;
      callbacks.onSlideProgress(i, results[i]);
    }

    if (workerResult.success) {
      callbacks.onText(`\n✅ Slides ${batchLabel} fertig\n`);
    } else {
      callbacks.onText(
        `\n❌ Slides ${batchLabel} fehlgeschlagen: ${workerResult.error ?? 'unknown'}\n`
      );
    }
  }

  // --- Phase 3: Summary ---
  const done = results.filter(r => r.status === 'done').length;
  const failed = results.filter(r => r.status === 'failed').length;
  callbacks.onText(`\n\n📊 Fertig: ${String(done)}/${String(plan.slides.length)} Slides erstellt`);
  if (failed > 0) {
    callbacks.onText(` (${String(failed)} fehlgeschlagen)`);
  }
  callbacks.onText('\n');

  callbacks.onComplete(results);
}
