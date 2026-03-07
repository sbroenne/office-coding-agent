import { type FC } from 'react';
import { Codicon } from '@/components/Codicon';
import { MarkdownContent } from './MarkdownContent';
import { ToolGroup } from './ToolGroup';
import { ActionBar } from './ActionBar';
import type { ChatMessage, ToolCallPart, TextPart } from '@/types';

interface AssistantMessageProps {
  message: ChatMessage;
  isRunning: boolean;
  isLast: boolean;
  onRegenerate?: (messageId: string) => void;
  onFeedback?: (messageId: string, kind: 'positive' | 'negative') => void;
}

/**
 * Groups consecutive tool-call parts together, interspersed with text parts.
 * Tool parts are split into separate groups whenever phaseIndex changes,
 * producing one Working box per intent phase (matching VS Code's IChatTask behavior).
 * Returns an array of "segments": either a single text part or a group of tool-call parts.
 */
function segmentParts(
  content: ChatMessage['content']
): (
  | { type: 'text'; part: TextPart }
  | { type: 'tools'; parts: ToolCallPart[]; phaseIndex: number }
)[] {
  const segments: (
    | { type: 'text'; part: TextPart }
    | { type: 'tools'; parts: ToolCallPart[]; phaseIndex: number }
  )[] = [];

  let currentTools: ToolCallPart[] = [];
  let currentPhase = -1;

  for (const part of content) {
    if (part.type === 'tool-call') {
      const phase = part.phaseIndex ?? 0;
      if (currentTools.length > 0 && phase !== currentPhase) {
        // Phase boundary → flush current tool group and start a new one
        segments.push({ type: 'tools', parts: currentTools, phaseIndex: currentPhase });
        currentTools = [];
      }
      currentPhase = phase;
      currentTools.push(part);
    } else if (part.type === 'text') {
      if (currentTools.length > 0) {
        segments.push({ type: 'tools', parts: currentTools, phaseIndex: currentPhase });
        currentTools = [];
      }
      segments.push({ type: 'text', part });
    }
  }

  if (currentTools.length > 0) {
    segments.push({ type: 'tools', parts: currentTools, phaseIndex: currentPhase });
  }

  return segments;
}

export const AssistantMessage: FC<AssistantMessageProps> = ({
  message,
  isRunning,
  isLast,
  onRegenerate,
  onFeedback,
}) => {
  const { thinkingText, content, status } = message;

  const segments = segmentParts(content);
  const textParts = content.filter((p): p is TextPart => p.type === 'text');
  const hasText = textParts.some(p => p.text.trim().length > 0);
  const hasTools = content.some(p => p.type === 'tool-call');

  const isMessageRunning = status?.type === 'running';
  const isError = status?.type === 'incomplete' && status.reason === 'error';

  // When the agent ends with task_complete (no text response follows), surface
  // the summary as readable text — matching VS Code where the model always writes
  // a text response after tool use.
  const taskCompleteSummary = (() => {
    if (hasText || isMessageRunning) return null;
    const part = content.find(
      (p): p is ToolCallPart => p.type === 'tool-call' && p.toolName === 'task_complete'
    );
    if (!part?.argsText) return null;
    try {
      const args = JSON.parse(part.argsText) as Record<string, unknown>;
      const s = args.summary;
      return typeof s === 'string' && s.trim() ? s.trim() : null;
    } catch {
      return null;
    }
  })();

  // Full message text for copy button (include task_complete summary when present)
  const fullText = taskCompleteSummary
    ? textParts.map(p => p.text).join('') || taskCompleteSummary
    : textParts.map(p => p.text).join('');

  // Show inline "Thinking…" shimmer ONLY before any tools or text appear.
  // Once tools exist, the Working box takes over all thinking indication
  // (its spinner shows thinkingText between tool steps). Once text streams,
  // the text itself is the visible progress — no shimmer needed.
  const showThinking =
    isLast && isRunning && isMessageRunning && !!thinkingText && !hasText && !hasTools;

  return (
    <div
      className="aui-assistant-message-root group/message fade-in slide-in-from-bottom-1 relative w-full animate-in py-2 duration-150"
      data-role="assistant"
    >
      <div className="aui-assistant-message-content flex flex-col wrap-break-word px-4 text-foreground text-[13px] leading-[1.5em]">
        {/* Inline working progress */}
        {showThinking && (
          <div
            className="inline-working-progress progress-container shimmer-progress"
            style={{ order: -2 }}
          >
            <Codicon name="circle-filled" className="progress-status-icon shrink-0" />
            <span className="progress-step">{thinkingText}</span>
          </div>
        )}

        {/* Tool groups and text */}
        {segments.map((seg, i) => {
          if (seg.type === 'tools') {
            // Only the LAST tool segment in a running message shows the live spinner/running state
            const isLastToolSegment = segments.slice(i + 1).every(s => s.type !== 'tools');
            return (
              <ToolGroup
                key={`tools-${seg.phaseIndex}-${i}`}
                parts={seg.parts}
                isMessageRunning={isLast && isRunning && isMessageRunning && isLastToolSegment}
                thinkingText={isLastToolSegment ? thinkingText : null}
              />
            );
          }
          // Only render non-empty text parts (or last part while streaming)
          const text = seg.part.text;
          if (!text && !(isLast && isRunning)) return null;
          return <MarkdownContent key={`text-${i}`} text={text} />;
        })}

        {/* task_complete summary — shown when the agent ends with task_complete but no text */}
        {taskCompleteSummary && <MarkdownContent text={taskCompleteSummary} />}

        {/* Error indicator */}
        {isError && status.error && (
          <div
            className="mt-2 rounded-[var(--vscode-cornerRadius-medium)] border p-3 text-sm"
            style={{
              borderColor: 'var(--vscode-errorForeground)',
              background: 'color-mix(in srgb, var(--vscode-errorForeground) 10%, transparent)',
              color: 'var(--vscode-errorForeground)',
            }}
          >
            <span className="line-clamp-2">{status.error}</span>
          </div>
        )}

        <ActionBar
          messageText={fullText}
          isRunning={isRunning && isLast}
          isLast={isLast}
          onRegenerate={onRegenerate ? () => onRegenerate(message.id) : undefined}
          onFeedback={onFeedback ? kind => onFeedback(message.id, kind) : undefined}
        />
      </div>
    </div>
  );
};
