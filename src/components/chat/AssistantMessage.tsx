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
 * Returns an array of "segments": either a single text part or a group of tool-call parts.
 */
function segmentParts(
  content: ChatMessage['content']
): ({ type: 'text'; part: TextPart } | { type: 'tools'; parts: ToolCallPart[] })[] {
  const segments: ({ type: 'text'; part: TextPart } | { type: 'tools'; parts: ToolCallPart[] })[] =
    [];

  let currentTools: ToolCallPart[] = [];

  for (const part of content) {
    if (part.type === 'tool-call') {
      currentTools.push(part);
    } else if (part.type === 'text') {
      if (currentTools.length > 0) {
        segments.push({ type: 'tools', parts: currentTools });
        currentTools = [];
      }
      segments.push({ type: 'text', part });
    }
  }

  if (currentTools.length > 0) {
    segments.push({ type: 'tools', parts: currentTools });
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
  const hasRunningTool = content.some(p => p.type === 'tool-call' && p.status?.type === 'running');

  // Full message text for copy button
  const fullText = textParts.map(p => p.text).join('');

  const isMessageRunning = status?.type === 'running';
  const isError = status?.type === 'incomplete' && status.reason === 'error';

  // Show inline working progress: only on last message, when thinking, no text yet, no running tool
  const showThinking =
    isLast && isRunning && isMessageRunning && !!thinkingText && !hasText && !hasRunningTool;

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
            return <ToolGroup key={i} parts={seg.parts} />;
          }
          // Only render non-empty text parts (or last part while streaming)
          const text = seg.part.text;
          if (!text && !(isLast && isRunning)) return null;
          return <MarkdownContent key={i} text={text} />;
        })}

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
