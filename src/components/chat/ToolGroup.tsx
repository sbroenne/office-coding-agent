import { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { ToolProgress } from './ToolProgress';
import type { ToolCallPart } from '@/types';

interface WorkingCollapsibleProps {
  isRunning: boolean;
  /** Phase label from report_intent — the Working box header text (VS Code: IChatTask.content). */
  phaseLabel?: string;
  /** Show the inter-step spinner when true (no tool actively executing). */
  showSpinner: boolean;
  /** Label for the inter-step spinner — the actual thinking/intent text. */
  spinnerLabel: string;
  children?: ReactNode;
}

const WorkingCollapsible: FC<WorkingCollapsibleProps> = ({
  isRunning,
  phaseLabel,
  showSpinner,
  spinnerLabel,
  children,
}) => {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasAutoCollapsed = useRef(false);

  useEffect(() => {
    if (!isRunning && !hasAutoCollapsed.current) {
      hasAutoCollapsed.current = true;
      setIsExpanded(false);
    }
  }, [isRunning]);

  const toggle = useCallback(() => setIsExpanded(prev => !prev), []);

  // VS Code: IChatTask shows task.content in BOTH running and done states.
  // Use the phase label if available; fall back to "Working".
  const headerLabel = phaseLabel ?? 'Working';

  return (
    <div
      className={cn(
        'chat-thinking-box',
        isRunning && 'chat-thinking-active',
        !isExpanded && 'chat-thinking-collapsed'
      )}
      style={{ order: -1 }}
    >
      <button
        type="button"
        onClick={toggle}
        className="chat-thinking-header"
        aria-expanded={isExpanded}
      >
        <Codicon
          name={isRunning ? 'circle-filled' : 'check'}
          className={cn('chat-thinking-header-icon', !isRunning && 'text-muted-foreground')}
        />
        <span
          className={cn(isRunning ? 'chat-thinking-title-shimmer' : 'chat-thinking-title-done')}
        >
          {headerLabel}
        </span>
        <Codicon
          name={isExpanded ? 'chevron-down' : 'chevron-right'}
          className="chat-collapsible-hover-chevron"
        />
      </button>

      <div
        className="chat-thinking-collapsible"
        style={{ display: isExpanded ? undefined : 'none' }}
      >
        {children}
        {showSpinner && (
          <div
            className="chat-thinking-spinner-item chat-thinking-tool-wrapper"
            data-testid="working-spinner"
          >
            <span className="chat-thinking-icon">
              <Codicon name="circle-filled" className="text-xs" />
            </span>
            <span className="chat-thinking-spinner-label">{spinnerLabel}</span>
          </div>
        )}
      </div>
    </div>
  );
};

interface ToolGroupProps {
  parts: ToolCallPart[];
  /** Whether the parent message is still streaming (not just tools) */
  isMessageRunning?: boolean;
  /** Current thinking/intent text — shown as spinner label between tool steps. */
  thinkingText?: string | null;
}

export const ToolGroup: FC<ToolGroupProps> = ({
  parts,
  isMessageRunning = false,
  thinkingText,
}) => {
  const hasRunningTool = parts.some(p => p.status?.type === 'running');
  // Show as "running" if any tool is running OR the message is still streaming
  const isRunning = hasRunningTool || isMessageRunning;
  // Show spinner between tool steps: message still running, no tool actively executing,
  // and we have thinking text to display. This is the inter-step "Thinking…" indicator.
  const showSpinner = isRunning && !hasRunningTool && !!thinkingText;
  const spinnerLabel = thinkingText ?? 'Thinking…';
  // Phase label: use the first part's phaseLabel (all parts in a group share the same phase)
  const phaseLabel = parts[0]?.phaseLabel;

  return (
    <WorkingCollapsible
      isRunning={isRunning}
      phaseLabel={phaseLabel}
      showSpinner={showSpinner}
      spinnerLabel={spinnerLabel}
    >
      {parts.map(part => (
        <ToolProgress key={part.toolCallId} part={part} />
      ))}
    </WorkingCollapsible>
  );
};
