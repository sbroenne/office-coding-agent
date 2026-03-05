import { memo, useState, useCallback } from 'react';
import {
  type ToolCallMessagePartStatus,
  type ToolCallMessagePartComponent,
} from '@assistant-ui/react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { humanizeToolName } from '@/utils/humanizeToolName';
import { toolResultSummary } from '@/utils/toolResultSummary';
import { getToolIcon } from '@/utils/toolIcon';

type ToolStatus = ToolCallMessagePartStatus['type'];

/**
 * VS Code-style progress line for a tool invocation.
 *
 * Matches VS Code's `.progress-container` layout:
 *   [tool-icon] [shimmer-text]          — while running
 *   [check]     [muted past-tense text] — when complete
 *
 * No borders, no cards. Flat inline progress line.
 */
const ToolFallbackImpl: ToolCallMessagePartComponent = ({ toolName, argsText, result, status }) => {
  const [isExpanded, setIsExpanded] = useState(false);
  const toggle = useCallback(() => setIsExpanded(prev => !prev), []);

  const statusType: ToolStatus = status?.type ?? 'complete';
  const isRunning = statusType === 'running';
  const isCancelled = status?.type === 'incomplete' && status.reason === 'cancelled';
  const isError = statusType === 'incomplete' && !isCancelled;

  const friendlyName = humanizeToolName(toolName);
  const toolIcon = getToolIcon(toolName);
  const summary = !isRunning && result !== undefined ? toolResultSummary(result) : null;
  const hasDetails = !!(argsText || result !== undefined || (status?.type === 'incomplete' && status.error));

  return (
    <div
      className={cn('chat-tool-progress-wrapper', isCancelled && 'opacity-60')}
      data-slot="tool-fallback-root"
    >
      {/* Progress line: [icon] [text] [hover-chevron] */}
      <button
        type="button"
        onClick={hasDetails ? toggle : undefined}
        className={cn(
          'progress-container',
          isRunning && 'shimmer-progress',
          !hasDetails && 'pointer-events-none'
        )}
        aria-expanded={hasDetails ? isExpanded : undefined}
        data-slot="tool-fallback-trigger"
      >
        {/* Tool-type icon on the chain-of-thought line */}
        <span className="chat-thinking-icon">
          <Codicon name={toolIcon} className="text-xs" />
        </span>

        {/* Status icon: hidden in shimmer mode, check when done */}
        <Codicon
          name={
            isRunning
              ? 'circle-filled'
              : isError
                ? 'error'
                : isCancelled
                  ? 'circle-slash'
                  : 'check'
          }
          className={cn(
            'progress-status-icon shrink-0',
            isError && 'text-[var(--vscode-errorForeground)]',
            isCancelled && 'text-muted-foreground'
          )}
        />

        {/* Tool name + summary */}
        <span
          className={cn(
            'progress-step',
            isCancelled && 'line-through'
          )}
        >
          {friendlyName}
        </span>

        {summary && (
          <span className="progress-summary">{String(summary)}</span>
        )}

        {/* Hover chevron (appears on hover, like VS Code) */}
        {hasDetails && (
          <Codicon
            name={isExpanded ? 'chevron-down' : 'chevron-right'}
            className="chat-collapsible-hover-chevron"
          />
        )}
      </button>

      {/* Expandable tool details */}
      {isExpanded && hasDetails && (
        <div className="tool-details-expanded">
          {/* Error */}
          {status?.type === 'incomplete' && !!status.error && (
            <div className="tool-details-section">
              <p
                className="text-xs font-semibold"
                style={{ color: 'var(--vscode-errorForeground)' }}
              >
                {isCancelled ? 'Cancelled:' : 'Error:'}
              </p>
              <p style={{ color: 'var(--vscode-errorForeground)' }}>
                {typeof status.error === 'string'
                  ? status.error
                  : JSON.stringify(status.error as object)}
              </p>
            </div>
          )}

          {/* Input */}
          {argsText && (
            <div className="tool-details-section">
              <p className="tool-details-label">Input</p>
              <pre className="tool-details-code">{argsText}</pre>
            </div>
          )}

          {/* Output */}
          {result !== undefined && !isCancelled && (
            <div className="tool-details-section">
              <p className="tool-details-label">Output</p>
              <pre className="tool-details-code">
                {typeof result === 'string' ? result : JSON.stringify(result as object, null, 2)}
              </pre>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

const ToolFallback = memo(ToolFallbackImpl) as ToolCallMessagePartComponent;
ToolFallback.displayName = 'ToolFallback';

export { ToolFallback };
