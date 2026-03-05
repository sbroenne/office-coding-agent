import { memo, useState, useCallback } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { humanizeToolName } from '@/utils/humanizeToolName';
import { toolResultSummary } from '@/utils/toolResultSummary';
import { getToolIcon } from '@/utils/toolIcon';
import type { ToolCallPart } from '@/types';

interface ToolProgressProps {
  part: ToolCallPart;
}

const ToolProgressImpl: React.FC<ToolProgressProps> = ({ part }) => {
  const { toolName, argsText, result, status } = part;
  const [isExpanded, setIsExpanded] = useState(false);
  const toggle = useCallback(() => setIsExpanded(prev => !prev), []);

  const statusType = status?.type ?? 'complete';
  const isRunning = statusType === 'running';
  const isCancelled = status?.type === 'incomplete' && status.reason === 'cancelled';
  const isError = statusType === 'incomplete' && !isCancelled;

  const friendlyName = humanizeToolName(toolName);
  const toolIcon = getToolIcon(toolName);
  const summary = !isRunning && result !== undefined ? toolResultSummary(result) : null;
  const hasDetails = !!(
    argsText ||
    result !== undefined ||
    (status?.type === 'incomplete' && status.error)
  );

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
        <span className="chat-thinking-icon">
          <Codicon name={toolIcon} className="text-xs" />
        </span>

        <Codicon
          name={
            isRunning ? 'circle-filled' : isError ? 'error' : isCancelled ? 'circle-slash' : 'check'
          }
          className={cn(
            'progress-status-icon shrink-0',
            isError && 'text-[var(--vscode-errorForeground)]',
            isCancelled && 'text-muted-foreground'
          )}
        />

        <span className={cn('progress-step', isCancelled && 'line-through')}>{friendlyName}</span>

        {summary && <span className="progress-summary">{String(summary)}</span>}

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

          {argsText && (
            <div className="tool-details-section">
              <p className="tool-details-label">Input</p>
              <pre className="tool-details-code">{argsText}</pre>
            </div>
          )}

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

export const ToolProgress = memo(ToolProgressImpl);
