import { type FC, type ReactNode, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { ToolProgress } from './ToolProgress';
import type { ToolCallPart } from '@/types';

const TOOL_LABELS = ['Processing', 'Preparing', 'Loading', 'Analyzing', 'Evaluating'];

function pickLabel(pool: string[]): string {
  return pool[Math.floor(Math.random() * pool.length)];
}

interface WorkingCollapsibleProps {
  isRunning: boolean;
  toolCount: number;
  children?: ReactNode;
}

const WorkingCollapsible: FC<WorkingCollapsibleProps> = ({ isRunning, toolCount, children }) => {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasAutoCollapsed = useRef(false);

  const spinnerLabel = useMemo(() => (isRunning ? pickLabel(TOOL_LABELS) : ''), [isRunning]);

  useEffect(() => {
    if (!isRunning && !hasAutoCollapsed.current) {
      hasAutoCollapsed.current = true;
      setIsExpanded(false);
    }
  }, [isRunning]);

  const toggle = useCallback(() => setIsExpanded(prev => !prev), []);

  const completionTitle = `Finished with ${toolCount} step${toolCount !== 1 ? 's' : ''}`;

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
          {isRunning ? 'Working' : completionTitle}
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
        {isRunning && (
          <div className="chat-thinking-spinner-item chat-thinking-tool-wrapper">
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
}

export const ToolGroup: FC<ToolGroupProps> = ({ parts }) => {
  const isRunning = parts.some(p => p.status?.type === 'running');

  return (
    <WorkingCollapsible isRunning={isRunning} toolCount={parts.length}>
      {parts.map(part => (
        <ToolProgress key={part.toolCallId} part={part} />
      ))}
    </WorkingCollapsible>
  );
};
