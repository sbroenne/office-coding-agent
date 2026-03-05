import React, { type FC } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';

interface ActionBarProps {
  messageText: string;
  isRunning?: boolean;
  isLast?: boolean;
  onRegenerate?: () => void;
  onFeedback?: (kind: 'positive' | 'negative') => void;
  className?: string;
}

export const ActionBar: FC<ActionBarProps> = ({
  messageText,
  isRunning = false,
  isLast = false,
  onRegenerate,
  onFeedback,
  className,
}) => {
  const [isCopied, setIsCopied] = React.useState(false);

  const handleCopy = () => {
    if (!messageText) return;
    void navigator.clipboard.writeText(messageText).then(() => {
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 3000);
    });
  };

  // Hide when running or when it's not the last message (autohide behaviour matching VS Code)
  if (isRunning || (!isLast && false)) {
    return null;
  }

  const btnClass = cn(
    'h-6 w-6 rounded flex items-center justify-center p-1 transition-colors',
    'text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground'
  );

  return (
    <div
      className={cn(
        'aui-assistant-action-bar flex items-center gap-0.5 mt-1',
        'opacity-0 transition-opacity duration-150 group-hover/message:opacity-100',
        className
      )}
    >
      <button
        type="button"
        onClick={handleCopy}
        title={isCopied ? 'Copied!' : 'Copy'}
        aria-label={isCopied ? 'Copied!' : 'Copy'}
        className={btnClass}
      >
        <Codicon name={isCopied ? 'check' : 'copy'} className="text-sm" />
      </button>

      {onRegenerate && (
        <button
          type="button"
          onClick={onRegenerate}
          title="Regenerate response"
          aria-label="Regenerate response"
          className={btnClass}
        >
          <Codicon name="refresh" className="text-sm" />
        </button>
      )}

      {onFeedback && (
        <>
          <button
            type="button"
            onClick={() => onFeedback('positive')}
            title="Good response"
            aria-label="Good response"
            className={btnClass}
          >
            <Codicon name="thumbsup" className="text-sm" />
          </button>
          <button
            type="button"
            onClick={() => onFeedback('negative')}
            title="Bad response"
            aria-label="Bad response"
            className={btnClass}
          >
            <Codicon name="thumbsdown" className="text-sm" />
          </button>
        </>
      )}
    </div>
  );
};
