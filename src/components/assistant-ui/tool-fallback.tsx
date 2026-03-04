import { memo, useCallback, useRef, useState } from 'react';
import {
  useScrollLock,
  type ToolCallMessagePartStatus,
  type ToolCallMessagePartComponent,
} from '@assistant-ui/react';
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from '@/components/ui/collapsible';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { humanizeToolName } from '@/utils/humanizeToolName';
import { toolResultSummary } from '@/utils/toolResultSummary';

const ANIMATION_DURATION = 200;

export type ToolFallbackRootProps = Omit<
  React.ComponentProps<typeof Collapsible>,
  'open' | 'onOpenChange'
> & {
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
  defaultOpen?: boolean;
};

function ToolFallbackRoot({
  className,
  open: controlledOpen,
  onOpenChange: controlledOnOpenChange,
  defaultOpen = false,
  children,
  ...props
}: ToolFallbackRootProps) {
  const collapsibleRef = useRef<HTMLDivElement>(null);
  const [uncontrolledOpen, setUncontrolledOpen] = useState(defaultOpen);
  const lockScroll = useScrollLock(collapsibleRef, ANIMATION_DURATION);

  const isControlled = controlledOpen !== undefined;
  const isOpen = isControlled ? controlledOpen : uncontrolledOpen;

  const handleOpenChange = useCallback(
    (open: boolean) => {
      if (!open) {
        lockScroll();
      }
      if (!isControlled) {
        setUncontrolledOpen(open);
      }
      controlledOnOpenChange?.(open);
    },
    [lockScroll, isControlled, controlledOnOpenChange]
  );

  return (
    <Collapsible
      ref={collapsibleRef}
      data-slot="tool-fallback-root"
      open={isOpen}
      onOpenChange={handleOpenChange}
      className={cn('aui-tool-fallback-root group/tool-fallback-root w-full', className)}
      style={
        {
          '--animation-duration': `${ANIMATION_DURATION}ms`,
        } as React.CSSProperties
      }
      {...props}
    >
      {children}
    </Collapsible>
  );
}

type ToolStatus = ToolCallMessagePartStatus['type'];

/** Maps tool status to codicon name */
const statusCodiconMap: Record<ToolStatus, string> = {
  running: 'loading~spin',
  complete: 'check',
  incomplete: 'error',
  'requires-action': 'warning',
};

function ToolFallbackTrigger({
  toolName,
  status,
  result,
  className,
  ...props
}: React.ComponentProps<typeof CollapsibleTrigger> & {
  toolName: string;
  status?: ToolCallMessagePartStatus;
  result?: unknown;
}) {
  const statusType = status?.type ?? 'complete';
  const isRunning = statusType === 'running';
  const isCancelled = status?.type === 'incomplete' && status.reason === 'cancelled';

  const codiconName = statusCodiconMap[statusType];
  const friendlyName = humanizeToolName(toolName);
  const summary = !isRunning && result !== undefined ? toolResultSummary(result) : null;

  return (
    <CollapsibleTrigger
      data-slot="tool-fallback-trigger"
      className={cn(
        'aui-tool-fallback-trigger group/trigger flex w-full items-center gap-2 px-3 text-sm transition-colors',
        className
      )}
      {...props}
    >
      <Codicon
        name={codiconName}
        className={cn(
          'aui-tool-fallback-trigger-icon shrink-0 text-sm',
          isCancelled && 'text-muted-foreground',
          isRunning && 'codicon-modifier-spin'
        )}
      />
      <span
        data-slot="tool-fallback-trigger-label"
        className={cn(
          'aui-tool-fallback-trigger-label-wrapper relative inline-block text-left leading-none',
          isCancelled && 'text-muted-foreground line-through'
        )}
      >
        <span className={cn(isRunning && 'chat-thinking-shimmer-text')}>
          <b>{friendlyName}</b>
        </span>
      </span>
      {summary && (
        <span className="ml-1 min-w-0 flex-1 truncate text-xs text-muted-foreground">
          {summary}
        </span>
      )}
      <Codicon
        name="chevron-down"
        data-slot="tool-fallback-trigger-chevron"
        className={cn(
          'aui-tool-fallback-trigger-chevron ml-auto shrink-0 text-sm',
          'transition-transform duration-200 ease-out',
          'group-data-[state=closed]/trigger:-rotate-90',
          'group-data-[state=open]/trigger:rotate-0'
        )}
      />
    </CollapsibleTrigger>
  );
}

function ToolFallbackContent({
  className,
  children,
  ...props
}: React.ComponentProps<typeof CollapsibleContent>) {
  return (
    <CollapsibleContent
      data-slot="tool-fallback-content"
      className={cn(
        'aui-tool-fallback-content relative overflow-hidden text-sm outline-none',
        'group/collapsible-content ease-out',
        className
      )}
      {...props}
    >
      <div
        className="mt-2 flex flex-col gap-2 pt-2"
        style={{ borderTop: '1px solid var(--vscode-widget-border)' }}
      >
        {children}
      </div>
    </CollapsibleContent>
  );
}

function ToolFallbackArgs({
  argsText,
  className,
  ...props
}: React.ComponentProps<'div'> & {
  argsText?: string;
}) {
  if (!argsText) return null;

  return (
    <div
      data-slot="tool-fallback-args"
      className={cn('aui-tool-fallback-args px-3', className)}
      {...props}
    >
      <p className="mb-1 text-xs font-semibold text-muted-foreground">Input</p>
      <pre
        className="aui-tool-fallback-args-value whitespace-pre-wrap rounded-[var(--vscode-cornerRadius-medium)] p-2 text-xs"
        style={{
          fontFamily: 'var(--vscode-monospace-font)',
          background: 'var(--vscode-editorWidget-background)',
        }}
      >
        {argsText}
      </pre>
    </div>
  );
}

function ToolFallbackResult({
  result,
  className,
  ...props
}: React.ComponentProps<'div'> & {
  result?: unknown;
}) {
  if (result === undefined) return null;

  return (
    <div
      data-slot="tool-fallback-result"
      className={cn('aui-tool-fallback-result px-3 pt-2', className)}
      style={{ borderTop: '1px dashed var(--vscode-widget-border)' }}
      {...props}
    >
      <p className="aui-tool-fallback-result-header mb-1 text-xs font-semibold text-muted-foreground">
        Output
      </p>
      <pre
        className="aui-tool-fallback-result-content whitespace-pre-wrap rounded-[var(--vscode-cornerRadius-medium)] p-2 text-xs"
        style={{
          fontFamily: 'var(--vscode-monospace-font)',
          background: 'var(--vscode-editorWidget-background)',
        }}
      >
        {typeof result === 'string' ? result : JSON.stringify(result, null, 2)}
      </pre>
    </div>
  );
}

function ToolFallbackError({
  status,
  className,
  ...props
}: React.ComponentProps<'div'> & {
  status?: ToolCallMessagePartStatus;
}) {
  if (status?.type !== 'incomplete') return null;

  const error = status.error;
  const errorText = error ? (typeof error === 'string' ? error : JSON.stringify(error)) : null;

  if (!errorText) return null;

  const isCancelled = status.reason === 'cancelled';
  const headerText = isCancelled ? 'Cancelled reason:' : 'Error:';

  return (
    <div
      data-slot="tool-fallback-error"
      className={cn('aui-tool-fallback-error px-3', className)}
      {...props}
    >
      <p
        className="aui-tool-fallback-error-header font-semibold"
        style={{ color: 'var(--vscode-errorForeground)' }}
      >
        {headerText}
      </p>
      <p
        className="aui-tool-fallback-error-reason"
        style={{ color: 'var(--vscode-errorForeground)' }}
      >
        {errorText}
      </p>
    </div>
  );
}

const ToolFallbackImpl: ToolCallMessagePartComponent = ({ toolName, argsText, result, status }) => {
  const isCancelled = status?.type === 'incomplete' && status.reason === 'cancelled';

  return (
    <ToolFallbackRoot className={cn(isCancelled && 'opacity-60')}>
      <ToolFallbackTrigger toolName={toolName} status={status} result={result as unknown} />
      <ToolFallbackContent>
        <ToolFallbackError status={status} />
        <ToolFallbackArgs argsText={argsText} className={cn(isCancelled && 'opacity-60')} />
        {/* eslint-disable-next-line @typescript-eslint/no-unsafe-assignment */}
        {!isCancelled && <ToolFallbackResult result={result} />}
      </ToolFallbackContent>
    </ToolFallbackRoot>
  );
};

const ToolFallback = memo(ToolFallbackImpl) as unknown as ToolCallMessagePartComponent & {
  Root: typeof ToolFallbackRoot;
  Trigger: typeof ToolFallbackTrigger;
  Content: typeof ToolFallbackContent;
  Args: typeof ToolFallbackArgs;
  Result: typeof ToolFallbackResult;
  Error: typeof ToolFallbackError;
};

ToolFallback.displayName = 'ToolFallback';
ToolFallback.Root = ToolFallbackRoot;
ToolFallback.Trigger = ToolFallbackTrigger;
ToolFallback.Content = ToolFallbackContent;
ToolFallback.Args = ToolFallbackArgs;
ToolFallback.Result = ToolFallbackResult;
ToolFallback.Error = ToolFallbackError;

export {
  ToolFallback,
  ToolFallbackRoot,
  ToolFallbackTrigger,
  ToolFallbackContent,
  ToolFallbackArgs,
  ToolFallbackResult,
  ToolFallbackError,
};
