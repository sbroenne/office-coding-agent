import { MarkdownText } from '@/components/assistant-ui/markdown-text';
import { ToolFallback } from '@/components/assistant-ui/tool-fallback';
import { TooltipIconButton } from '@/components/assistant-ui/tooltip-icon-button';
import {
  ActionBarPrimitive,
  AuiIf,
  BranchPickerPrimitive,
  ComposerPrimitive,
  ErrorPrimitive,
  MessagePrimitive,
  ThreadPrimitive,
  useAuiState,
} from '@assistant-ui/react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { useThinkingText } from '@/contexts/ThinkingContext';

export const Thread: FC<{ leftToolbar?: ReactNode; rightToolbar?: ReactNode }> = ({
  leftToolbar,
  rightToolbar,
}) => {
  return (
    <ThreadPrimitive.Root className="aui-root aui-thread-root flex flex-1 min-h-0 flex-col bg-background">
      <ThreadPrimitive.Viewport
        turnAnchor="bottom"
        className="aui-thread-viewport relative flex flex-1 min-h-0 flex-col overflow-x-hidden overflow-y-auto scroll-smooth pt-2"
      >
        <AuiIf condition={s => s.thread.isEmpty}>
          <ThreadWelcome />
        </AuiIf>

        <ThreadPrimitive.Messages
          components={{
            UserMessage,
            UserEditComposer,
            AssistantMessage,
          }}
        />

        <ThreadLevelThinkingIndicator />

        <ThreadPrimitive.ViewportFooter className="aui-thread-viewport-footer sticky bottom-0 mt-auto flex w-full flex-col gap-2 overflow-visible bg-background px-3 pb-3">
          <ThreadScrollToBottom />
          <Composer leftToolbar={leftToolbar} rightToolbar={rightToolbar} />
        </ThreadPrimitive.ViewportFooter>
      </ThreadPrimitive.Viewport>
    </ThreadPrimitive.Root>
  );
};

const ThreadScrollToBottom: FC = () => {
  return (
    <ThreadPrimitive.ScrollToBottom asChild>
      <TooltipIconButton
        tooltip="Scroll to bottom"
        variant="ghost"
        className="aui-thread-scroll-to-bottom absolute -top-10 z-10 self-center rounded-full bg-transparent p-3 disabled:invisible hover:bg-[var(--vscode-toolbar-hoverBackground)]"
      >
        <Codicon name="chevron-down" className="text-base" />
      </TooltipIconButton>
    </ThreadPrimitive.ScrollToBottom>
  );
};

interface SuggestionItem {
  prompt: string;
  autoSend: boolean;
}

const SUGGESTIONS: SuggestionItem[] = [
  { prompt: 'Summarize my data', autoSend: true },
  { prompt: 'Create a chart from selected data', autoSend: true },
  { prompt: 'Format the table as currency', autoSend: true },
  { prompt: 'Find and highlight duplicates', autoSend: true },
  { prompt: 'Add a formula to calculate totals', autoSend: true },
  { prompt: 'Clean up and organize my sheet', autoSend: true },
];

const ThreadWelcome: FC = () => {
  return (
    <div className="aui-thread-welcome-root my-auto flex w-full grow flex-col px-4">
      <div className="aui-thread-welcome-center flex w-full grow flex-col items-center justify-center">
        <div className="aui-thread-welcome-message flex w-full flex-col items-center justify-center">
          <div
            className="mb-2 flex items-center justify-center rounded-full"
            style={{
              width: 28,
              height: 28,
              background: 'var(--vscode-chat-avatarBackground)',
              color: 'var(--vscode-chat-avatarForeground)',
            }}
          >
            <Codicon name="copilot" className="text-sm" />
          </div>
          <h1 className="aui-thread-welcome-message-inner fade-in slide-in-from-bottom-1 animate-in fill-mode-both font-semibold text-base duration-200">
            How can I help?
          </h1>
        </div>

        <div className="mt-4 flex w-full flex-col">
          {SUGGESTIONS.map((suggestion, idx) => (
            <ThreadPrimitive.Suggestion
              key={suggestion.prompt}
              {...suggestion}
              className="fade-in slide-in-from-bottom-1 animate-in fill-mode-both flex items-center gap-2 rounded-[var(--vscode-cornerRadius-medium)] px-2 py-1.5 text-left text-[12px] transition-colors duration-150 hover:bg-[var(--vscode-list-hoverBackground)]"
              style={{
                animationDelay: `${100 + idx * 50}ms`,
                color: 'var(--vscode-textLink-foreground)',
              }}
            >
              <Codicon name="sparkle" className="shrink-0 text-[11px]" />
              {suggestion.prompt}
            </ThreadPrimitive.Suggestion>
          ))}
        </div>
      </div>
    </div>
  );
};

const Composer: FC<{ leftToolbar?: ReactNode; rightToolbar?: ReactNode }> = ({
  leftToolbar,
  rightToolbar,
}) => {
  return (
    <ComposerPrimitive.Root className="aui-composer-root relative flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] overflow-hidden outline-none transition-colors has-[textarea:focus-visible]:border-[var(--vscode-focusBorder)]">
      <ComposerPrimitive.Input
        placeholder="Send a message..."
        className="aui-composer-input max-h-32 min-h-[34px] w-full resize-none bg-transparent px-3 py-2 text-[13px] outline-none placeholder:text-[var(--vscode-input-placeholderForeground)] focus-visible:outline-none focus-visible:ring-0"
        rows={1}
        autoFocus
        aria-label="Message input"
      />
      <div className="aui-composer-action flex items-center justify-between px-1.5 pb-1">
        <div className="flex items-center gap-0.5">{leftToolbar}</div>
        <div className="flex items-center gap-0.5">
          {rightToolbar}
          <AuiIf condition={s => !s.thread.isRunning}>
            <ComposerPrimitive.Send asChild>
              <TooltipIconButton
                tooltip="Send"
                variant="ghost"
                className="aui-composer-send h-7 w-7 rounded-md transition-opacity"
              >
                <Codicon name="send" className="text-base" />
              </TooltipIconButton>
            </ComposerPrimitive.Send>
          </AuiIf>
          <AuiIf condition={s => s.thread.isRunning}>
            <ComposerPrimitive.Cancel asChild>
              <TooltipIconButton
                tooltip="Stop"
                variant="ghost"
                className="aui-composer-cancel h-7 w-7 rounded-md transition-opacity"
              >
                <Codicon name="debug-stop" className="text-base" />
              </TooltipIconButton>
            </ComposerPrimitive.Cancel>
          </AuiIf>
        </div>
      </div>
    </ComposerPrimitive.Root>
  );
};

const MessageError: FC = () => {
  return (
    <MessagePrimitive.Error>
      <ErrorPrimitive.Root className="aui-message-error-root mt-2 rounded-[var(--vscode-cornerRadius-medium)] border border-[var(--vscode-errorForeground)] bg-[var(--vscode-errorForeground)]/10 p-3 text-[var(--vscode-errorForeground)] text-sm">
        <ErrorPrimitive.Message className="aui-message-error-message line-clamp-2" />
      </ErrorPrimitive.Root>
    </MessagePrimitive.Error>
  );
};

/**
 * Thread-level thinking indicator with VS Code shimmer effect.
 * Rendered once after ThreadPrimitive.Messages so it always appears below the
 * last message. Reads directly from ThinkingContext.
 */
const ThreadLevelThinkingIndicator: FC = () => {
  const thinkingText = useThinkingText();
  if (thinkingText === null) return null;
  return (
    <div className="aui-assistant-message-root" data-role="assistant">
      <div
        className="aui-assistant-thinking-indicator fade-in animate-in duration-150 flex items-center gap-2 px-4 py-1.5"
        style={{ fontSize: 13 }}
      >
        <span className="chat-thinking-shimmer-text">{thinkingText}</span>
      </div>
    </div>
  );
};

const AssistantActionBar: FC = () => {
  return (
    <ActionBarPrimitive.Root
      hideWhenRunning
      autohide="not-last"
      className="aui-assistant-action-bar flex items-center gap-0.5 mt-1 opacity-0 transition-opacity duration-150 group-hover/message:opacity-100 data-[floating]:absolute data-[floating]:-bottom-2 data-[floating]:right-0"
    >
      <ActionBarPrimitive.Copy asChild>
        <TooltipIconButton
          tooltip="Copy"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
        >
          <Codicon name="copy" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.Copy>
      <ActionBarPrimitive.Reload asChild>
        <TooltipIconButton
          tooltip="Regenerate response"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
        >
          <Codicon name="refresh" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.Reload>
      <ActionBarPrimitive.FeedbackPositive asChild>
        <TooltipIconButton
          tooltip="Good response"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
        >
          <Codicon name="thumbsup" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackPositive>
      <ActionBarPrimitive.FeedbackNegative asChild>
        <TooltipIconButton
          tooltip="Bad response"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
        >
          <Codicon name="thumbsdown" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackNegative>
      <BranchPickerPrimitive.Root hideWhenSingleBranch className="flex items-center gap-0.5">
        <BranchPickerPrimitive.Previous asChild>
          <TooltipIconButton
            tooltip="Previous response"
            variant="ghost"
            className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
          >
            <Codicon name="chevron-left" className="text-sm" />
          </TooltipIconButton>
        </BranchPickerPrimitive.Previous>
        <span className="text-xs text-muted-foreground/60 tabular-nums">
          <BranchPickerPrimitive.Number />
          {' / '}
          <BranchPickerPrimitive.Count />
        </span>
        <BranchPickerPrimitive.Next asChild>
          <TooltipIconButton
            tooltip="Next response"
            variant="ghost"
            className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
          >
            <Codicon name="chevron-right" className="text-sm" />
          </TooltipIconButton>
        </BranchPickerPrimitive.Next>
      </BranchPickerPrimitive.Root>
    </ActionBarPrimitive.Root>
  );
};

/**
 * ToolGroup wrapper — rendered by assistant-ui around consecutive tool-call
 * parts. Uses CSS `order: -1` so tool cards appear visually ABOVE the text
 * response (matching VS Code Copilot Chat) while keeping text at DOM index 0
 * to prevent React 18 useSyncExternalStore tearing.
 *
 * When there are 2+ tools, shows a collapsible "Used N tools" header
 * (matching VS Code Copilot Chat's grouped tool invocations). The group
 * auto-collapses once all tools complete and stays expanded while any tool
 * is still running. A single tool renders without the group header.
 */
const ToolGroup: FC<{ startIndex: number; endIndex: number; children?: ReactNode }> = ({
  startIndex,
  endIndex,
  children,
}) => {
  const toolCount = endIndex - startIndex + 1;

  // Read whether any tool in this range is still running
  const isRunning = useAuiState(s => {
    for (let i = startIndex; i <= endIndex; i++) {
      const part = s.message.parts[i];
      if (part?.type === 'tool-call' && part.status?.type === 'running') return true;
    }
    return false;
  });

  // Single tool — no group header, just render the card directly
  if (toolCount <= 1) {
    return (
      <div className="aui-tool-group" style={{ order: -1 }}>
        {children}
      </div>
    );
  }

  return (
    <ToolGroupCollapsible isRunning={isRunning} toolCount={toolCount}>
      {children}
    </ToolGroupCollapsible>
  );
};

/**
 * Inner collapsible component for ToolGroup (2+ tools).
 * Extracted so the hooks are only active when we actually need the collapsible.
 */
const ToolGroupCollapsible: FC<{
  isRunning: boolean;
  toolCount: number;
  children?: ReactNode;
}> = ({ isRunning, toolCount, children }) => {
  const [isExpanded, setIsExpanded] = useState(true);
  const hasAutoCollapsed = useRef(false);

  // Auto-collapse once when every tool in the group completes
  useEffect(() => {
    if (!isRunning && !hasAutoCollapsed.current) {
      hasAutoCollapsed.current = true;
      setIsExpanded(false);
    }
  }, [isRunning]);

  const toggle = useCallback(() => setIsExpanded(prev => !prev), []);

  const headerText = isRunning ? `Running tools\u2026` : `Used ${toolCount} tools`;

  return (
    <div className="aui-tool-group" style={{ order: -1 }}>
      <button
        type="button"
        onClick={toggle}
        className="aui-tool-group-trigger flex w-full items-center gap-1.5 py-1 text-sm text-left"
        aria-expanded={isExpanded}
        data-slot="tool-group-trigger"
      >
        <Codicon
          name={isExpanded ? 'chevron-down' : 'chevron-right'}
          className="text-xs text-muted-foreground shrink-0"
        />
        <Codicon
          name={isRunning ? 'loading~spin' : 'check'}
          className={cn('shrink-0 text-sm', isRunning && 'codicon-modifier-spin')}
        />
        <span className={cn('font-medium', isRunning && 'chat-thinking-shimmer-text')}>
          {headerText}
        </span>
      </button>
      <div className="aui-tool-group-content" style={{ display: isExpanded ? undefined : 'none' }}>
        {children}
      </div>
    </div>
  );
};

const AssistantMessage: FC = () => {
  return (
    <MessagePrimitive.Root
      className="aui-assistant-message-root group/message fade-in slide-in-from-bottom-1 relative w-full animate-in py-2 duration-150"
      data-role="assistant"
    >
      {/* Copilot avatar header */}
      <div className="mb-1.5 flex items-center gap-2 px-4">
        <div
          className="flex items-center justify-center rounded-full"
          style={{
            width: 22,
            height: 22,
            background: 'var(--vscode-chat-avatarBackground)',
            color: 'var(--vscode-chat-avatarForeground)',
          }}
        >
          <Codicon name="copilot" className="text-[12px]" />
        </div>
        <span style={{ fontSize: 13, fontWeight: 600 }} className="text-foreground">
          Copilot
        </span>
      </div>
      <div className="aui-assistant-message-content flex flex-col wrap-break-word px-4 text-foreground text-[13px] leading-[1.5em]">
        <MessagePrimitive.Parts
          components={{
            Text: MarkdownText,
            ToolGroup,
            tools: { Fallback: ToolFallback },
            Empty: () => null,
          }}
          unstable_showEmptyOnNonTextEnd={false}
        />
        <AssistantActionBar />
        <MessageError />
      </div>
    </MessagePrimitive.Root>
  );
};

const UserActionBar: FC = () => {
  return (
    <ActionBarPrimitive.Root
      hideWhenRunning
      autohide="always"
      className="aui-user-action-bar absolute right-2 top-2 flex items-center opacity-0 transition-opacity duration-150 group-hover/message:opacity-100"
    >
      <ActionBarPrimitive.Edit asChild>
        <TooltipIconButton
          tooltip="Edit message"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground"
        >
          <Codicon name="edit" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.Edit>
    </ActionBarPrimitive.Root>
  );
};

const UserEditComposer: FC = () => {
  return (
    <ComposerPrimitive.Root className="flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] px-1 pt-2 outline-none transition-colors has-[textarea:focus-visible]:border-[var(--vscode-focusBorder)]">
      <ComposerPrimitive.Input
        className="max-h-32 min-h-10 w-full resize-none bg-transparent px-3 pt-1 pb-2 text-sm outline-none placeholder:text-[var(--vscode-input-placeholderForeground)] focus-visible:ring-0"
        rows={1}
      />
      <div className="mx-1 mb-1.5 flex items-center justify-end gap-1 border-t border-border/40 pt-1">
        <ComposerPrimitive.Cancel asChild>
          <TooltipIconButton
            tooltip="Cancel edit"
            variant="ghost"
            className="h-7 px-2 rounded-md text-xs text-muted-foreground hover:text-foreground"
          >
            Cancel
          </TooltipIconButton>
        </ComposerPrimitive.Cancel>
        <ComposerPrimitive.Send asChild>
          <TooltipIconButton tooltip="Send edit" variant="ghost" className="h-7 w-7 rounded-md">
            <Codicon name="send" className="text-base" />
          </TooltipIconButton>
        </ComposerPrimitive.Send>
      </div>
    </ComposerPrimitive.Root>
  );
};

const UserMessage: FC = () => {
  return (
    <MessagePrimitive.Root
      className="aui-user-message-root fade-in slide-in-from-bottom-1 group/message relative flex w-full animate-in px-4 py-2 duration-150"
      data-role="user"
    >
      <div className="aui-user-message-content wrap-break-word w-full text-foreground text-[13px] leading-[1.5em]">
        <MessagePrimitive.Content />
      </div>
      <UserActionBar />
    </MessagePrimitive.Root>
  );
};
