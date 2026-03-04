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
} from '@assistant-ui/react';
import { Codicon } from '@/components/Codicon';
import type { FC, ReactNode } from 'react';
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
        className="aui-thread-scroll-to-bottom absolute -top-8 z-10 self-center rounded-[var(--vscode-cornerRadius-small)] bg-[var(--vscode-editorWidget-background)] border border-border disabled:invisible"
        style={{ width: 26, height: 26, color: 'var(--vscode-icon-foreground)' }}
      >
        <Codicon name="chevron-down" className="text-[14px]" />
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
            <Codicon name="copilot" className="text-[12px]" />
          </div>
          <h1 className="aui-thread-welcome-message-inner fade-in slide-in-from-bottom-1 animate-in fill-mode-both font-semibold text-[16px] duration-200">
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
      <div className="aui-composer-action flex items-center justify-between px-1 pb-1">
        <div className="flex items-center gap-1">{leftToolbar}</div>
        <div className="flex items-center gap-1">
          {rightToolbar}
          <AuiIf condition={s => !s.thread.isRunning}>
            <ComposerPrimitive.Send asChild>
              <TooltipIconButton
                tooltip="Send"
                variant="ghost"
                className="aui-composer-send rounded-[var(--vscode-cornerRadius-small)] transition-opacity"
                style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
              >
                <Codicon name="send" className="text-[14px]" />
              </TooltipIconButton>
            </ComposerPrimitive.Send>
          </AuiIf>
          <AuiIf condition={s => s.thread.isRunning}>
            <ComposerPrimitive.Cancel asChild>
              <TooltipIconButton
                tooltip="Stop"
                variant="ghost"
                className="aui-composer-cancel rounded-[var(--vscode-cornerRadius-small)] transition-opacity"
                style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
              >
                <Codicon name="debug-stop" className="text-[14px]" />
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
          className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
          style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
        >
          <Codicon name="copy" className="text-[12px]" />
        </TooltipIconButton>
      </ActionBarPrimitive.Copy>
      <ActionBarPrimitive.Reload asChild>
        <TooltipIconButton
          tooltip="Regenerate response"
          variant="ghost"
          className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
          style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
        >
          <Codicon name="refresh" className="text-[12px]" />
        </TooltipIconButton>
      </ActionBarPrimitive.Reload>
      <ActionBarPrimitive.FeedbackPositive asChild>
        <TooltipIconButton
          tooltip="Good response"
          variant="ghost"
          className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
          style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
        >
          <Codicon name="thumbsup" className="text-[12px]" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackPositive>
      <ActionBarPrimitive.FeedbackNegative asChild>
        <TooltipIconButton
          tooltip="Bad response"
          variant="ghost"
          className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
          style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
        >
          <Codicon name="thumbsdown" className="text-[12px]" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackNegative>
      <BranchPickerPrimitive.Root hideWhenSingleBranch className="flex items-center gap-0.5">
        <BranchPickerPrimitive.Previous asChild>
          <TooltipIconButton
            tooltip="Previous response"
            variant="ghost"
            className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
            style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
          >
            <Codicon name="chevron-left" className="text-[12px]" />
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
            className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
            style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
          >
            <Codicon name="chevron-right" className="text-[12px]" />
          </TooltipIconButton>
        </BranchPickerPrimitive.Next>
      </BranchPickerPrimitive.Root>
    </ActionBarPrimitive.Root>
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
      <div className="aui-assistant-message-content wrap-break-word px-4 text-foreground text-[13px] leading-[1.5em]">
        <MessagePrimitive.Parts
          components={{
            Text: MarkdownText,
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
          className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
          style={{ width: 20, height: 20, color: 'var(--vscode-icon-foreground)' }}
        >
          <Codicon name="edit" className="text-[12px]" />
        </TooltipIconButton>
      </ActionBarPrimitive.Edit>
    </ActionBarPrimitive.Root>
  );
};

const UserEditComposer: FC = () => {
  return (
    <ComposerPrimitive.Root className="flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] px-1 pt-2 outline-none transition-colors has-[textarea:focus-visible]:border-[var(--vscode-focusBorder)]">
      <ComposerPrimitive.Input
        className="max-h-32 min-h-[34px] w-full resize-none bg-transparent px-3 py-2 text-[13px] outline-none placeholder:text-[var(--vscode-input-placeholderForeground)] focus-visible:outline-none focus-visible:ring-0"
        rows={1}
      />
      <div className="mx-1 mb-1 flex items-center justify-end gap-1">
        <ComposerPrimitive.Cancel asChild>
          <TooltipIconButton
            tooltip="Cancel edit"
            variant="ghost"
            className="rounded-[var(--vscode-cornerRadius-small)] px-2 text-[12px] hover:bg-accent"
            style={{ height: 22, color: 'var(--vscode-icon-foreground)' }}
          >
            Cancel
          </TooltipIconButton>
        </ComposerPrimitive.Cancel>
        <ComposerPrimitive.Send asChild>
          <TooltipIconButton
            tooltip="Send edit"
            variant="ghost"
            className="rounded-[var(--vscode-cornerRadius-small)] hover:bg-accent"
            style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
          >
            <Codicon name="send" className="text-[14px]" />
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
