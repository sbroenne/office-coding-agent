import { MarkdownText } from '@/components/assistant-ui/markdown-text';
import { ToolFallback } from '@/components/assistant-ui/tool-fallback';
import { TooltipIconButton } from '@/components/assistant-ui/tooltip-icon-button';
import {
  ActionBarPrimitive,
  AuiIf,
  ComposerPrimitive,
  ErrorPrimitive,
  MessagePrimitive,
  ThreadPrimitive,
} from '@assistant-ui/react';
import {
  ArrowDownIcon,
  ArrowUpIcon,
  CheckIcon,
  CopyIcon,
  LoaderIcon,
  RefreshCwIcon,
  SparklesIcon,
  SquareIcon,
} from 'lucide-react';
import type { FC } from 'react';

export const Thread: FC = () => {
  return (
    <ThreadPrimitive.Root className="aui-root aui-thread-root flex h-full flex-col bg-background">
      <ThreadPrimitive.Viewport
        turnAnchor="top"
        className="aui-thread-viewport relative flex flex-1 flex-col overflow-x-hidden overflow-y-auto scroll-smooth px-3 pt-3"
      >
        <AuiIf condition={s => s.thread.isEmpty}>
          <ThreadWelcome />
        </AuiIf>

        <ThreadPrimitive.Messages
          components={{
            UserMessage,
            AssistantMessage,
          }}
        />

        <ThreadPrimitive.ViewportFooter className="aui-thread-viewport-footer sticky bottom-0 mt-auto flex w-full flex-col gap-3 overflow-visible rounded-t-2xl bg-background pb-3">
          <ThreadScrollToBottom />
          <Composer />
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
        variant="outline"
        className="aui-thread-scroll-to-bottom absolute -top-10 z-10 self-center rounded-full p-3 disabled:invisible dark:bg-background dark:hover:bg-accent"
      >
        <ArrowDownIcon />
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
    <div className="aui-thread-welcome-root my-auto flex w-full grow flex-col">
      <div className="aui-thread-welcome-center flex w-full grow flex-col items-center justify-center">
        <div className="aui-thread-welcome-message flex w-full flex-col justify-center px-2">
          <h1 className="aui-thread-welcome-message-inner fade-in slide-in-from-bottom-1 animate-in fill-mode-both font-semibold text-xl duration-200">
            Hello there!
          </h1>
          <p className="aui-thread-welcome-message-inner fade-in slide-in-from-bottom-1 animate-in fill-mode-both text-muted-foreground text-base delay-75 duration-200">
            How can I help you today?
          </p>
        </div>

        <div className="mt-6 flex w-full flex-col gap-2 px-2">
          {SUGGESTIONS.map((suggestion, idx) => (
            <ThreadPrimitive.Suggestion
              key={suggestion.prompt}
              {...suggestion}
              className="fade-in slide-in-from-bottom-1 animate-in fill-mode-both flex items-center gap-2 rounded-xl border border-border bg-background px-3 py-2.5 text-left text-sm text-foreground shadow-sm transition-colors duration-150 hover:bg-accent hover:shadow-md"
              style={{ animationDelay: `${100 + idx * 50}ms` }}
            >
              <SparklesIcon className="size-3.5 shrink-0 text-primary" />
              {suggestion.prompt}
            </ThreadPrimitive.Suggestion>
          ))}
        </div>
      </div>
    </div>
  );
};

const Composer: FC = () => {
  return (
    <ComposerPrimitive.Root className="aui-composer-root relative flex w-full flex-col rounded-2xl border border-input bg-background px-1 pt-2 outline-none transition-shadow has-[textarea:focus-visible]:border-ring has-[textarea:focus-visible]:ring-2 has-[textarea:focus-visible]:ring-ring/20">
      <ComposerPrimitive.Input
        placeholder="Send a message..."
        className="aui-composer-input mb-1 max-h-32 min-h-10 w-full resize-none bg-transparent px-3 pt-1 pb-2 text-sm outline-none placeholder:text-muted-foreground focus-visible:ring-0"
        rows={1}
        autoFocus
        aria-label="Message input"
      />
      <ComposerAction />
    </ComposerPrimitive.Root>
  );
};

const ComposerAction: FC = () => {
  return (
    <div className="aui-composer-action mx-2 mb-2 flex items-center justify-end">
      <AuiIf condition={s => !s.thread.isRunning}>
        <ComposerPrimitive.Send asChild>
          <TooltipIconButton
            tooltip="Send"
            variant="default"
            className="aui-composer-send size-8 rounded-full p-2 transition-opacity"
          >
            <ArrowUpIcon />
          </TooltipIconButton>
        </ComposerPrimitive.Send>
      </AuiIf>
      <AuiIf condition={s => s.thread.isRunning}>
        <ComposerPrimitive.Cancel asChild>
          <TooltipIconButton
            tooltip="Cancel"
            variant="default"
            className="aui-composer-cancel size-8 rounded-full p-2 transition-opacity"
          >
            <SquareIcon className="size-4" />
          </TooltipIconButton>
        </ComposerPrimitive.Cancel>
      </AuiIf>
    </div>
  );
};

const MessageError: FC = () => {
  return (
    <MessagePrimitive.Error>
      <ErrorPrimitive.Root className="aui-message-error-root mt-2 rounded-md border border-destructive bg-destructive/10 p-3 text-destructive text-sm dark:bg-destructive/5 dark:text-red-200">
        <ErrorPrimitive.Message className="aui-message-error-message line-clamp-2" />
      </ErrorPrimitive.Root>
    </MessagePrimitive.Error>
  );
};

const ThinkingIndicator: FC = () => {
  return (
    <AuiIf condition={s => s.message.isLast && s.thread.isRunning}>
      <div className="aui-thinking-indicator flex items-center gap-2 px-1 py-1 text-muted-foreground text-sm">
        <LoaderIcon className="size-3.5 animate-spin" />
        <span className="animate-pulse">Thinking…</span>
      </div>
    </AuiIf>
  );
};

const AssistantMessage: FC = () => {
  return (
    <MessagePrimitive.Root
      className="aui-assistant-message-root fade-in slide-in-from-bottom-1 relative w-full animate-in py-3 duration-150"
      data-role="assistant"
    >
      <div className="aui-assistant-message-content wrap-break-word px-1 text-foreground text-sm leading-relaxed">
        <MessagePrimitive.Parts
          components={{
            Text: MarkdownText,
            tools: { Fallback: ToolFallback },
          }}
        />
        <ThinkingIndicator />
        <MessageError />
      </div>

      <div className="aui-assistant-message-footer mt-1 ml-1 flex">
        <AssistantActionBar />
      </div>
    </MessagePrimitive.Root>
  );
};

const AssistantActionBar: FC = () => {
  return (
    <ActionBarPrimitive.Root
      hideWhenRunning
      autohide="not-last"
      autohideFloat="single-branch"
      className="aui-assistant-action-bar-root -ml-1 flex gap-1 text-muted-foreground data-floating:absolute data-floating:rounded-md data-floating:border data-floating:bg-background data-floating:p-1 data-floating:shadow-sm"
    >
      <ActionBarPrimitive.Copy asChild>
        <TooltipIconButton tooltip="Copy">
          <AuiIf condition={s => s.message.isCopied}>
            <CheckIcon />
          </AuiIf>
          <AuiIf condition={s => !s.message.isCopied}>
            <CopyIcon />
          </AuiIf>
        </TooltipIconButton>
      </ActionBarPrimitive.Copy>
      <ActionBarPrimitive.Reload asChild>
        <TooltipIconButton tooltip="Regenerate">
          <RefreshCwIcon />
        </TooltipIconButton>
      </ActionBarPrimitive.Reload>
    </ActionBarPrimitive.Root>
  );
};

const UserMessage: FC = () => {
  return (
    <MessagePrimitive.Root
      className="aui-user-message-root fade-in slide-in-from-bottom-1 flex w-full animate-in justify-end py-3 duration-150"
      data-role="user"
    >
      <div className="aui-user-message-content wrap-break-word max-w-[85%] rounded-2xl bg-muted px-3 py-2 text-foreground text-sm">
        <MessagePrimitive.Content />
      </div>
    </MessagePrimitive.Root>
  );
};
