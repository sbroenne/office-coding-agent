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
import { ArrowDownIcon, CircleStop, LoaderIcon, SparklesIcon } from 'lucide-react';
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
        className="aui-thread-viewport relative flex flex-1 min-h-0 flex-col overflow-x-hidden overflow-y-auto scroll-smooth px-3 pt-3"
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

        <ThreadLevelThinkingIndicator />

        <ThreadPrimitive.ViewportFooter className="aui-thread-viewport-footer sticky bottom-0 mt-auto flex w-full flex-col gap-3 overflow-visible rounded-t-2xl bg-background pb-3">
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

const Composer: FC<{ leftToolbar?: ReactNode; rightToolbar?: ReactNode }> = ({
  leftToolbar,
  rightToolbar,
}) => {
  return (
    <ComposerPrimitive.Root className="aui-composer-root relative flex w-full flex-col rounded-2xl border border-input bg-background px-1 pt-2 outline-none transition-shadow has-[textarea:focus-visible]:border-ring has-[textarea:focus-visible]:ring-2 has-[textarea:focus-visible]:ring-ring/20">
      <ComposerPrimitive.Input
        placeholder="Send a message..."
        className="aui-composer-input max-h-32 min-h-10 w-full resize-none bg-transparent px-3 pt-1 pb-2 text-sm outline-none placeholder:text-muted-foreground focus-visible:ring-0"
        rows={1}
        autoFocus
        aria-label="Message input"
      />
      {/* VS Code-style bottom bar: pickers left, model+send right */}
      <div className="aui-composer-action mx-1 mb-1.5 flex items-center justify-between border-t border-border/40 pt-1">
        <div className="flex items-center gap-0.5 pl-1">{leftToolbar}</div>
        <div className="flex items-center gap-0.5 pr-0.5">
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
                <CircleStop className="size-4" />
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
      <ErrorPrimitive.Root className="aui-message-error-root mt-2 rounded-md border border-destructive bg-destructive/10 p-3 text-destructive text-sm dark:bg-destructive/5 dark:text-red-200">
        <ErrorPrimitive.Message className="aui-message-error-message line-clamp-2" />
      </ErrorPrimitive.Root>
    </MessagePrimitive.Error>
  );
};

/**
 * Thread-level thinking indicator.  Rendered once after ThreadPrimitive.Messages
 * so it always appears below the last message.  It reads directly from
 * ThinkingContext — no dependency on the per-message AUI store, which avoids
 * the timing gap between flushSync renders and the deferred useEffect adapter
 * sync in useExternalStoreRuntime.
 */
const ThreadLevelThinkingIndicator: FC = () => {
  const thinkingText = useThinkingText();
  if (thinkingText === null) return null;
  return (
    <div className="aui-assistant-thinking-indicator fade-in animate-in duration-150 mt-1 flex items-center gap-2 px-4 py-2 text-muted-foreground text-sm">
      <LoaderIcon className="size-3.5 animate-spin" />
      <span className="animate-pulse">{thinkingText}</span>
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
          className="h-6 w-6 rounded text-muted-foreground/60 hover:text-muted-foreground"
        >
          <Codicon name="copy" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.Copy>
      <ActionBarPrimitive.FeedbackPositive asChild>
        <TooltipIconButton
          tooltip="Good response"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:text-muted-foreground"
        >
          <Codicon name="thumbsup" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackPositive>
      <ActionBarPrimitive.FeedbackNegative asChild>
        <TooltipIconButton
          tooltip="Bad response"
          variant="ghost"
          className="h-6 w-6 rounded text-muted-foreground/60 hover:text-muted-foreground"
        >
          <Codicon name="thumbsdown" className="text-sm" />
        </TooltipIconButton>
      </ActionBarPrimitive.FeedbackNegative>
    </ActionBarPrimitive.Root>
  );
};

const AssistantMessage: FC = () => {
  return (
    <MessagePrimitive.Root
      className="aui-assistant-message-root group/message fade-in slide-in-from-bottom-1 relative w-full animate-in py-3 duration-150"
      data-role="assistant"
    >
      <div className="aui-assistant-message-content wrap-break-word px-1 text-foreground text-sm leading-relaxed">
        <MessagePrimitive.Parts
          components={{
            Text: MarkdownText,
            tools: { Fallback: ToolFallback },
            // Prevent EmptyPartFallback from rendering MarkdownText outside a valid
            // part context — this avoids the "MessagePartText can only be used inside
            // text or reasoning message parts" crash when the message has 0 parts or
            // ends in a tool-call / source part.
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
