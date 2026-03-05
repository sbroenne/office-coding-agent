import { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { UserMessage } from './UserMessage';
import { AssistantMessage } from './AssistantMessage';
import { ChatComposer } from './ChatComposer';
import { useChatActions } from '@/contexts/ChatActionsContext';
import { detectOfficeHost } from '@/services/office/host';
import type { ChatMessage } from '@/types';

// ─── Welcome Screen ───────────────────────────────────────────────────────────

interface SuggestionItem {
  prompt: string;
}

const SUGGESTIONS_BY_HOST: Record<string, SuggestionItem[]> = {
  excel: [
    { prompt: 'Summarize my data' },
    { prompt: 'Create a chart from selected data' },
    { prompt: 'Format the table as currency' },
    { prompt: 'Find and highlight duplicates' },
    { prompt: 'Add a formula to calculate totals' },
    { prompt: 'Clean up and organize my sheet' },
  ],
  powerpoint: [
    { prompt: 'Create a title slide for my presentation' },
    { prompt: 'Add a slide with 3 key points' },
    { prompt: 'Create a chart slide from data' },
    { prompt: 'Redesign this slide with better layout' },
    { prompt: 'Add speaker notes to all slides' },
    { prompt: 'Create a 5-slide presentation about...' },
  ],
  word: [
    { prompt: 'Summarize this document' },
    { prompt: 'Fix grammar and improve clarity' },
    { prompt: 'Create a table of contents' },
    { prompt: 'Translate this paragraph to...' },
    { prompt: 'Write an executive summary' },
    { prompt: 'Format headings and structure' },
  ],
  outlook: [
    { prompt: 'Draft a reply to this email' },
    { prompt: 'Summarize this email thread' },
    { prompt: 'Make this email more professional' },
    { prompt: 'Write a follow-up email' },
  ],
};

const DEFAULT_SUGGESTIONS: SuggestionItem[] = [
  { prompt: 'What can you help me with?' },
  { prompt: 'Tell me about your capabilities' },
];

const ThreadWelcome: FC = () => {
  const { send } = useChatActions();
  const host = detectOfficeHost();
  const suggestions = SUGGESTIONS_BY_HOST[host] ?? DEFAULT_SUGGESTIONS;

  return (
    <div className="aui-thread-welcome-root my-auto flex w-full grow flex-col px-4">
      <div className="aui-thread-welcome-center flex w-full grow flex-col items-center justify-center">
        <div className="aui-thread-welcome-message flex w-full flex-col items-center justify-center">
          <div style={{ marginBottom: 24 }}>
            <Codicon
              name="copilot"
              className="text-[40px] text-[var(--vscode-descriptionForeground)]"
            />
          </div>
          <h1
            className="aui-thread-welcome-message-inner fade-in slide-in-from-bottom-1 animate-in fill-mode-both font-semibold duration-200"
            style={{ fontSize: 13 }}
          >
            How can I help?
          </h1>
        </div>

        <div className="chat-welcome-suggested-prompts mt-4">
          <p className="chat-welcome-suggested-prompts-title">Try asking</p>
          {suggestions.map((suggestion, idx) => (
            <button
              key={suggestion.prompt}
              onClick={() => {
                void send(suggestion.prompt);
              }}
              className="chat-welcome-suggested-prompt fade-in animate-in fill-mode-both duration-150"
              style={{ animationDelay: `${100 + idx * 50}ms` }}
            >
              {suggestion.prompt}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── Scroll-to-bottom button ──────────────────────────────────────────────────

const ScrollToBottomButton: FC<{ onClick: () => void }> = ({ onClick }) => (
  <button
    type="button"
    onClick={onClick}
    title="Scroll to bottom"
    className={cn(
      'absolute -top-10 z-10 self-center rounded-full bg-transparent p-3',
      'hover:bg-[var(--vscode-toolbar-hoverBackground)] transition-colors'
    )}
    style={{ color: 'var(--vscode-icon-foreground)' }}
  >
    <Codicon name="chevron-down" className="text-base" />
  </button>
);

// ─── MessageList ──────────────────────────────────────────────────────────────

interface MessageListProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onEdit?: (messageId: string, newText: string) => void;
  onRegenerate?: (messageId: string) => void;
  onFeedback?: (messageId: string, kind: 'positive' | 'negative') => void;
  leftToolbar?: ReactNode;
  rightToolbar?: ReactNode;
}

export const MessageList: FC<MessageListProps> = ({
  messages,
  isRunning,
  onSend,
  onCancel,
  onEdit,
  onRegenerate,
  onFeedback,
  leftToolbar,
  rightToolbar,
}) => {
  const viewportRef = useRef<HTMLDivElement>(null);
  const bottomRef = useRef<HTMLDivElement>(null);
  const [showScrollToBottom, setShowScrollToBottom] = useState(false);
  // Track whether the user has scrolled up (broken auto-scroll)
  const userScrolledUp = useRef(false);

  const scrollToBottom = useCallback((behavior: ScrollBehavior = 'smooth') => {
    bottomRef.current?.scrollIntoView({ behavior, block: 'end' });
    userScrolledUp.current = false;
    setShowScrollToBottom(false);
  }, []);

  // Auto-scroll when messages change, unless user has scrolled up
  useEffect(() => {
    if (!userScrolledUp.current) {
      scrollToBottom('smooth');
    }
  }, [messages, scrollToBottom]);

  // Detect when user scrolls up
  const handleScroll = useCallback(() => {
    const el = viewportRef.current;
    if (!el) return;
    const distanceFromBottom = el.scrollHeight - el.scrollTop - el.clientHeight;
    const isAtBottom = distanceFromBottom < 40;
    userScrolledUp.current = !isAtBottom;
    setShowScrollToBottom(!isAtBottom);
  }, []);

  const isEmpty = messages.length === 0;

  return (
    <div className="aui-root aui-thread-root flex flex-1 min-h-0 flex-col bg-background">
      <div
        ref={viewportRef}
        onScroll={handleScroll}
        className="aui-thread-viewport relative flex flex-1 min-h-0 flex-col overflow-x-hidden overflow-y-auto scroll-smooth pt-2"
      >
        {isEmpty && <ThreadWelcome />}

        {messages.map((msg, idx) => {
          const isLast = idx === messages.length - 1;
          if (msg.role === 'user') {
            return <UserMessage key={msg.id} message={msg} onEdit={onEdit} />;
          }
          return (
            <AssistantMessage
              key={msg.id}
              message={msg}
              isRunning={isRunning}
              isLast={isLast}
              onRegenerate={onRegenerate}
              onFeedback={onFeedback}
            />
          );
        })}

        <div ref={bottomRef} style={{ height: 1 }} />

        {/* Sticky footer: scroll-to-bottom + composer */}
        <div className="aui-thread-viewport-footer sticky bottom-0 mt-auto flex w-full flex-col gap-2 overflow-visible bg-background px-3 pb-3">
          {showScrollToBottom && <ScrollToBottomButton onClick={() => scrollToBottom('smooth')} />}
          <ChatComposer
            onSend={onSend}
            onCancel={onCancel}
            isRunning={isRunning}
            leftToolbar={leftToolbar}
            rightToolbar={rightToolbar}
          />
        </div>
      </div>
    </div>
  );
};
