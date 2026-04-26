import React, { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { getLocalApiBase } from '@/lib/api';

interface SlashItem {
  type: 'skill' | 'prompt';
  name: string;
  description?: string;
  plugin?: string;
  source?: string;
}

interface SlashItemsResponse {
  skills?: SlashItem[];
  prompts?: SlashItem[];
}

interface SlashSuggestion {
  type: 'skill' | 'prompt';
  value: string;
  name: string;
  description?: string;
  plugin?: string;
  source?: string;
}

interface ChatComposerProps {
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onEnqueue?: (text: string) => void;
  isRunning: boolean;
  queuedCount?: number;
  /** Previous user messages for Up/Down arrow history navigation. */
  history?: string[];
  placeholder?: string;
  leftToolbar?: ReactNode;
  rightToolbar?: ReactNode;
  autoFocus?: boolean;
}

export const ChatComposer: FC<ChatComposerProps> = ({
  onSend,
  onCancel,
  onEnqueue,
  isRunning,
  queuedCount = 0,
  history = [],
  placeholder = 'Send a message...',
  leftToolbar,
  rightToolbar,
  autoFocus = true,
}) => {
  const [text, setText] = useState('');
  const [slashItems, setSlashItems] = useState<SlashItemsResponse>({});
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // History navigation state: -1 = composing new text, 0 = most recent, etc.
  const historyIndexRef = useRef(-1);
  // Stash the in-progress draft when the user starts navigating history
  const draftRef = useRef('');

  // Auto-resize textarea as content changes
  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${el.scrollHeight}px`;
  }, [text]);

  useEffect(() => {
    void fetch(`${getLocalApiBase()}/api/slash-items`)
      .then(res => (res.ok ? res.json() : {}))
      .then((data: SlashItemsResponse) => setSlashItems(data))
      .catch(() => setSlashItems({}));
  }, []);

  const rootSlashMatch = /^\/([\w-]*)$/i.exec(text);
  const slashQuery = rootSlashMatch?.[1]?.toLowerCase() ?? '';
  const slashMode = rootSlashMatch ? 'slash' : undefined;
  const skillSuggestions = (slashItems.skills ?? [])
    .filter(item => {
      if (!slashQuery) return true;
      return [item.name, item.description, item.plugin].some(
        value => value?.toLowerCase().includes(slashQuery) ?? false
      );
    })
    .map(
      (item): SlashSuggestion => ({
        type: 'skill',
        value: `/${item.name} `,
        name: `/${item.name}`,
        description: item.description,
        plugin: item.plugin,
      })
    );
  const promptSuggestions = (slashItems.prompts ?? [])
    .filter(item => {
      if (!slashQuery) return true;
      return [item.name, item.description, item.source].some(
        value => value?.toLowerCase().includes(slashQuery) ?? false
      );
    })
    .map(
      (item): SlashSuggestion => ({
        type: 'prompt',
        value: `/${item.name} `,
        name: `/${item.name}`,
        description: item.description,
        source: item.source,
      })
    );
  const visibleSlashSuggestions = [...skillSuggestions, ...promptSuggestions].slice(0, 8);

  const applySlashSuggestion = useCallback((suggestion: SlashSuggestion) => {
    setText(suggestion.value);
    historyIndexRef.current = -1;
    draftRef.current = '';
    requestAnimationFrame(() => {
      textareaRef.current?.focus();
      textareaRef.current?.setSelectionRange(suggestion.value.length, suggestion.value.length);
    });
  }, []);

  const handleTextChange = useCallback((e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(e.target.value);
  }, []);

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    setText('');
    historyIndexRef.current = -1;
    draftRef.current = '';
    void onSend(trimmed);
  }, [text, onSend]);

  const handleEnqueue = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    if (!isRunning) {
      setText('');
      historyIndexRef.current = -1;
      draftRef.current = '';
      void onSend(trimmed);
      return;
    }
    setText('');
    historyIndexRef.current = -1;
    draftRef.current = '';
    onEnqueue?.(trimmed);
  }, [text, isRunning, onSend, onEnqueue]);

  /** Navigate through history (reversed: index 0 = most recent). */
  const navigateHistory = useCallback(
    (direction: 'up' | 'down') => {
      if (history.length === 0) return;

      const idx = historyIndexRef.current;

      if (direction === 'up') {
        if (idx === -1) {
          // Save current draft before navigating
          draftRef.current = text;
        }
        const nextIdx = Math.min(idx + 1, history.length - 1);
        if (nextIdx === idx && idx !== -1) return; // already at oldest
        historyIndexRef.current = nextIdx;
        setText(history[nextIdx]);
      } else {
        // down
        if (idx <= -1) return; // already at draft
        const nextIdx = idx - 1;
        historyIndexRef.current = nextIdx;
        if (nextIdx === -1) {
          setText(draftRef.current);
        } else {
          setText(history[nextIdx]);
        }
      }
    },
    [history, text]
  );

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        handleSend();
      } else if (e.key === 'q' && e.shiftKey && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        handleEnqueue();
      } else if (e.key === 'ArrowUp') {
        // Only navigate history when cursor is at start of input (single-line behavior)
        const el = textareaRef.current;
        if (el?.selectionStart === 0 && el.selectionEnd === 0) {
          e.preventDefault();
          navigateHistory('up');
        }
      } else if (e.key === 'ArrowDown') {
        // Only navigate history when cursor is at end of input
        const el = textareaRef.current;
        if (el && el.selectionStart === el.value.length && el.selectionEnd === el.value.length) {
          e.preventDefault();
          navigateHistory('down');
        }
      }
    },
    [handleSend, handleEnqueue, navigateHistory]
  );

  // Dynamic placeholder: hint at queue shortcut when agent is running
  const activePlaceholder = isRunning ? 'Type here, Ctrl+Shift+Q to queue...' : placeholder;

  return (
    <div className="aui-composer-root relative flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] overflow-hidden outline-none transition-colors focus-within:border-[var(--vscode-focusBorder)]">
      {slashMode && (
        <div
          className="max-h-56 overflow-y-auto border-b border-border p-1"
          role="listbox"
          aria-label="slash suggestions"
        >
          {visibleSlashSuggestions.length > 0 ? (
            visibleSlashSuggestions.map(suggestion => (
              <button
                key={`${suggestion.type}:${suggestion.plugin ?? suggestion.source ?? ''}:${suggestion.name}`}
                type="button"
                onClick={() => applySlashSuggestion(suggestion)}
                className="flex w-full items-start gap-2 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1.5 text-left text-sm hover:bg-accent"
                role="option"
              >
                <Codicon
                  name={suggestion.type === 'skill' ? 'lightbulb-sparkle' : 'symbol-keyword'}
                  className="mt-0.5 text-[14px] text-[var(--vscode-icon-foreground)]"
                />
                <span className="min-w-0 flex-1">
                  <span className="block truncate font-medium text-foreground">
                    {suggestion.name}
                  </span>
                  {[suggestion.description, suggestion.plugin ?? suggestion.source].some(
                    Boolean
                  ) && (
                    <span className="block truncate text-xs text-muted-foreground">
                      {[suggestion.description, suggestion.plugin ?? suggestion.source]
                        .filter(Boolean)
                        .join(' — ')}
                    </span>
                  )}
                </span>
              </button>
            ))
          ) : (
            <div className="px-2 py-2 text-xs text-muted-foreground">
              No matching slash commands found.
            </div>
          )}
        </div>
      )}
      <textarea
        ref={textareaRef}
        value={text}
        onChange={handleTextChange}
        onKeyDown={handleKeyDown}
        placeholder={activePlaceholder}
        rows={1}
        autoFocus={autoFocus}
        aria-label="Message input"
        className={cn(
          'aui-composer-input max-h-32 min-h-[34px] w-full resize-none bg-transparent px-3 py-2',
          'text-[13px] outline-none placeholder:text-[var(--vscode-input-placeholderForeground)]',
          'focus-visible:outline-none focus-visible:ring-0'
        )}
        style={{ overflow: 'hidden' }}
      />
      <div className="aui-composer-action flex items-center justify-between px-1.5 pb-1">
        <div className="flex items-center gap-0.5">
          {leftToolbar}
          {/* Queue badge — shows number of enqueued prompts */}
          {queuedCount > 0 && (
            <span
              className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1.5 text-[11px] leading-[18px]"
              style={{
                color: 'var(--vscode-badge-foreground)',
                background: 'var(--vscode-badge-background)',
              }}
              title={`${queuedCount} prompt${queuedCount !== 1 ? 's' : ''} queued (Ctrl+Shift+Q)`}
              data-testid="queue-badge"
            >
              <Codicon name="layers" className="text-[10px]" />
              {queuedCount}
            </span>
          )}
        </div>
        <div className="flex items-center gap-0.5">
          {rightToolbar}

          {/* Shortcut hint — shown when running, no text, and no queued items */}
          {isRunning && !text.trim() && queuedCount === 0 && (
            <span
              className="text-[11px] px-1"
              style={{ color: 'var(--vscode-descriptionForeground)' }}
            >
              <kbd
                className="rounded px-1 py-0.5 text-[10px]"
                style={{
                  background: 'var(--vscode-keybindingLabel-background)',
                  border:
                    '1px solid var(--vscode-keybindingLabel-border, var(--vscode-widget-border))',
                  color: 'var(--vscode-keybindingLabel-foreground)',
                }}
              >
                Ctrl+Shift+Q
              </kbd>{' '}
              to queue
            </span>
          )}

          {/* Enqueue button — shown when running and text is entered */}
          {isRunning && text.trim() && onEnqueue && (
            <button
              type="button"
              onClick={handleEnqueue}
              title="Queue prompt (Ctrl+Shift+Q)"
              className="aui-composer-enqueue h-7 rounded-md flex items-center justify-center gap-1 px-1.5 transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
              style={{ color: 'var(--vscode-textLink-foreground)', fontSize: 11 }}
              data-testid="enqueue-button"
            >
              <Codicon name="add" className="text-xs" />
              <span>Queue</span>
            </button>
          )}

          {/* Send button — always visible and functional */}
          <button
            type="button"
            onClick={handleSend}
            title="Send"
            className={cn(
              'aui-composer-send h-7 w-7 rounded-md flex items-center justify-center transition-colors',
              'hover:bg-[var(--vscode-toolbar-hoverBackground)]',
              !text.trim() && 'opacity-40'
            )}
            style={{ color: 'var(--vscode-icon-foreground)' }}
          >
            <Codicon name="send" className="text-base" />
          </button>

          {/* Stop button — shown when running */}
          {isRunning && (
            <button
              type="button"
              onClick={onCancel}
              title="Stop"
              className="aui-composer-cancel h-7 w-7 rounded-md flex items-center justify-center transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
              style={{ color: 'var(--vscode-icon-foreground)' }}
            >
              <Codicon name="debug-stop" className="text-base" />
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
