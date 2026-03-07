import React, { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import type { PluginPrompt } from '@/types/plugin';

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
  /** Plugin prompt templates surfaced by the `/` slash command menu. */
  slashCommands?: PluginPrompt[];
  /** Called when the slash menu selects a prompt with an associated agent name. */
  onAgentSelect?: (agentName: string) => void;
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
  slashCommands = [],
  onAgentSelect,
}) => {
  const [text, setText] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // History navigation state: -1 = composing new text, 0 = most recent, etc.
  const historyIndexRef = useRef(-1);
  // Stash the in-progress draft when the user starts navigating history
  const draftRef = useRef('');

  // ─── Slash command menu state ────────────────────────────────────────────────
  const [slashQuery, setSlashQuery] = useState<string | null>(null);
  const [slashIndex, setSlashIndex] = useState(0);
  const slashMenuRef = useRef<HTMLDivElement>(null);

  const filteredCommands =
    slashQuery !== null
      ? slashCommands.filter(
          cmd =>
            slashQuery === '' ||
            cmd.name.toLowerCase().includes(slashQuery.toLowerCase()) ||
            cmd.description.toLowerCase().includes(slashQuery.toLowerCase())
        )
      : [];

  // Auto-resize textarea as content changes
  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${el.scrollHeight}px`;
  }, [text]);

  // Keep slash menu index in bounds when filter changes
  useEffect(() => {
    setSlashIndex(0);
  }, [slashQuery]);

  // Close slash menu when clicking outside
  useEffect(() => {
    if (slashQuery === null) return;
    const handler = (e: MouseEvent) => {
      if (
        slashMenuRef.current &&
        !slashMenuRef.current.contains(e.target as Node) &&
        !textareaRef.current?.contains(e.target as Node)
      ) {
        setSlashQuery(null);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [slashQuery]);

  const handleTextChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      const val = e.target.value;
      setText(val);

      // Open slash menu when first character is `/` and slash commands exist
      if (slashCommands.length > 0 && val.startsWith('/')) {
        setSlashQuery(val.slice(1));
      } else {
        setSlashQuery(null);
      }
    },
    [slashCommands]
  );

  /** Select a slash command: fill the textarea, optionally switch agent. */
  const selectSlashCommand = useCallback(
    (cmd: PluginPrompt) => {
      // Replace ${input:varname} placeholders with <varname> for clarity
      const filled = cmd.body.replace(/\$\{input:([^}]+)\}/g, '<$1>');
      setText(filled);
      setSlashQuery(null);
      if (cmd.agent) {
        onAgentSelect?.(cmd.agent);
      }
      setTimeout(() => textareaRef.current?.focus(), 0);
    },
    [onAgentSelect]
  );

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    setText('');
    setSlashQuery(null);
    historyIndexRef.current = -1;
    draftRef.current = '';
    void onSend(trimmed);
  }, [text, onSend]);

  const handleEnqueue = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    if (!isRunning) {
      setText('');
      setSlashQuery(null);
      historyIndexRef.current = -1;
      draftRef.current = '';
      void onSend(trimmed);
      return;
    }
    setText('');
    setSlashQuery(null);
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
      // Slash menu navigation
      if (slashQuery !== null && filteredCommands.length > 0) {
        if (e.key === 'ArrowUp') {
          e.preventDefault();
          setSlashIndex(i => Math.max(0, i - 1));
          return;
        }
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          setSlashIndex(i => Math.min(filteredCommands.length - 1, i + 1));
          return;
        }
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          selectSlashCommand(filteredCommands[slashIndex]);
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          setSlashQuery(null);
          return;
        }
      }

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
    [
      slashQuery,
      filteredCommands,
      slashIndex,
      selectSlashCommand,
      handleSend,
      handleEnqueue,
      navigateHistory,
    ]
  );

  // Dynamic placeholder: hint at queue shortcut when agent is running
  const activePlaceholder = isRunning ? 'Type here, Ctrl+Shift+Q to queue...' : placeholder;

  return (
    <div className="aui-composer-root relative flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] overflow-hidden outline-none transition-colors focus-within:border-[var(--vscode-focusBorder)]">
      {/* Slash command menu — floats above the textarea */}
      {slashQuery !== null && filteredCommands.length > 0 && (
        <div
          ref={slashMenuRef}
          role="listbox"
          aria-label="Slash commands"
          className="absolute bottom-full left-0 z-50 mb-1 w-full max-h-56 overflow-y-auto rounded-[var(--vscode-cornerRadius-medium)] border border-[var(--vscode-widget-border,var(--vscode-input-border))] bg-[var(--vscode-quickInput-background,var(--vscode-input-background))] py-1 shadow-lg"
          style={{ boxShadow: '0 2px 8px var(--vscode-widget-shadow, rgba(0,0,0,0.3))' }}
        >
          {filteredCommands.map((cmd, i) => (
            <div
              key={cmd.name}
              role="option"
              aria-selected={i === slashIndex}
              onMouseEnter={() => setSlashIndex(i)}
              onMouseDown={e => {
                e.preventDefault();
                selectSlashCommand(cmd);
              }}
              className={cn(
                'flex cursor-pointer items-start gap-2 px-3 py-1.5 text-sm',
                i === slashIndex
                  ? 'bg-[var(--vscode-list-activeSelectionBackground)] text-[var(--vscode-list-activeSelectionForeground)]'
                  : 'text-[var(--vscode-foreground)] hover:bg-[var(--vscode-list-hoverBackground)]'
              )}
            >
              <Codicon
                name="sparkle"
                className="mt-0.5 shrink-0 text-[12px] text-[var(--vscode-textLink-foreground)]"
              />
              <div className="min-w-0">
                <span className="font-medium">/{cmd.name}</span>
                {cmd.description && (
                  <span
                    className="ml-2 truncate text-xs"
                    style={{
                      color:
                        i === slashIndex
                          ? 'var(--vscode-list-activeSelectionForeground)'
                          : 'var(--vscode-descriptionForeground)',
                    }}
                  >
                    {cmd.description}
                  </span>
                )}
              </div>
              {cmd.argumentHint && (
                <span
                  className="ml-auto shrink-0 text-xs italic"
                  style={{
                    color:
                      i === slashIndex
                        ? 'var(--vscode-list-activeSelectionForeground)'
                        : 'var(--vscode-descriptionForeground)',
                    opacity: 0.7,
                  }}
                >
                  {cmd.argumentHint}
                </span>
              )}
            </div>
          ))}
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
        aria-autocomplete={slashQuery !== null ? 'list' : undefined}
        aria-haspopup={slashQuery !== null ? 'listbox' : undefined}
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
