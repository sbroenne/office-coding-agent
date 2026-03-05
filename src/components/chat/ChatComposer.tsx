import React, { type FC, type ReactNode, useCallback, useEffect, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';

interface ChatComposerProps {
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  isRunning: boolean;
  placeholder?: string;
  leftToolbar?: ReactNode;
  rightToolbar?: ReactNode;
  autoFocus?: boolean;
}

export const ChatComposer: FC<ChatComposerProps> = ({
  onSend,
  onCancel,
  isRunning,
  placeholder = 'Send a message...',
  leftToolbar,
  rightToolbar,
  autoFocus = true,
}) => {
  const [text, setText] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Auto-resize textarea as content changes
  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${el.scrollHeight}px`;
  }, [text]);

  const handleSend = useCallback(() => {
    const trimmed = text.trim();
    if (!trimmed) return;
    setText('');
    void onSend(trimmed);
  }, [text, onSend]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    },
    [handleSend]
  );

  return (
    <div className="aui-composer-root relative flex w-full flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] overflow-hidden outline-none transition-colors focus-within:border-[var(--vscode-focusBorder)]">
      <textarea
        ref={textareaRef}
        value={text}
        onChange={e => setText(e.target.value)}
        onKeyDown={handleKeyDown}
        placeholder={placeholder}
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
        <div className="flex items-center gap-0.5">{leftToolbar}</div>
        <div className="flex items-center gap-0.5">
          {rightToolbar}

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
