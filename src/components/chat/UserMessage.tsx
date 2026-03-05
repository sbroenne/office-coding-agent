import { type FC, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import type { ChatMessage } from '@/types';

interface UserMessageProps {
  message: ChatMessage;
  onEdit?: (messageId: string, newText: string) => void;
}

export const UserMessage: FC<UserMessageProps> = ({ message, onEdit }) => {
  const [editing, setEditing] = useState(false);
  const [editText, setEditText] = useState('');

  const textContent = message.content
    .filter(p => p.type === 'text')
    .map(p => (p.type === 'text' ? p.text : ''))
    .join('');

  const startEdit = () => {
    setEditText(textContent);
    setEditing(true);
  };

  const cancelEdit = () => {
    setEditing(false);
    setEditText('');
  };

  const submitEdit = () => {
    const trimmed = editText.trim();
    if (!trimmed || !onEdit) {
      cancelEdit();
      return;
    }
    onEdit(message.id, trimmed);
    setEditing(false);
    setEditText('');
  };

  if (editing) {
    return (
      <div className="fade-in slide-in-from-bottom-1 relative flex w-full animate-in px-4 py-2 duration-150">
        <div className="ml-auto w-full max-w-[90%]">
          <div className="flex flex-col rounded-[var(--vscode-cornerRadius-large)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] overflow-hidden outline-none transition-colors focus-within:border-[var(--vscode-focusBorder)]">
            <textarea
              autoFocus
              className="max-h-32 min-h-10 w-full resize-none bg-transparent px-3 pt-2 pb-1 text-[13px] outline-none placeholder:text-[var(--vscode-input-placeholderForeground)] focus-visible:ring-0"
              rows={1}
              value={editText}
              onChange={e => setEditText(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  submitEdit();
                } else if (e.key === 'Escape') {
                  cancelEdit();
                }
              }}
            />
            <div className="mx-1 mb-1.5 flex items-center justify-end gap-1 border-t border-border/40 pt-1">
              <button
                type="button"
                onClick={cancelEdit}
                title="Cancel edit"
                className="h-7 px-2 rounded-md text-xs text-muted-foreground hover:text-foreground hover:bg-accent transition-colors"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={submitEdit}
                title="Send edit"
                className="flex h-7 w-7 items-center justify-center rounded-md transition-colors hover:bg-accent"
              >
                <Codicon name="send" className="text-base" />
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      className="aui-user-message-root fade-in slide-in-from-bottom-1 group/message relative flex w-full animate-in px-4 py-2 duration-150"
      data-role="user"
    >
      <div
        className={cn(
          'aui-user-message-bubble wrap-break-word text-foreground text-[13px] leading-[1.5em]',
          onEdit && 'cursor-pointer'
        )}
        onClick={onEdit ? startEdit : undefined}
        title={onEdit ? 'Click to edit' : undefined}
      >
        {textContent}
      </div>

      {onEdit && (
        <button
          type="button"
          onClick={startEdit}
          title="Edit message"
          aria-label="Edit message"
          className={cn(
            'aui-user-action-bar absolute right-2 top-2',
            'flex h-6 w-6 items-center justify-center rounded p-1',
            'opacity-0 transition-opacity duration-150 group-hover/message:opacity-100',
            'text-muted-foreground/60 hover:bg-[var(--vscode-toolbar-hoverBackground)] hover:text-muted-foreground'
          )}
        >
          <Codicon name="edit" className="text-sm" />
        </button>
      )}
    </div>
  );
};
