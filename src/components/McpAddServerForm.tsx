import React, { useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import type { McpTransportType } from '@/types';

interface McpAddServerFormProps {
  /** Pre-filled values for editing an existing server */
  initial?: {
    name: string;
    transport: McpTransportType;
    command?: string;
    args?: string[];
    url?: string;
    headers?: Record<string, string>;
    description?: string;
  };
  /** Whether the name field should be read-only (edit mode) */
  editMode?: boolean;
  existingNames: Set<string>;
  onSubmit: (config: {
    name: string;
    description?: string;
    transport: McpTransportType;
    command?: string;
    args?: string[];
    url?: string;
    headers?: Record<string, string>;
  }) => void;
  onCancel: () => void;
}

export const McpAddServerForm: React.FC<McpAddServerFormProps> = ({
  initial,
  editMode,
  existingNames,
  onSubmit,
  onCancel,
}) => {
  const [name, setName] = useState(initial?.name ?? '');
  const [description, setDescription] = useState(initial?.description ?? '');
  const [transport, setTransport] = useState<McpTransportType>(initial?.transport ?? 'stdio');
  const [command, setCommand] = useState(initial?.command ?? 'npx');
  const [args, setArgs] = useState(initial?.args?.join(' ') ?? '');
  const [url, setUrl] = useState(initial?.url ?? '');
  const [headersText, setHeadersText] = useState(
    initial?.headers ? JSON.stringify(initial.headers, null, 2) : ''
  );
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    const trimmedName = name.trim();
    if (!trimmedName) {
      setError('Name is required');
      return;
    }
    if (!editMode && existingNames.has(trimmedName)) {
      setError('A server with this name already exists');
      return;
    }

    if (transport === 'stdio') {
      if (!command.trim()) {
        setError('Command is required for stdio transport');
        return;
      }
      const parsedArgs = args.trim() ? args.trim().split(/\s+/) : [];
      onSubmit({
        name: trimmedName,
        description: description.trim() || undefined,
        transport: 'stdio',
        command: command.trim(),
        args: parsedArgs,
      });
    } else {
      if (!url.trim()) {
        setError('URL is required for HTTP/SSE transport');
        return;
      }
      let headers: Record<string, string> | undefined;
      if (headersText.trim()) {
        try {
          headers = JSON.parse(headersText) as Record<string, string>;
        } catch {
          setError('Headers must be valid JSON');
          return;
        }
      }
      onSubmit({
        name: trimmedName,
        description: description.trim() || undefined,
        transport,
        url: url.trim(),
        headers,
      });
    }
  };

  return (
    <form
      onSubmit={handleSubmit}
      className="space-y-3 rounded-md border border-border bg-muted/30 p-3"
    >
      <div className="flex items-center justify-between">
        <h4 className="text-xs font-medium">{editMode ? 'Edit Server' : 'Add Server'}</h4>
        <button
          type="button"
          onClick={onCancel}
          className="rounded-sm p-0.5 text-muted-foreground hover:text-foreground"
        >
          <Codicon name="close" className="text-sm" />
        </button>
      </div>

      <div className="space-y-2">
        <div>
          <label className="text-[10px] font-medium text-muted-foreground">Name</label>
          <input
            type="text"
            value={name}
            onChange={e => setName(e.target.value)}
            readOnly={editMode}
            placeholder="my-server"
            className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
          />
        </div>

        <div>
          <label className="text-[10px] font-medium text-muted-foreground">
            Description (optional)
          </label>
          <input
            type="text"
            value={description}
            onChange={e => setDescription(e.target.value)}
            placeholder="What this server does"
            className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
          />
        </div>

        <div>
          <label className="text-[10px] font-medium text-muted-foreground">Transport</label>
          <div className="mt-0.5 flex gap-2">
            {(['stdio', 'http', 'sse'] as const).map(t => (
              <label key={t} className="flex items-center gap-1 text-xs">
                <input
                  type="radio"
                  name="transport"
                  value={t}
                  checked={transport === t}
                  onChange={() => setTransport(t)}
                  className="size-3"
                />
                {t}
              </label>
            ))}
          </div>
        </div>

        {transport === 'stdio' ? (
          <>
            <div>
              <label className="text-[10px] font-medium text-muted-foreground">Command</label>
              <input
                type="text"
                value={command}
                onChange={e => setCommand(e.target.value)}
                placeholder="npx"
                className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
              />
            </div>
            <div>
              <label className="text-[10px] font-medium text-muted-foreground">
                Arguments (space-separated)
              </label>
              <input
                type="text"
                value={args}
                onChange={e => setArgs(e.target.value)}
                placeholder="-y @microsoft/workiq mcp"
                className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
              />
            </div>
          </>
        ) : (
          <>
            <div>
              <label className="text-[10px] font-medium text-muted-foreground">URL</label>
              <input
                type="text"
                value={url}
                onChange={e => setUrl(e.target.value)}
                placeholder="https://example.com/mcp"
                className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
              />
            </div>
            <div>
              <label className="text-[10px] font-medium text-muted-foreground">
                Headers (JSON, optional)
              </label>
              <textarea
                value={headersText}
                onChange={e => setHeadersText(e.target.value)}
                placeholder='{"Authorization": "Bearer ..."}'
                rows={2}
                className="mt-0.5 w-full rounded-md border border-border bg-background px-2 py-1 font-mono text-[10px] focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-[-1px]"
              />
            </div>
          </>
        )}
      </div>

      {error && <p className="text-[10px] text-[var(--vscode-errorForeground)]">{error}</p>}

      <div className="flex justify-end gap-2">
        <Button type="button" variant="ghost" size="sm" onClick={onCancel}>
          Cancel
        </Button>
        <Button type="submit" size="sm">
          {editMode ? 'Save' : 'Add Server'}
        </Button>
      </div>
    </form>
  );
};
