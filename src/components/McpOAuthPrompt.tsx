import React, { useEffect, useMemo, useState } from 'react';
import { Codicon } from '@/components/Codicon';

export interface McpOAuthPromptRequest {
  serverName: string;
  defaultLoginHint?: string;
  reason?: 'connect' | 'sign-in' | 'retry' | 'switch' | 'chat-required';
  blocking?: boolean;
}

interface McpOAuthPromptProps {
  request: McpOAuthPromptRequest | null;
  onClose: () => void;
  onSignIn: (serverName: string, loginHint?: string) => Promise<string | undefined>;
}

function getAliasFromLoginHint(loginHint: string | undefined): string {
  const trimmed = loginHint?.trim();
  if (!trimmed) return '';
  return trimmed.toLowerCase().endsWith('@microsoft.com')
    ? trimmed.slice(0, trimmed.indexOf('@'))
    : trimmed;
}

export const McpOAuthPrompt: React.FC<McpOAuthPromptProps> = ({ request, onClose, onSignIn }) => {
  const [aliasOverride, setAliasOverride] = useState<string | null>(null);
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setAliasOverride(request ? getAliasFromLoginHint(request.defaultLoginHint) : null);
    setError(null);
    setIsSigningIn(false);
  }, [request]);

  const alias = aliasOverride ?? '';
  const title = useMemo(() => {
    if (!request) return '';
    return request.reason === 'switch'
      ? `Switch account for ${request.serverName}`
      : `Sign in to ${request.serverName}`;
  }, [request]);

  if (!request) return null;

  const isChatRequired = request.blocking === true || request.reason === 'chat-required';
  const description = isChatRequired
    ? `Office Coding Agent needs you to sign in to ${request.serverName} before this MCP server can be used.`
    : `Choose the account Office Coding Agent should use for this MCP server.`;

  const startSignIn = async () => {
    const loginHint = alias.trim();
    setIsSigningIn(true);
    setError(null);
    try {
      await onSignIn(request.serverName, loginHint || undefined);
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setIsSigningIn(false);
    }
  };

  return (
    <div
      className="absolute inset-0 z-[60] flex items-center justify-center bg-[rgba(0,0,0,0.35)] p-3"
      role="presentation"
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className="w-full max-w-[420px] rounded-[var(--vscode-cornerRadius-large)] border border-border bg-[var(--vscode-editorWidget-background)] p-4 text-sm text-foreground shadow-lg"
      >
        <div className="mb-3 flex items-start gap-2">
          <Codicon
            name="account"
            className="mt-0.5 shrink-0 text-[16px] text-[var(--vscode-textLink-foreground)]"
          />
          <div className="min-w-0">
            <h2 className="text-[13px] font-semibold">{title}</h2>
            <p className="mt-1 text-xs text-muted-foreground">{description}</p>
          </div>
        </div>

        <label htmlFor="mcp-oauth-alias" className="mb-1 block text-xs font-medium">
          Username or alias
        </label>
        <input
          id="mcp-oauth-alias"
          className="w-full rounded-[var(--vscode-cornerRadius-medium)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] px-2 py-1.5 text-[13px] text-[var(--vscode-input-foreground)] outline-none focus:border-[var(--vscode-focusBorder)]"
          placeholder="alias or user@domain.com"
          value={alias}
          disabled={isSigningIn}
          onChange={event => setAliasOverride(event.target.value)}
        />
        <p className="mt-1 text-xs text-muted-foreground">
          Optional. Enter an alias or full email address.
        </p>
        {error && <p className="mt-2 text-xs text-[var(--vscode-errorForeground)]">{error}</p>}

        <div className="mt-4 flex justify-end gap-2">
          <button
            type="button"
            className="rounded-[var(--vscode-cornerRadius-medium)] border border-border px-3 py-1 text-xs hover:bg-accent disabled:opacity-60"
            disabled={isSigningIn}
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="rounded-[var(--vscode-cornerRadius-medium)] border border-[var(--vscode-button-background)] bg-[var(--vscode-button-background)] px-3 py-1 text-xs text-[var(--vscode-button-foreground)] hover:bg-[var(--vscode-button-hoverBackground)] disabled:opacity-60"
            disabled={isSigningIn}
            onClick={() => void startSignIn()}
          >
            {isSigningIn
              ? 'Connecting...'
              : request.reason === 'switch'
                ? 'Switch account'
                : 'Sign in'}
          </button>
        </div>
      </div>
    </div>
  );
};
