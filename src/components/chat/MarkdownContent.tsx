import React, { memo, useState, useCallback, Children, isValidElement } from 'react';
import Markdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useChatActions } from '@/contexts/ChatActionsContext';
import type { Components } from 'react-markdown';

// Helper to type react-markdown v10 component props
type MdProps<T extends keyof React.JSX.IntrinsicElements> = React.JSX.IntrinsicElements[T] & {
  node?: unknown;
};

// ─── Copy to clipboard hook ───────────────────────────────────────────────────

const useCopyToClipboard = ({ copiedDuration = 3000 } = {}) => {
  const [isCopied, setIsCopied] = useState(false);
  const copyToClipboard = (value: string) => {
    if (!value) return;
    void navigator.clipboard.writeText(value).then(() => {
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), copiedDuration);
    });
  };
  return { isCopied, copyToClipboard };
};

// ─── CodeHeader ───────────────────────────────────────────────────────────────

const CodeHeader: React.FC<{ language: string; code: string }> = ({ language, code }) => {
  const { isCopied, copyToClipboard } = useCopyToClipboard();

  return (
    <div
      className="flex items-center justify-between rounded-t-[var(--vscode-cornerRadius-medium)] border border-b-0 px-4 py-2 text-sm"
      style={{
        background: 'var(--vscode-editorWidget-background)',
        borderColor: 'var(--vscode-widget-border)',
        color: 'var(--vscode-editor-foreground)',
      }}
    >
      <span className="lowercase">{language}</span>
      <button
        onClick={() => copyToClipboard(code)}
        title={isCopied ? 'Copied!' : 'Copy'}
        className="flex h-6 w-6 items-center justify-center rounded p-1 transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
        style={{ color: 'var(--vscode-icon-foreground)' }}
      >
        {isCopied ? (
          <Codicon name="check" className="text-sm" />
        ) : (
          <Codicon name="copy" className="text-sm" />
        )}
      </button>
    </div>
  );
};

// ─── Choice Cards ─────────────────────────────────────────────────────────────

interface ChoiceItem {
  label: string;
  description?: string;
}

function parseChoices(code: string): ChoiceItem[] | null {
  try {
    const parsed: unknown = JSON.parse(code.trim());
    if (!Array.isArray(parsed)) return null;
    return parsed.filter(
      (item): item is ChoiceItem =>
        typeof item === 'object' && item !== null && typeof (item as ChoiceItem).label === 'string'
    );
  } catch {
    return null;
  }
}

const ChoiceCards: React.FC<{ code: string }> = ({ code }) => {
  const { send } = useChatActions();
  const choices = parseChoices(code);
  const [freeformText, setFreeformText] = useState('');

  const handleChoice = useCallback(
    (label: string) => {
      void send(label);
    },
    [send]
  );

  const handleFreeform = useCallback(() => {
    const trimmed = freeformText.trim();
    if (!trimmed) return;
    void send(trimmed);
    setFreeformText('');
  }, [send, freeformText]);

  if (!choices || choices.length === 0) return null;

  return (
    <div className="aui-choices-wrapper chat-question-carousel-container">
      <div className="chat-question-input-container">
        <div className="chat-question-list">
          {choices.map((choice, i) => (
            <button
              key={choice.label}
              onClick={() => handleChoice(choice.label)}
              title={choice.description}
              className="chat-question-list-item"
            >
              <span className="chat-question-list-number">{i + 1}</span>
              <span className="chat-question-list-label">
                <span className="chat-question-list-label-title">{choice.label}</span>
                {choice.description && (
                  <span className="chat-question-list-label-desc">{choice.description}</span>
                )}
              </span>
            </button>
          ))}
        </div>
        <div className="chat-question-freeform">
          <span className="chat-question-freeform-number">{choices.length + 1}</span>
          <textarea
            className="chat-question-freeform-textarea"
            placeholder="Enter custom answer"
            rows={1}
            value={freeformText}
            onChange={e => setFreeformText(e.target.value)}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleFreeform();
              }
            }}
          />
        </div>
      </div>
    </div>
  );
};

// ─── Suggestion Links ─────────────────────────────────────────────────────────

const SuggestionLinks: React.FC<{ code: string }> = ({ code }) => {
  const { send } = useChatActions();
  const suggestions = parseChoices(code);

  const handleClick = useCallback(
    (label: string) => {
      void send(label);
    },
    [send]
  );

  if (!suggestions || suggestions.length === 0) return null;

  return (
    <div className="aui-suggestions-wrapper mt-3 flex flex-col items-start gap-1.5">
      {suggestions.map(suggestion => (
        <button
          key={suggestion.label}
          onClick={() => handleClick(suggestion.label)}
          className="flex items-center gap-1.5 border-none bg-transparent p-0 text-left text-xs transition-colors hover:text-foreground"
          style={{ color: 'var(--vscode-textLink-foreground)' }}
        >
          <Codicon name="sparkle" className="shrink-0 text-xs" />
          {suggestion.label}
        </button>
      ))}
    </div>
  );
};

// ─── Custom markdown components ───────────────────────────────────────────────

type PreProps = React.ComponentPropsWithoutRef<'pre'> & { node?: unknown };

const PreComponent: React.FC<PreProps> = ({ children, ...props }) => {
  // Extract language and code text from the inner <code> element
  const codeEl = Children.toArray(children).find(
    (c): c is React.ReactElement<{ className?: string; children?: React.ReactNode }> =>
      isValidElement(c)
  ) as React.ReactElement<{ className?: string; children?: React.ReactNode }> | undefined;

  const className = codeEl?.props?.className ?? '';
  const match = /language-(\S+)/.exec(className);
  const language = match ? match[1] : '';
  const rawCode = codeEl?.props?.children;
  const code = (
    Array.isArray(rawCode) ? rawCode.join('') : typeof rawCode === 'string' ? rawCode : ''
  ).replace(/\n$/, '');

  if (language === 'choices') return <ChoiceCards code={code} />;
  if (language === 'suggestions') return <SuggestionLinks code={code} />;

  // Exclude the node prop from DOM spreading
  const { node: _node, ...domProps } = props as PreProps;

  return (
    <>
      {language && <CodeHeader language={language} code={code} />}
      <pre
        className={cn(
          'aui-md-pre overflow-x-auto p-3 text-xs leading-relaxed',
          language ? 'rounded-t-none' : 'rounded-[var(--vscode-cornerRadius-medium)]'
        )}
        style={{
          fontFamily: 'var(--vscode-monospace-font)',
          fontSize: 'var(--vscode-chat-font-size-body-xs)',
          background: 'var(--vscode-textPreformat-background)',
          border: '1px solid var(--vscode-widget-border)',
          borderTop: language ? 0 : undefined,
          borderBottomLeftRadius: 'var(--vscode-cornerRadius-medium)',
          borderBottomRightRadius: 'var(--vscode-cornerRadius-medium)',
        }}
        {...domProps}
      >
        {children}
      </pre>
    </>
  );
};

type CodeProps = React.ComponentPropsWithoutRef<'code'> & { node?: unknown };

const CodeComponent: React.FC<CodeProps> = ({ className, children, ...props }) => {
  const isBlock = (className ?? '').includes('language-');
  // Exclude node from DOM spreading
  const { node: _node, ...domProps } = props as CodeProps;

  if (isBlock) {
    return (
      <code
        className={className}
        style={{ fontFamily: 'var(--vscode-monospace-font)' }}
        {...domProps}
      >
        {children}
      </code>
    );
  }
  return (
    <code
      className={cn('aui-md-inline-code', className)}
      style={{
        fontFamily: 'var(--vscode-monospace-font)',
        fontSize: 'var(--vscode-chat-font-size-body-xs)',
        color: 'var(--vscode-textPreformat-foreground)',
        background: 'var(--vscode-textPreformat-background)',
        padding: '1px 3px',
        borderRadius: 4,
        border: '1px solid var(--vscode-textPreformat-border)',
      }}
      {...domProps}
    >
      {children}
    </code>
  );
};

const markdownComponents: Components = {
  h1: ({ className, node: _node, ...props }: MdProps<'h1'>) => (
    <h1
      className={cn('aui-md-h1 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{ fontSize: 'var(--vscode-chat-font-size-body-xxl)', margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h2: ({ className, node: _node, ...props }: MdProps<'h2'>) => (
    <h2
      className={cn('aui-md-h2 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{ fontSize: 'var(--vscode-chat-font-size-body-xl)', margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h3: ({ className, node: _node, ...props }: MdProps<'h3'>) => (
    <h3
      className={cn('aui-md-h3 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{ fontSize: 'var(--vscode-chat-font-size-body-l)', margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h4: ({ className, node: _node, ...props }: MdProps<'h4'>) => (
    <h4
      className={cn('aui-md-h4 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h5: ({ className, node: _node, ...props }: MdProps<'h5'>) => (
    <h5
      className={cn('aui-md-h5 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h6: ({ className, node: _node, ...props }: MdProps<'h6'>) => (
    <h6
      className={cn('aui-md-h6 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  p: ({ className, node: _node, ...props }: MdProps<'p'>) => (
    <p
      className={cn('aui-md-p leading-normal first:mt-0 last:mb-0', className)}
      style={{ margin: '0 0 16px 0' }}
      {...props}
    />
  ),
  a: ({ className, node: _node, ...props }: MdProps<'a'>) => (
    <a
      className={cn('aui-md-a underline underline-offset-2', className)}
      style={{ color: 'var(--vscode-textLink-foreground)' }}
      onMouseEnter={e => {
        (e.currentTarget as HTMLElement).style.color = 'var(--vscode-textLink-activeForeground)';
      }}
      onMouseLeave={e => {
        (e.currentTarget as HTMLElement).style.color = 'var(--vscode-textLink-foreground)';
      }}
      {...props}
    />
  ),
  blockquote: ({ className, node: _node, ...props }: MdProps<'blockquote'>) => (
    <blockquote
      className={cn('aui-md-blockquote my-2.5 italic', className)}
      style={{
        padding: '0 16px 0 10px',
        borderLeft: '5px solid var(--vscode-textBlockQuote-border)',
        background: 'var(--vscode-textBlockQuote-background)',
        borderRadius: 2,
      }}
      {...props}
    />
  ),
  ul: ({ className, node: _node, ...props }: MdProps<'ul'>) => (
    <ul
      className={cn('aui-md-ul list-disc marker:text-muted-foreground [&>li]:mt-1', className)}
      style={{ paddingInlineStart: 24, margin: '4px 0' }}
      {...props}
    />
  ),
  ol: ({ className, node: _node, ...props }: MdProps<'ol'>) => (
    <ol
      className={cn('aui-md-ol list-decimal marker:text-muted-foreground [&>li]:mt-1', className)}
      style={{ paddingInlineStart: 28, margin: '4px 0' }}
      {...props}
    />
  ),
  hr: ({ className, node: _node, ...props }: MdProps<'hr'>) => (
    <hr className={cn('aui-md-hr my-2 border-muted-foreground/20', className)} {...props} />
  ),
  table: ({ className, node: _node, ...props }: MdProps<'table'>) => (
    <div className="my-2 overflow-x-auto">
      <table
        className={cn('aui-md-table w-full', className)}
        style={{
          borderCollapse: 'separate',
          border: '1px solid var(--vscode-chat-requestBorder)',
          borderRadius: 'var(--vscode-cornerRadius-medium)',
        }}
        {...props}
      />
    </div>
  ),
  th: ({ className, node: _node, ...props }: MdProps<'th'>) => (
    <th
      className={cn(
        'aui-md-th px-2 py-1 text-left font-medium first:rounded-tl last:rounded-tr [[align=center]]:text-center [[align=right]]:text-right',
        className
      )}
      style={{
        borderBottom: '1px solid var(--vscode-chat-requestBorder)',
        borderRight: '1px solid var(--vscode-chat-requestBorder)',
      }}
      {...props}
    />
  ),
  td: ({ className, node: _node, ...props }: MdProps<'td'>) => (
    <td
      className={cn(
        'aui-md-td px-2 py-1 text-left [[align=center]]:text-center [[align=right]]:text-right',
        className
      )}
      style={{
        borderBottom: '1px solid var(--vscode-chat-requestBorder)',
        borderRight: '1px solid var(--vscode-chat-requestBorder)',
      }}
      {...props}
    />
  ),
  tr: ({ className, node: _node, ...props }: MdProps<'tr'>) => (
    <tr className={cn('aui-md-tr m-0 p-0', className)} {...props} />
  ),
  li: ({ className, node: _node, ...props }: MdProps<'li'>) => (
    <li
      className={cn('aui-md-li leading-normal', className)}
      style={{ margin: '4px 0' }}
      {...props}
    />
  ),
  sup: ({ className, node: _node, ...props }: MdProps<'sup'>) => (
    <sup className={cn('aui-md-sup [&>a]:text-xs [&>a]:no-underline', className)} {...props} />
  ),
  pre: PreComponent as Components['pre'],
  code: CodeComponent as Components['code'],
};

// ─── MarkdownContent ──────────────────────────────────────────────────────────

interface MarkdownContentProps {
  text: string;
  className?: string;
}

const MarkdownContentImpl: React.FC<MarkdownContentProps> = ({ text, className }) => {
  return (
    <div className={cn('aui-md', className)}>
      <Markdown remarkPlugins={[remarkGfm]} components={markdownComponents}>
        {text}
      </Markdown>
    </div>
  );
};

export const MarkdownContent = memo(MarkdownContentImpl);
