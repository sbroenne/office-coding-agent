import {
  type CodeHeaderProps,
  MarkdownTextPrimitive,
  unstable_memoizeMarkdownComponents as memoizeMarkdownComponents,
  useIsMarkdownCodeBlock,
} from '@assistant-ui/react-markdown';
import { useThreadRuntime } from '@assistant-ui/react';
import remarkGfm from 'remark-gfm';
import { type FC, Children, isValidElement, memo, useState, useCallback } from 'react';
import { Codicon } from '@/components/Codicon';
import { TooltipIconButton } from '@/components/assistant-ui/tooltip-icon-button';
import { cn } from '@/lib/utils';

const MarkdownTextImpl = () => {
  return (
    <MarkdownTextPrimitive
      remarkPlugins={[remarkGfm]}
      className="aui-md"
      components={defaultComponents}
    />
  );
};

export const MarkdownText = memo(MarkdownTextImpl);

const useCopyToClipboard = ({
  copiedDuration = 3000,
}: {
  copiedDuration?: number;
} = {}) => {
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

const CodeHeader: FC<CodeHeaderProps> = ({ language, code }) => {
  const { isCopied, copyToClipboard } = useCopyToClipboard();
  const onCopy = () => {
    if (code) copyToClipboard(code);
  };

  // choices blocks: render cards instead of code header + code block
  if (language === 'choices') {
    return <ChoiceCards code={code} />;
  }

  // suggestions blocks: render VS Code-style follow-up links
  if (language === 'suggestions') {
    return <SuggestionLinks code={code} />;
  }

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
      <TooltipIconButton tooltip={isCopied ? 'Copied!' : 'Copy'} side="left" onClick={onCopy}>
        {isCopied ? (
          <Codicon name="check" className="text-sm" />
        ) : (
          <Codicon name="copy" className="text-sm" />
        )}
      </TooltipIconButton>
    </div>
  );
};

// ─── Choice Cards ───
// Parses ```choices JSON blocks into clickable cards.

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

const ChoiceCards: FC<{ code: string }> = ({ code }) => {
  const threadRuntime = useThreadRuntime();
  const choices = parseChoices(code);
  const [freeformText, setFreeformText] = useState('');

  const handleChoice = useCallback(
    (label: string) => {
      threadRuntime.append({
        role: 'user',
        content: [{ type: 'text', text: label }],
      });
    },
    [threadRuntime]
  );

  const handleFreeform = useCallback(() => {
    const trimmed = freeformText.trim();
    if (!trimmed) return;
    threadRuntime.append({
      role: 'user',
      content: [{ type: 'text', text: trimmed }],
    });
    setFreeformText('');
  }, [threadRuntime, freeformText]);

  if (!choices || choices.length === 0) return null;

  return (
    <div className="aui-choices-wrapper mt-2 flex flex-col">
      {choices.map((choice, i) => (
        <button
          key={choice.label}
          onClick={() => handleChoice(choice.label)}
          title={choice.description}
          className="flex items-baseline gap-3 rounded-[var(--vscode-cornerRadius-medium)] px-2 py-1 text-left text-sm transition-colors hover:bg-[var(--vscode-list-hoverBackground)]"
        >
          <span className="w-4 shrink-0 text-right text-xs text-muted-foreground select-none">
            {i + 1}
          </span>
          <span className="text-foreground">{choice.label}</span>
        </button>
      ))}
      <div className="flex items-baseline gap-3 px-2 py-1">
        <span className="w-4 shrink-0 text-right text-xs text-muted-foreground select-none">
          {choices.length + 1}
        </span>
        <textarea
          className="flex-1 resize-none bg-transparent text-sm text-foreground outline-none border-b transition-colors placeholder:text-[var(--vscode-input-placeholderForeground)]"
          style={{ borderColor: 'var(--vscode-widget-border)' }}
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
  );
};

// ─── Suggestion Links ───
// Parses ```suggestions JSON blocks into VS Code-style follow-up links.
// Same JSON format as choices: [{"label": "..."}]
// Rendered as plain text-link buttons with a sparkle icon (no border, no freeform).

const SuggestionLinks: FC<{ code: string }> = ({ code }) => {
  const threadRuntime = useThreadRuntime();
  const suggestions = parseChoices(code); // reuse same parser

  const handleClick = useCallback(
    (label: string) => {
      threadRuntime.append({
        role: 'user',
        content: [{ type: 'text', text: label }],
      });
    },
    [threadRuntime]
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

const defaultComponents = memoizeMarkdownComponents({
  h1: ({ className, ...props }) => (
    <h1
      className={cn('aui-md-h1 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{
        fontSize: 'var(--vscode-chat-font-size-body-xxl)',
        margin: '1.5em 0 0.875em 0',
      }}
      {...props}
    />
  ),
  h2: ({ className, ...props }) => (
    <h2
      className={cn('aui-md-h2 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{
        fontSize: 'var(--vscode-chat-font-size-body-xl)',
        margin: '1.5em 0 0.875em 0',
      }}
      {...props}
    />
  ),
  h3: ({ className, ...props }) => (
    <h3
      className={cn('aui-md-h3 scroll-m-20 font-semibold first:mt-0 last:mb-0', className)}
      style={{
        fontSize: 'var(--vscode-chat-font-size-body-l)',
        margin: '1.5em 0 0.875em 0',
      }}
      {...props}
    />
  ),
  h4: ({ className, ...props }) => (
    <h4
      className={cn('aui-md-h4 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h5: ({ className, ...props }) => (
    <h5
      className={cn('aui-md-h5 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  h6: ({ className, ...props }) => (
    <h6
      className={cn('aui-md-h6 font-medium text-sm first:mt-0 last:mb-0', className)}
      style={{ margin: '1.5em 0 0.875em 0' }}
      {...props}
    />
  ),
  p: ({ className, ...props }) => (
    <p
      className={cn('aui-md-p leading-normal first:mt-0 last:mb-0', className)}
      style={{ margin: '0 0 16px 0' }}
      {...props}
    />
  ),
  a: ({ className, ...props }) => (
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
  blockquote: ({ className, ...props }) => (
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
  ul: ({ className, ...props }) => (
    <ul
      className={cn('aui-md-ul list-disc marker:text-muted-foreground [&>li]:mt-1', className)}
      style={{ paddingInlineStart: 24, margin: '4px 0' }}
      {...props}
    />
  ),
  ol: ({ className, ...props }) => (
    <ol
      className={cn('aui-md-ol list-decimal marker:text-muted-foreground [&>li]:mt-1', className)}
      style={{ paddingInlineStart: 28, margin: '4px 0' }}
      {...props}
    />
  ),
  hr: ({ className, ...props }) => (
    <hr className={cn('aui-md-hr my-2 border-muted-foreground/20', className)} {...props} />
  ),
  table: ({ className, ...props }) => (
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
  th: ({ className, ...props }) => (
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
  td: ({ className, ...props }) => (
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
  tr: ({ className, ...props }) => <tr className={cn('aui-md-tr m-0 p-0', className)} {...props} />,
  li: ({ className, ...props }) => (
    <li
      className={cn('aui-md-li leading-normal', className)}
      style={{ margin: '4px 0' }}
      {...props}
    />
  ),
  sup: ({ className, ...props }) => (
    <sup className={cn('aui-md-sup [&>a]:text-xs [&>a]:no-underline', className)} {...props} />
  ),
  pre: ({ className, children, ...props }) => {
    // Hide the raw pre/code wrapper for custom-rendered blocks (choices, suggestions)
    const isCustomBlock = Children.toArray(children).some(
      kid =>
        isValidElement<{ className?: string }>(kid) &&
        (kid.props.className?.includes('language-choices') === true ||
          kid.props.className?.includes('language-suggestions') === true)
    );
    if (isCustomBlock) return null;
    return (
      <pre
        className={cn(
          'aui-md-pre overflow-x-auto rounded-t-none p-3 text-xs leading-relaxed',
          className
        )}
        style={{
          fontFamily: 'var(--vscode-monospace-font)',
          fontSize: 'var(--vscode-chat-font-size-body-xs)',
          background: 'var(--vscode-textPreformat-background)',
          border: '1px solid var(--vscode-widget-border)',
          borderTop: 0,
          borderBottomLeftRadius: 'var(--vscode-cornerRadius-medium)',
          borderBottomRightRadius: 'var(--vscode-cornerRadius-medium)',
        }}
        {...props}
      >
        {children}
      </pre>
    );
  },
  code: function Code({ className, ...props }) {
    const isCodeBlock = useIsMarkdownCodeBlock();
    if (isCodeBlock) {
      return (
        <code
          className={cn(className)}
          style={{ fontFamily: 'var(--vscode-monospace-font)' }}
          {...props}
        />
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
        {...props}
      />
    );
  },
  CodeHeader,
});
