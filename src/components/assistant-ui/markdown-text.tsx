import {
  type CodeHeaderProps,
  MarkdownTextPrimitive,
  unstable_memoizeMarkdownComponents as memoizeMarkdownComponents,
  useIsMarkdownCodeBlock,
} from '@assistant-ui/react-markdown';
import { useThreadRuntime } from '@assistant-ui/react';
import remarkGfm from 'remark-gfm';
import { type FC, memo, useState, useCallback } from 'react';
import { CheckIcon, CopyIcon, SparklesIcon } from 'lucide-react';
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

  return (
    <div className="flex items-center justify-between rounded-t-lg border border-b-0 bg-zinc-900 px-4 py-2 text-sm text-white">
      <span className="lowercase">{language}</span>
      <TooltipIconButton tooltip={isCopied ? 'Copied!' : 'Copy'} side="left" onClick={onCopy}>
        {isCopied ? <CheckIcon /> : <CopyIcon />}
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

  const handleClick = useCallback(
    (label: string) => {
      threadRuntime.append({
        role: 'user',
        content: [{ type: 'text', text: label }],
      });
    },
    [threadRuntime]
  );

  if (!choices || choices.length === 0) return null;

  return (
    <div className="aui-choices-wrapper my-2 flex flex-col gap-1.5">
      {choices.map(choice => (
        <button
          key={choice.label}
          onClick={() => handleClick(choice.label)}
          className="flex items-center gap-2 rounded-xl border border-border bg-background px-3 py-2 text-left text-sm text-foreground shadow-sm transition-all hover:bg-accent hover:shadow-md active:scale-[0.98]"
        >
          <SparklesIcon className="size-3.5 shrink-0 text-primary" />
          <div className="min-w-0">
            <div className="font-medium">{choice.label}</div>
            {choice.description && (
              <div className="text-xs text-muted-foreground">{choice.description}</div>
            )}
          </div>
        </button>
      ))}
    </div>
  );
};

const defaultComponents = memoizeMarkdownComponents({
  h1: ({ className, ...props }) => (
    <h1
      className={cn(
        'aui-md-h1 mb-6 mt-2 scroll-m-20 text-2xl font-bold tracking-tight first:mt-0 last:mb-0',
        className
      )}
      {...props}
    />
  ),
  h2: ({ className, ...props }) => (
    <h2
      className={cn(
        'aui-md-h2 mb-4 mt-6 scroll-m-20 text-xl font-semibold tracking-tight first:mt-0 last:mb-0',
        className
      )}
      {...props}
    />
  ),
  h3: ({ className, ...props }) => (
    <h3
      className={cn(
        'aui-md-h3 mb-2 mt-4 scroll-m-20 text-lg font-semibold tracking-tight first:mt-0 last:mb-0',
        className
      )}
      {...props}
    />
  ),
  h4: ({ className, ...props }) => (
    <h4
      className={cn(
        'aui-md-h4 mt-2 mb-1 scroll-m-20 font-medium text-sm first:mt-0 last:mb-0',
        className
      )}
      {...props}
    />
  ),
  h5: ({ className, ...props }) => (
    <h5
      className={cn('aui-md-h5 mt-2 mb-1 font-medium text-sm first:mt-0 last:mb-0', className)}
      {...props}
    />
  ),
  h6: ({ className, ...props }) => (
    <h6
      className={cn('aui-md-h6 mt-2 mb-1 font-medium text-sm first:mt-0 last:mb-0', className)}
      {...props}
    />
  ),
  p: ({ className, ...props }) => (
    <p
      className={cn('aui-md-p my-2.5 leading-normal first:mt-0 last:mb-0', className)}
      {...props}
    />
  ),
  a: ({ className, ...props }) => (
    <a
      className={cn(
        'aui-md-a text-primary underline underline-offset-2 hover:text-primary/80',
        className
      )}
      {...props}
    />
  ),
  blockquote: ({ className, ...props }) => (
    <blockquote
      className={cn(
        'aui-md-blockquote my-2.5 border-muted-foreground/30 border-l-2 pl-3 text-muted-foreground italic',
        className
      )}
      {...props}
    />
  ),
  ul: ({ className, ...props }) => (
    <ul
      className={cn(
        'aui-md-ul my-2 ml-4 list-disc marker:text-muted-foreground [&>li]:mt-1',
        className
      )}
      {...props}
    />
  ),
  ol: ({ className, ...props }) => (
    <ol
      className={cn(
        'aui-md-ol my-2 ml-4 list-decimal marker:text-muted-foreground [&>li]:mt-1',
        className
      )}
      {...props}
    />
  ),
  hr: ({ className, ...props }) => (
    <hr className={cn('aui-md-hr my-2 border-muted-foreground/20', className)} {...props} />
  ),
  table: ({ className, ...props }) => (
    <table
      className={cn(
        'aui-md-table my-2 w-full border-separate border-spacing-0 overflow-y-auto',
        className
      )}
      {...props}
    />
  ),
  th: ({ className, ...props }) => (
    <th
      className={cn(
        'aui-md-th bg-muted px-2 py-1 text-left font-medium first:rounded-tl-lg last:rounded-tr-lg [[align=center]]:text-center [[align=right]]:text-right',
        className
      )}
      {...props}
    />
  ),
  td: ({ className, ...props }) => (
    <td
      className={cn(
        'aui-md-td border-muted-foreground/20 border-b border-l px-2 py-1 text-left last:border-r [[align=center]]:text-center [[align=right]]:text-right',
        className
      )}
      {...props}
    />
  ),
  tr: ({ className, ...props }) => (
    <tr
      className={cn(
        'aui-md-tr m-0 border-b p-0 first:border-t [&:last-child>td:first-child]:rounded-bl-lg [&:last-child>td:last-child]:rounded-br-lg',
        className
      )}
      {...props}
    />
  ),
  li: ({ className, ...props }) => (
    <li className={cn('aui-md-li leading-normal', className)} {...props} />
  ),
  sup: ({ className, ...props }) => (
    <sup className={cn('aui-md-sup [&>a]:text-xs [&>a]:no-underline', className)} {...props} />
  ),
  pre: ({ className, ...props }) => (
    <pre
      className={cn(
        'aui-md-pre overflow-x-auto rounded-t-none rounded-b-lg border border-border/50 border-t-0 bg-muted/30 p-3 text-xs leading-relaxed',
        className
      )}
      {...props}
    />
  ),
  code: function Code({ className, ...props }) {
    const isCodeBlock = useIsMarkdownCodeBlock();
    return (
      <code
        className={cn(
          !isCodeBlock &&
            'aui-md-inline-code rounded-md border border-border/50 bg-muted/50 px-1.5 py-0.5 font-mono text-[0.85em]',
          className
        )}
        {...props}
      />
    );
  },
  CodeHeader,
});
