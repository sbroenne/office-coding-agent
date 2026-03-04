import * as React from 'react';
import { cn } from '@/lib/utils';

function Input({ className, type, ...props }: React.ComponentProps<'input'>) {
  return (
    <input
      type={type}
      data-slot="input"
      className={cn(
        'flex h-[26px] w-full rounded-[var(--vscode-cornerRadius-small)] border border-[var(--vscode-input-border)] bg-[var(--vscode-input-background)] px-2 py-1 text-[13px] text-[var(--vscode-input-foreground)] placeholder:text-[var(--vscode-input-placeholderForeground)] focus-visible:border-[var(--vscode-focusBorder)] focus-visible:outline-none focus-visible:ring-0 transition-colors file:border-0 file:bg-transparent file:text-[13px] file:font-medium disabled:cursor-not-allowed disabled:opacity-50',
        className
      )}
      {...props}
    />
  );
}

export { Input };
