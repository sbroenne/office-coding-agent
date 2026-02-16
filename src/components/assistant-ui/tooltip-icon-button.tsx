import { forwardRef } from 'react';
import { Slottable } from '@radix-ui/react-slot';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { Button, type buttonVariants } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import type { VariantProps } from 'class-variance-authority';

export const TooltipIconButton = forwardRef<
  HTMLButtonElement,
  React.ComponentProps<'button'> &
    VariantProps<typeof buttonVariants> & {
      tooltip: string;
      side?: 'top' | 'bottom' | 'left' | 'right';
      asChild?: boolean;
    }
>(({ children, tooltip, side = 'bottom', className, variant = 'ghost', ...rest }, ref) => {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <Button
          variant={variant}
          size="icon"
          {...rest}
          className={cn('size-6 p-1', className)}
          ref={ref}
        >
          <Slottable>{children}</Slottable>
          <span className="sr-only">{tooltip}</span>
        </Button>
      </TooltipTrigger>
      <TooltipContent side={side}>{tooltip}</TooltipContent>
    </Tooltip>
  );
});

TooltipIconButton.displayName = 'TooltipIconButton';
