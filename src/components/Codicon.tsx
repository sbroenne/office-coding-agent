/**
 * Thin wrapper around @vscode/codicons icon font.
 *
 * Usage:  <Codicon name="robot" className="text-base" />
 *
 * Icon names: https://microsoft.github.io/vscode-codicons/dist/codicon.html
 */
import React from 'react';
import { cn } from '@/lib/utils';

interface CodiconProps {
  /** Codicon name without the "codicon-" prefix, e.g. "robot", "extensions", "lightbulb-sparkle" */
  name: string;
  className?: string;
  'aria-hidden'?: boolean;
}

export const Codicon: React.FC<CodiconProps> = ({
  name,
  className,
  'aria-hidden': ariaHidden = true,
}) => <i className={cn(`codicon codicon-${name}`, className)} aria-hidden={ariaHidden} />;
