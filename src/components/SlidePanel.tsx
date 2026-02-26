import React from 'react';
import { ArrowLeft } from 'lucide-react';

interface SlidePanelProps {
  open: boolean;
  onClose: () => void;
  title: string;
  description?: string;
  children: React.ReactNode;
}

/**
 * Full-width slide-in panel that replaces the chat view.
 * Slides in from the right with a back-button header.
 */
export const SlidePanel: React.FC<SlidePanelProps> = ({
  open,
  onClose,
  title,
  description,
  children,
}) => {
  return (
    <div
      className={`absolute inset-0 z-10 flex flex-col bg-background transition-transform duration-300 ease-in-out ${
        open ? 'translate-x-0' : 'translate-x-full'
      }`}
      aria-hidden={!open}
    >
      {/* Panel header */}
      <div className="flex items-center gap-2 border-b border-border px-3 py-1.5 shrink-0">
        <button
          onClick={onClose}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="Back"
          title="Back"
        >
          <ArrowLeft className="size-4" />
        </button>
        <div className="min-w-0 flex-1">
          <h2 className="text-sm font-semibold truncate">{title}</h2>
          {description && (
            <p className="text-[10px] text-muted-foreground truncate">{description}</p>
          )}
        </div>
      </div>

      {/* Scrollable content */}
      <div className="flex-1 overflow-y-auto">{open && children}</div>
    </div>
  );
};
