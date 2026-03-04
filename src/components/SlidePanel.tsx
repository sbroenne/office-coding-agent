import React, { useEffect, useRef, useCallback } from 'react';
import { Codicon } from '@/components/Codicon';

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
 * Traps focus when open and restores focus on close.
 */
export const SlidePanel: React.FC<SlidePanelProps> = ({
  open,
  onClose,
  title,
  description,
  children,
}) => {
  const panelRef = useRef<HTMLDivElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);

  // Escape key handler
  const handleKeyDown = useCallback(
    (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onClose();
      }
    },
    [onClose]
  );

  // Focus management: trap focus on open, restore on close
  useEffect(() => {
    if (open) {
      previousFocusRef.current = document.activeElement as HTMLElement;
      // Wait for transition to complete before focusing
      const timer = setTimeout(() => {
        const backButton = panelRef.current?.querySelector<HTMLButtonElement>(
          'button[aria-label="Back"]'
        );
        backButton?.focus();
      }, 310);
      document.addEventListener('keydown', handleKeyDown);
      return () => {
        clearTimeout(timer);
        document.removeEventListener('keydown', handleKeyDown);
      };
    } else {
      // Restore focus to the element that opened the panel
      previousFocusRef.current?.focus();
      previousFocusRef.current = null;
    }
  }, [open, handleKeyDown]);

  // Focus trap: keep Tab within the panel
  const handleFocusTrap = useCallback((e: React.KeyboardEvent) => {
    if (e.key !== 'Tab') return;
    const panel = panelRef.current;
    if (!panel) return;

    const focusable = panel.querySelectorAll<HTMLElement>(
      'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
    );
    if (focusable.length === 0) return;

    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  }, []);

  // Sync the `inert` attribute imperatively — React 18 strips unknown HTML attrs
  useEffect(() => {
    const el = panelRef.current;
    if (!el) return;
    if (!open) {
      el.setAttribute('inert', '');
    } else {
      el.removeAttribute('inert');
    }
  }, [open]);

  return (
    <div
      ref={panelRef}
      role="dialog"
      aria-modal="true"
      aria-label={title}
      className={`absolute inset-0 z-50 flex flex-col bg-background will-change-transform transition-transform duration-300 ease-in-out ${
        open ? 'translate-x-0' : 'translate-x-full'
      }`}
      aria-hidden={!open}
      onKeyDown={handleFocusTrap}
    >
      {/* Panel header */}
      <div className="flex h-[35px] items-center gap-2 border-b border-border px-3 shrink-0">
        <button
          onClick={onClose}
          className="inline-flex h-8 w-8 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="Back"
          title="Back"
        >
          <Codicon name="arrow-left" className="text-base" />
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
