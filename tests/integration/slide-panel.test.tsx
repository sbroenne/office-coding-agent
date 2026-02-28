/**
 * Integration tests for SlidePanel.
 *
 * Verifies dialog role, aria attributes, back-button, Escape-to-close,
 * title / description rendering, and children-only-when-open behaviour.
 */
import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { SlidePanel } from '@/components/SlidePanel';

describe('Integration: SlidePanel', () => {
  it('renders with role="dialog" and aria-modal', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Settings">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog');
    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(dialog).toHaveAttribute('aria-label', 'Settings');
  });

  it('renders title in the header', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="My Panel">
        <p>content</p>
      </SlidePanel>
    );

    expect(screen.getByText('My Panel')).toBeInTheDocument();
  });

  it('renders description when provided', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel" description="Some description">
        <p>content</p>
      </SlidePanel>
    );

    expect(screen.getByText('Some description')).toBeInTheDocument();
  });

  it('does not render description when not provided', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    expect(screen.queryByText('Some description')).not.toBeInTheDocument();
  });

  it('renders children when open', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>Visible child</p>
      </SlidePanel>
    );

    expect(screen.getByText('Visible child')).toBeInTheDocument();
  });

  it('does not render children when closed', () => {
    render(
      <SlidePanel open={false} onClose={() => {}} title="Panel">
        <p>Hidden child</p>
      </SlidePanel>
    );

    // children are conditionally rendered: {open && children}
    expect(screen.queryByText('Hidden child')).not.toBeInTheDocument();
  });

  it('sets aria-hidden="true" when closed', () => {
    render(
      <SlidePanel open={false} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog', { hidden: true });
    expect(dialog).toHaveAttribute('aria-hidden', 'true');
  });

  it('sets aria-hidden="false" when open', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog');
    expect(dialog).toHaveAttribute('aria-hidden', 'false');
  });

  it('calls onClose when Back button is clicked', async () => {
    const onClose = vi.fn();

    render(
      <SlidePanel open={true} onClose={onClose} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    await userEvent.click(screen.getByLabelText('Back'));
    expect(onClose).toHaveBeenCalledOnce();
  });

  it('calls onClose when Escape key is pressed', async () => {
    const onClose = vi.fn();

    render(
      <SlidePanel open={true} onClose={onClose} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    // fireEvent because keydown listeners are on document
    fireEvent.keyDown(document, { key: 'Escape' });
    expect(onClose).toHaveBeenCalledOnce();
  });

  it('does not call onClose on Escape when closed', () => {
    const onClose = vi.fn();

    render(
      <SlidePanel open={false} onClose={onClose} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    fireEvent.keyDown(document, { key: 'Escape' });
    expect(onClose).not.toHaveBeenCalled();
  });

  it('renders back button with correct aria-label and title', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const backBtn = screen.getByLabelText('Back');
    expect(backBtn).toHaveAttribute('title', 'Back');
  });

  it('applies translate-x-0 when open and translate-x-full when closed', () => {
    const { rerender } = render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog');
    expect(dialog.className).toContain('translate-x-0');

    rerender(
      <SlidePanel open={false} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialogClosed = screen.getByRole('dialog', { hidden: true });
    expect(dialogClosed.className).toContain('translate-x-full');
  });

  // Bug regression: closed SlidePanel should be inert to prevent keyboard focus
  // from reaching off-screen Back buttons and other interactive elements.
  it('has inert attribute when closed to prevent keyboard focus leak', () => {
    render(
      <SlidePanel open={false} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog', { hidden: true });
    // The inert attribute should be present when the panel is closed
    expect(dialog).toHaveAttribute('inert');
  });

  it('does not have inert attribute when open', () => {
    render(
      <SlidePanel open={true} onClose={() => {}} title="Panel">
        <p>content</p>
      </SlidePanel>
    );

    const dialog = screen.getByRole('dialog');
    expect(dialog).not.toHaveAttribute('inert');
  });
});
