/**
 * Integration tests for SessionHistoryPicker.
 *
 * Renders the real Radix Popover component. Tests popover open/close,
 * session list rendering, relative time display, and callback wiring.
 */
import { describe, it, expect, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SessionHistoryPicker } from '@/components/SessionHistoryPicker';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';

function makeSession(overrides: Partial<SessionHistoryItem> = {}): SessionHistoryItem {
  return {
    id: 'session-1',
    title: 'Test Conversation',
    host: 'excel',
    updatedAt: Date.now() - 60_000, // 1 minute ago
    messages: [],
    ...overrides,
  };
}

describe('Integration: SessionHistoryPicker', () => {
  it('renders the trigger button with history icon', () => {
    renderWithProviders(
      <SessionHistoryPicker
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByLabelText('Session history')).toBeInTheDocument();
  });

  it('opens popover when trigger is clicked', async () => {
    renderWithProviders(
      <SessionHistoryPicker
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('Session history')).toBeInTheDocument();
    });
  });

  it('shows empty state when no sessions exist', async () => {
    renderWithProviders(
      <SessionHistoryPicker
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('No previous conversations yet.')).toBeInTheDocument();
    });
  });

  it('shows session titles in the popover', async () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Budget Chat' }),
      makeSession({ id: 's2', title: 'Report Chat' }),
    ];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('Budget Chat')).toBeInTheDocument();
      expect(screen.getByText('Report Chat')).toBeInTheDocument();
    });
  });

  it('shows "active" label for the active session', async () => {
    const sessions = [makeSession({ id: 's1', title: 'Active Chat' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId="s1"
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('active')).toBeInTheDocument();
    });
  });

  it('calls onRestoreSession when session title is clicked', async () => {
    const onRestore = vi.fn();
    const sessions = [makeSession({ id: 's1', title: 'My Chat' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={onRestore}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('My Chat')).toBeInTheDocument();
    });

    await userEvent.click(screen.getByText('My Chat'));
    expect(onRestore).toHaveBeenCalledWith('s1');
  });

  it('calls onDeleteSession when delete button is clicked', async () => {
    const onDelete = vi.fn();
    const sessions = [makeSession({ id: 's1', title: 'Old Chat' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={onDelete}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByLabelText('Delete session')).toBeInTheDocument();
    });

    await userEvent.click(screen.getByLabelText('Delete session'));
    expect(onDelete).toHaveBeenCalledWith('s1');
  });

  it('shows "Manage history…" footer when sessions exist', async () => {
    const sessions = [makeSession({ id: 's1', title: 'Chat' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('Manage history…')).toBeInTheDocument();
      expect(screen.getByText('1 total')).toBeInTheDocument();
    });
  });

  it('calls onOpenPanel when Manage history is clicked', async () => {
    const onOpenPanel = vi.fn();
    const sessions = [makeSession({ id: 's1', title: 'Chat' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
        onOpenPanel={onOpenPanel}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('Manage history…')).toBeInTheDocument();
    });

    await userEvent.click(screen.getByText('Manage history…'));
    expect(onOpenPanel).toHaveBeenCalledWith('history');
  });

  it('shows host label for each session', async () => {
    const sessions = [makeSession({ id: 's1', title: 'Chat', host: 'excel' })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('excel')).toBeInTheDocument();
    });
  });

  it('displays relative time for non-active sessions', async () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Recent', updatedAt: Date.now() - 5 * 60_000 }), // 5m ago
    ];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      expect(screen.getByText('5m ago')).toBeInTheDocument();
    });
  });

  // Bug regression: long session titles should be truncated with CSS
  // to prevent tall list items from pushing other sessions out of view.
  it('applies truncate class to session title buttons', async () => {
    const longTitle = 'This is a very long session title that should be truncated to prevent layout breakage in the popover';
    const sessions = [makeSession({ id: 's-long', title: longTitle })];

    renderWithProviders(
      <SessionHistoryPicker
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByLabelText('Session history'));

    await waitFor(() => {
      const btn = screen.getByText(longTitle);
      expect(btn.className).toContain('truncate');
    });
  });
});
