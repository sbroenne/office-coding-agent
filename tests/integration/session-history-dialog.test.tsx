/**
 * Integration tests for SessionHistoryPanel.
 *
 * Renders the real component with test data. Validates session display,
 * restore/delete buttons, host filtering, and empty state.
 */
import { describe, it, expect, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SessionHistoryPanel } from '@/components/SessionHistoryDialog';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';

function makeSession(overrides: Partial<SessionHistoryItem> = {}): SessionHistoryItem {
  return {
    id: 'session-1',
    title: 'Test Conversation',
    host: 'excel',
    updatedAt: Date.now() - 60_000, // 1 minute ago
    messages: [{ role: 'user', content: 'hello' }],
    ...overrides,
  };
}

describe('Integration: SessionHistoryPanel', () => {
  it('shows empty state when no sessions exist', () => {
    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText('No saved conversations yet.')).toBeInTheDocument();
    expect(screen.getByText(/0 saved conversation/)).toBeInTheDocument();
  });

  it('shows session count text', () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Chat 1' }),
      makeSession({ id: 's2', title: 'Chat 2' }),
    ];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText(/2 saved conversations for excel/)).toBeInTheDocument();
  });

  it('renders session titles', () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Budget Planning' }),
      makeSession({ id: 's2', title: 'Data Analysis' }),
    ];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText('Budget Planning')).toBeInTheDocument();
    expect(screen.getByText('Data Analysis')).toBeInTheDocument();
  });

  it('shows "Active" for the active session', () => {
    const sessions = [makeSession({ id: 's1', title: 'Active Chat' })];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId="s1"
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText('Active')).toBeInTheDocument();
  });

  it('shows "Restore" for non-active sessions', () => {
    const sessions = [makeSession({ id: 's1', title: 'Old Chat' })];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId="s2"
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText('Restore')).toBeInTheDocument();
  });

  it('calls onRestoreSession when Restore is clicked', async () => {
    const onRestore = vi.fn();
    const sessions = [makeSession({ id: 's1', title: 'Chat' })];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={onRestore}
        onDeleteSession={() => {}}
      />
    );

    await userEvent.click(screen.getByText('Restore'));
    expect(onRestore).toHaveBeenCalledWith('s1');
  });

  it('calls onDeleteSession when Delete button is clicked', async () => {
    const onDelete = vi.fn();
    const sessions = [makeSession({ id: 's1', title: 'Chat' })];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={onDelete}
      />
    );

    await userEvent.click(screen.getByLabelText('Delete session'));
    expect(onDelete).toHaveBeenCalledWith('s1');
  });

  it('filters sessions by host', () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Excel Chat', host: 'excel' }),
      makeSession({ id: 's2', title: 'PPT Chat', host: 'powerpoint' }),
    ];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    expect(screen.getByText('Excel Chat')).toBeInTheDocument();
    expect(screen.queryByText('PPT Chat')).not.toBeInTheDocument();
    expect(screen.getByText(/1 saved conversation for excel/)).toBeInTheDocument();
  });

  it('sorts sessions by updatedAt descending (most recent first)', () => {
    const sessions = [
      makeSession({ id: 's1', title: 'Older', updatedAt: 1000 }),
      makeSession({ id: 's2', title: 'Newer', updatedAt: 2000 }),
    ];

    renderWithProviders(
      <SessionHistoryPanel
        host="excel"
        sessions={sessions}
        activeSessionId={null}
        onRestoreSession={() => {}}
        onDeleteSession={() => {}}
      />
    );

    const titles = screen.getAllByText(/Older|Newer/).map(el => el.textContent);
    expect(titles[0]).toBe('Newer');
    expect(titles[1]).toBe('Older');
  });
});
