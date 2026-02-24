/**
 * Integration test: ChatHeader — SkillPicker and New Conversation button.
 *
 * ChatHeader now contains only: SkillPicker and New Conversation button.
 * SettingsDialog and McpManagerDialog have been removed.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatHeader } from '@/components/ChatHeader';
import { useSettingsStore } from '@/stores/settingsStore';

const mockClearMessages = vi.fn();

describe('Integration: ChatHeader', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    mockClearMessages.mockClear();
  });

  it('renders skill picker and new conversation button', () => {
    renderWithProviders(
      <ChatHeader
        host="excel"
        onClearMessages={mockClearMessages}
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={vi.fn()}
        onDeleteSession={vi.fn()}
      />
    );

    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
    expect(screen.getByLabelText('New conversation')).toBeInTheDocument();
  });

  it('calls onClearMessages when New conversation is clicked', async () => {
    renderWithProviders(
      <ChatHeader
        host="excel"
        onClearMessages={mockClearMessages}
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={vi.fn()}
        onDeleteSession={vi.fn()}
      />
    );

    await userEvent.click(screen.getByLabelText('New conversation'));
    expect(mockClearMessages).toHaveBeenCalledOnce();
  });
});
