/**
 * Integration test: ChatHeader — New Conversation, panel buttons.
 *
 * ChatHeader contains: SessionHistoryPicker, Permissions, New Conversation.
 * SkillPicker and MCP pill live in the Composer toolbar (ChatPanel).
 * All settings panels are opened via onOpenPanel callback.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatHeader } from '@/components/ChatHeader';
import { useSettingsStore } from '@/stores/settingsStore';

const mockClearMessages = vi.fn();
const mockOpenPanel = vi.fn();

describe('Integration: ChatHeader', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    mockClearMessages.mockClear();
    mockOpenPanel.mockClear();
  });

  it('renders new conversation button', () => {
    renderWithProviders(
      <ChatHeader
        host="excel"
        onClearMessages={mockClearMessages}
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={vi.fn()}
        onDeleteSession={vi.fn()}
        onOpenPanel={mockOpenPanel}
      />
    );

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
        onOpenPanel={mockOpenPanel}
      />
    );

    await userEvent.click(screen.getByLabelText('New conversation'));
    expect(mockClearMessages).toHaveBeenCalledOnce();
  });

  it('Permissions button calls onOpenPanel with "permissions"', async () => {
    renderWithProviders(
      <ChatHeader
        host="excel"
        onClearMessages={mockClearMessages}
        sessions={[]}
        activeSessionId={null}
        onRestoreSession={vi.fn()}
        onDeleteSession={vi.fn()}
        onOpenPanel={mockOpenPanel}
      />
    );

    await userEvent.click(screen.getByLabelText('Permissions'));
    expect(mockOpenPanel).toHaveBeenCalledWith('permissions');
  });
});
