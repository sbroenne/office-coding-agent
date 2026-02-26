/**
 * Integration test: ChatHeader — SkillPicker, New Conversation, panel buttons.
 *
 * ChatHeader contains: SkillPicker, SessionHistoryPicker, Permissions, MCP, New Conversation.
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

  it('renders skill picker and new conversation button', () => {
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
        onOpenPanel={mockOpenPanel}
      />
    );

    await userEvent.click(screen.getByLabelText('New conversation'));
    expect(mockClearMessages).toHaveBeenCalledOnce();
  });

  it('MCP Servers button calls onOpenPanel with "mcp"', async () => {
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

    await userEvent.click(screen.getByLabelText('MCP Servers'));
    expect(mockOpenPanel).toHaveBeenCalledWith('mcp');
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
