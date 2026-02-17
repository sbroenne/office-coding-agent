/**
 * Integration test: ChatHeader — SkillPicker, New Conversation, SettingsDialog.
 *
 * ChatHeader now contains: "AI Chat" title, SkillPicker, New Conversation
 * button, and SettingsDialog. The ModelPicker and AgentPicker have moved
 * to ChatPanel's input toolbar.
 */
import React, { useState } from 'react';
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ChatHeader } from '@/components/ChatHeader';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Only mock external services, NOT components ───
vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

const mockClearMessages = vi.fn();

/** Wrapper that manages settingsOpen state so the controlled dialog works. */
const StatefulChatHeader: React.FC<{ onClearMessages: () => void }> = ({ onClearMessages }) => {
  const [settingsOpen, setSettingsOpen] = useState(false);
  return (
    <ChatHeader
      onClearMessages={onClearMessages}
      settingsOpen={settingsOpen}
      onSettingsOpenChange={setSettingsOpen}
    />
  );
};

describe('Integration: ChatHeader', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    mockClearMessages.mockClear();
  });

  it('renders title, skill picker, and toolbar buttons', () => {
    renderWithProviders(<StatefulChatHeader onClearMessages={mockClearMessages} />);

    expect(screen.getByText('AI Chat')).toBeInTheDocument();
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
    expect(screen.getByLabelText('New conversation')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Settings' })).toBeInTheDocument();
  });

  it('calls onClearMessages when New conversation is clicked', async () => {
    renderWithProviders(<StatefulChatHeader onClearMessages={mockClearMessages} />);

    await userEvent.click(screen.getByLabelText('New conversation'));
    expect(mockClearMessages).toHaveBeenCalledOnce();
  });

  it('Settings dialog toolbar button opens the dialog', async () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'Direct Open',
      resourceUrl: 'https://direct.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(<StatefulChatHeader onClearMessages={mockClearMessages} />);

    // Click the Settings toolbar button directly
    const settingsButton = screen.getByRole('button', { name: 'Settings' });
    await userEvent.click(settingsButton);

    // Dialog should open
    await waitFor(() => {
      expect(screen.getByRole('heading', { name: 'Settings' })).toBeInTheDocument();
    });
  });

  it('opens Settings dialog when settingsOpen prop is true', () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'Controlled Open',
      resourceUrl: 'https://controlled.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(
      <ChatHeader
        onClearMessages={mockClearMessages}
        settingsOpen={true}
        onSettingsOpenChange={() => {}}
      />
    );

    expect(screen.getByRole('heading', { name: 'Settings' })).toBeInTheDocument();
  });
});
