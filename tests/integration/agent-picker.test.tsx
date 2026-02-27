/**
 * Integration test: AgentPicker component.
 *
 * Renders the real AgentPicker with real Zustand store and real
 * bundled agents (loaded via rawMarkdownPlugin). Verifies selecting
 * agents updates the store and shows the current selection.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { AgentPicker } from '@/components/AgentPicker';
import { useSettingsStore } from '@/stores/settingsStore';
import { getAgents } from '@/services/agents';

const mockOpenPanel = vi.fn();

beforeEach(() => {
  useSettingsStore.getState().reset();
  mockOpenPanel.mockClear();
});

describe('Integration: AgentPicker', () => {
  it('renders button with default agent name', () => {
    renderWithProviders(<AgentPicker />);
    expect(screen.getByText('Excel')).toBeInTheDocument();
  });

  it('shows agent list when clicked', async () => {
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByText('Excel'));

    const agents = getAgents();
    for (const agent of agents) {
      // Agent name should appear as a radio option
      const items = screen.getAllByText(agent.metadata.name);
      expect(items.length).toBeGreaterThanOrEqual(1);
    }

    expect(screen.getByText('Manage agents…')).toBeInTheDocument();
  });

  it('shows agent description as secondary content', async () => {
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByText('Excel'));

    const agents = getAgents();
    const firstSentence = agents[0].metadata.description.split('.')[0];
    expect(screen.getByText(firstSentence)).toBeInTheDocument();
  });

  it('calls onOpenPanel when manage agents button is clicked', async () => {
    renderWithProviders(<AgentPicker onOpenPanel={mockOpenPanel} />);

    await userEvent.click(screen.getByText('Excel'));

    const manageButton = screen.getByRole('button', { name: /manage agents/i });
    manageButton.focus();
    await userEvent.keyboard('{Enter}');

    expect(mockOpenPanel).toHaveBeenCalledWith('agents');
  });

  it('store reflects the default active agent', () => {
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
    expect(useSettingsStore.getState().getActiveAgent()).toBe('Excel');
  });
});
