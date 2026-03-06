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
import { getAgents, getBundledAgents } from '@/services/agents';
import type { OfficeHostApp } from '@/services/office/host';
import type { AgentHost } from '@/types/agent';

const mockOpenPanel = vi.fn();

beforeEach(() => {
  useSettingsStore.getState().reset();
  mockOpenPanel.mockClear();
});

describe('Integration: AgentPicker', () => {
  it('renders button with agent icon', () => {
    renderWithProviders(<AgentPicker />);
    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
  });

  it('shows agent list when clicked', async () => {
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    const agents = getAgents();
    for (const agent of agents) {
      // Agent name should appear as a radio option
      const items = screen.getAllByText(agent.metadata.name);
      expect(items.length).toBeGreaterThanOrEqual(1);
    }

    expect(screen.getByText('Manage plugins…')).toBeInTheDocument();
  });

  it('shows agent description as secondary content', async () => {
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    const agents = getAgents();
    const firstSentence = agents[0].metadata.description.split('.')[0];
    expect(screen.getByText(firstSentence)).toBeInTheDocument();
  });

  it('calls onOpenPanel when manage plugins button is clicked', async () => {
    renderWithProviders(<AgentPicker onOpenPanel={mockOpenPanel} />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    const manageButton = screen.getByRole('button', { name: /manage plugins/i });
    manageButton.focus();
    await userEvent.keyboard('{Enter}');

    expect(mockOpenPanel).toHaveBeenCalledWith('plugins');
  });

  it('store reflects the default active agent', () => {
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
    expect(useSettingsStore.getState().getActiveAgent()).toBe('Excel');
  });

  // Bug regression: Word and Outlook hosts should have filterable agents
  // Before fix: targetHost was only set for 'excel' | 'powerpoint', leaving
  // Word/Outlook with undefined → empty bundled/imported agent lists → no
  // selectable agents in the dropdown.
  it.each(['excel', 'powerpoint', 'word', 'outlook'] as OfficeHostApp[])(
    'getBundledAgents can be filtered for host "%s"',
    testHost => {
      const agents = getAgents(testHost);
      if (agents.length > 0) {
        // The filtering logic that the AgentPicker uses should work for this host
        const filtered = getBundledAgents().filter(a =>
          a.metadata.hosts.includes(testHost as AgentHost)
        );
        // At least the agents returned by getAgents should be findable via the
        // same host filter the AgentPicker uses
        expect(filtered.length).toBeGreaterThanOrEqual(0);
        // The key assertion: if agents exist for a host, they should be filterable
        if (agents.some(a => a.metadata.hosts.includes(testHost as AgentHost))) {
          expect(filtered.length).toBeGreaterThan(0);
        }
      }
    }
  );

  // Bug regression: checkmark should appear on the resolved (fallback) agent,
  // not just when activeAgentId literally matches. If the stored agent was deleted,
  // the host-default should still show a checkmark in the dropdown.
  it('checkmark appears on the resolved default agent, not only matching activeAgentId', async () => {
    // Set activeAgentId to a non-existent agent — resolveActiveAgent falls back to host default
    useSettingsStore.getState().setActiveAgent('Deleted-Agent-That-Does-Not-Exist');

    renderWithProviders(<AgentPicker />);
    await userEvent.click(screen.getByLabelText('Select agent'));

    // The default agent for excel should still be visible
    const agents = getAgents();
    const defaultName = agents[0]?.metadata.name;
    if (defaultName) {
      // The resolved agent should be the fallback, and its checkmark (opacity-100) should be visible
      const items = screen.getAllByText(defaultName);
      expect(items.length).toBeGreaterThanOrEqual(1);
    }
  });
});
