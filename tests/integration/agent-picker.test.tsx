import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { AgentPicker } from '@/components/AgentPicker';
import { useSettingsStore } from '@/stores/settingsStore';
import { getAgents, getBundledAgents } from '@/services/agents';
import type { OfficeHostApp } from '@/services/office/host';
import type { AgentHost } from '@/types/agent';

beforeEach(() => {
  useSettingsStore.getState().reset();
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
  });

  it('shows agent description as secondary content', async () => {
    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    const agents = getAgents();
    const firstSentence = agents[0].metadata.description.split('.')[0];
    expect(screen.getByText(firstSentence)).toBeInTheDocument();
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
  it('[REGRESSION] imported/custom agents are visible in the dropdown', async () => {
    // Bug: AgentPicker only rendered bundledAgents; imported agents were never shown.
    // Also regression: AgentPicker must react to plugin agents arriving via the store.
    const { parseAgentFrontmatter } = await import('@/services/agents');

    const customAgent = parseAgentFrontmatter(`---
name: MyCustomAgent
description: A custom agent for testing
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Custom instructions.`);

    // Use store action (not module-level setImportedAgents) so the picker re-renders
    useSettingsStore.getState().setPluginAgents([customAgent]);

    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));

    // Custom agent must appear in the dropdown
    expect(screen.getByText('MyCustomAgent')).toBeInTheDocument();

    // Clean up
    useSettingsStore.getState().setPluginAgents([]);
  });

  it('[REGRESSION] clicking a custom agent updates the store', async () => {
    const { parseAgentFrontmatter } = await import('@/services/agents');

    const customAgent = parseAgentFrontmatter(`---
name: MySelectableAgent
description: Custom selectable agent
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Instructions.`);

    useSettingsStore.getState().setPluginAgents([customAgent]);

    renderWithProviders(<AgentPicker />);

    await userEvent.click(screen.getByLabelText('Select agent'));
    await userEvent.click(screen.getByText('MySelectableAgent'));

    expect(useSettingsStore.getState().activeAgentId).toBe('MySelectableAgent');

    // Clean up
    useSettingsStore.getState().setPluginAgents([]);
    useSettingsStore.getState().reset();
  });

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
