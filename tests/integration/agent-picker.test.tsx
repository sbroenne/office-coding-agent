import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { AgentPicker } from '@/components/AgentPicker';
import { useSettingsStore } from '@/stores/settingsStore';
import { getDefaultAgent, getPickerAgents } from '@/services/agents';
import type { OfficeHostApp } from '@/services/office/host';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: AgentPicker', () => {
  it('renders the picker button with the default label when only built-in agents exist', () => {
    renderWithProviders(<AgentPicker />);

    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
    expect(screen.getByText('Default')).toBeInTheDocument();
  });

  it('shows the generic default option instead of listing the bundled Office agent by name', async () => {
    const user = userEvent.setup();
    renderWithProviders(<AgentPicker />);

    await user.click(screen.getByLabelText('Select agent'));

    expect(screen.getByText('Use the built-in Office agent for this host')).toBeInTheDocument();
    expect(screen.queryByText('Excel')).not.toBeInTheDocument();
  });

  it('store reflects the default active agent', () => {
    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
    expect(useSettingsStore.getState().getActiveAgent()).toBe('Excel');
  });

  it.each(['excel', 'powerpoint', 'word', 'outlook'] as OfficeHostApp[])(
    'keeps built-in host defaults while exposing no picker-visible agents for host "%s"',
    testHost => {
      expect(getDefaultAgent(testHost)).toBeDefined();
      expect(getPickerAgents(testHost)).toEqual([]);
    }
  );

  it('[REGRESSION] imported/custom agents are visible in the dropdown', async () => {
    const user = userEvent.setup();
    const { parseAgentFrontmatter } = await import('@/services/agents');

    const customAgent = parseAgentFrontmatter(`---
name: MyCustomAgent
description: A custom agent for testing
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Custom instructions.`);

    useSettingsStore.getState().setPluginAgents([customAgent]);
    renderWithProviders(<AgentPicker />);

    await user.click(screen.getByLabelText('Select agent'));

    expect(screen.getByText('MyCustomAgent')).toBeInTheDocument();
    expect(screen.queryByText('Excel')).not.toBeInTheDocument();
  });

  it('[REGRESSION] clicking a custom agent updates the store', async () => {
    const user = userEvent.setup();
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

    await user.click(screen.getByLabelText('Select agent'));
    await user.click(screen.getByText('MySelectableAgent'));

    expect(useSettingsStore.getState().activeAgentId).toBe('MySelectableAgent');
  });

  it('[REGRESSION] hides built-in Office plugin agents but shows external plugin agents', async () => {
    const user = userEvent.setup();

    useSettingsStore.getState().setPluginAgents([
      {
        metadata: {
          name: 'office-excel',
          description: 'Built-in Office plugin agent',
          version: '1.0.0',
          hosts: ['excel'],
          defaultForHosts: [],
        },
        instructions: 'Built-in Office instructions.',
      },
      {
        metadata: {
          name: 'Contoso Excel Agent',
          description: 'External plugin agent',
          version: '1.0.0',
          hosts: ['excel'],
          defaultForHosts: [],
        },
        instructions: 'External plugin instructions.',
      },
    ]);

    renderWithProviders(<AgentPicker />);

    await user.click(screen.getByLabelText('Select agent'));

    expect(screen.queryByText('office-excel')).not.toBeInTheDocument();
    expect(screen.getByText('Contoso Excel Agent')).toBeInTheDocument();
  });

  it('allows returning to the hidden built-in default after choosing a custom agent', async () => {
    const user = userEvent.setup();
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
    useSettingsStore.getState().setActiveAgent('MySelectableAgent');
    renderWithProviders(<AgentPicker />);

    expect(screen.getByText('MySelectableAgent')).toBeInTheDocument();

    await user.click(screen.getByLabelText('Select agent'));
    await user.click(screen.getByText('Default'));

    expect(useSettingsStore.getState().activeAgentId).toBe('Excel');
    expect(screen.getByText('Default')).toBeInTheDocument();
  });

  it('falls back to the hidden built-in default when the stored selection no longer exists', () => {
    renderWithProviders(<AgentPicker />);

    expect(screen.getByText('Default')).toBeInTheDocument();
  });
});
