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
  it('renders the picker button with the default label', () => {
    renderWithProviders(<AgentPicker />);

    expect(screen.getByLabelText('Select agent')).toBeInTheDocument();
    expect(screen.getByText('Default')).toBeInTheDocument();
  });

  it('shows the generic default option instead of listing the bundled Office agent by name', async () => {
    const user = userEvent.setup();
    renderWithProviders(<AgentPicker />);

    await user.click(screen.getByLabelText('Select agent'));

    expect(screen.getByText('Use the built-in Office agent for this host')).toBeInTheDocument();
    expect(screen.getByText(/Plugin agents are managed by the Copilot CLI/)).toBeInTheDocument();
    expect(screen.getByText('copilot plugin')).toBeInTheDocument();
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

  it('falls back to the hidden built-in default when the stored selection no longer exists', () => {
    renderWithProviders(<AgentPicker />);

    expect(screen.getByText('Default')).toBeInTheDocument();
  });
});
