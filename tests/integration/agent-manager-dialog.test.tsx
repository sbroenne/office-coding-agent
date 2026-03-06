/**
 * Integration tests for AgentManagerPanel.
 *
 * Renders the real AgentManagerPanel. Bundled agents are shown read-only with
 * download-as-template buttons. Users can upload custom agents via ZIP or .md.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { AgentManagerPanel } from '@/components/AgentManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: AgentManagerPanel', () => {
  it('shows bundled agents in the read-only section', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    expect(screen.getAllByText('Excel').length).toBeGreaterThanOrEqual(1);
  });

  it('shows the Uploaded section with count', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText(/Uploaded \(0\)/i)).toBeInTheDocument();
  });

  it('bundled agents each have a "Download as template" button', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(
      screen.getByRole('button', { name: 'Download Excel as template' })
    ).toBeInTheDocument();
  });

  it('shows upload buttons for ZIP and .md', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByLabelText('Import agents ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import agent Markdown file')).toBeInTheDocument();
  });

  it('shows empty-state message when no agents are uploaded', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText(/No uploaded agents/i)).toBeInTheDocument();
  });

  it('shows uploaded agent with delete button when store has an imported agent', () => {
    const agent = {
      metadata: {
        name: 'my-agent',
        description: 'A test agent',
        version: '1.0.0',
        hosts: ['excel' as const],
        defaultForHosts: [],
      },
      instructions: 'Do something.',
    };
    useSettingsStore.getState().addImportedAgent(agent);

    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText('my-agent')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Remove my-agent' })).toBeInTheDocument();
    expect(screen.getByText(/Uploaded \(1\)/i)).toBeInTheDocument();
  });

  it('clicking Remove deletes the agent from the store', async () => {
    const agent = {
      metadata: {
        name: 'removable-agent',
        description: '',
        version: '1.0.0',
        hosts: ['excel' as const],
        defaultForHosts: [],
      },
      instructions: 'instructions',
    };
    useSettingsStore.getState().addImportedAgent(agent);

    renderWithProviders(<AgentManagerPanel />);

    await userEvent.click(screen.getByRole('button', { name: 'Remove removable-agent' }));

    expect(useSettingsStore.getState().importedAgents).toHaveLength(0);
    expect(screen.queryByText('removable-agent')).not.toBeInTheDocument();
  });
});
