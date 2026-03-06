/**
 * Integration tests for AgentManagerPanel.
 *
 * Renders the real AgentManagerPanel. Bundled agents are shown read-only with
 * download-as-template buttons. Custom agents are installed via plugins.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
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

  it('shows plugin hub message in the "From plugins" section', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText('From plugins')).toBeInTheDocument();
    expect(screen.getByText(/installed via the Copilot CLI Plugin Hub/i)).toBeInTheDocument();
  });

  it('bundled agents each have a "Download as template" button', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(
      screen.getByRole('button', { name: 'Download Excel as template' })
    ).toBeInTheDocument();
  });

  it('does not show import upload buttons', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.queryByLabelText('Import agents ZIP file')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Import agent Markdown file')).not.toBeInTheDocument();
  });
});
