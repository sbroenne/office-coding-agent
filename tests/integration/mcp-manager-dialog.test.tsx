/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel. Bundled servers are shown with status
 * indicators. Plugin servers are installed via the Copilot CLI Plugin Hub.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { McpManagerPanel } from '@/components/McpManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: McpManagerPanel', () => {
  it('renders with bundled servers and Built-in badges', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText('workiq')).toBeInTheDocument();
    expect(screen.getByText('powerbi')).toBeInTheDocument();
    expect(screen.getAllByText('Built-in').length).toBeGreaterThanOrEqual(2);
  });

  it('shows plugin servers note', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText(/Additional MCP servers are configured via plugins/i)).toBeInTheDocument();
  });

  it('does not show import or add buttons', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.queryByRole('button', { name: /Import/i })).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /Add/i })).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Import mcp.json file')).not.toBeInTheDocument();
  });

  it('bundled servers show Built-in badge', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getAllByText('Built-in').length).toBeGreaterThanOrEqual(2);
    expect(screen.getByText('workiq')).toBeInTheDocument();
    expect(screen.getByText('powerbi')).toBeInTheDocument();
  });
});
