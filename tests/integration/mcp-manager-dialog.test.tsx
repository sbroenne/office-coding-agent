/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel. Bundled servers are shown with status
 * indicators and toggle buttons. Users can upload custom servers via JSON.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
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

  it('shows the JSON upload button', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByLabelText('Import MCP servers from JSON file')).toBeInTheDocument();
  });

  it('each server has a toggle button', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByLabelText('Toggle workiq')).toBeInTheDocument();
    expect(screen.getByLabelText('Toggle powerbi')).toBeInTheDocument();
  });

  it('clicking toggle disables a server and updates the store', async () => {
    renderWithProviders(<McpManagerPanel />);

    const toggle = screen.getByLabelText('Toggle workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'true');

    await userEvent.click(toggle);

    expect(useSettingsStore.getState().disabledMcpServerNames).toContain('workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'false');
  });

  it('clicking toggle again re-enables a server', async () => {
    renderWithProviders(<McpManagerPanel />);

    const toggle = screen.getByLabelText('Toggle workiq');
    await userEvent.click(toggle);
    await userEvent.click(toggle);

    expect(useSettingsStore.getState().disabledMcpServerNames).not.toContain('workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'true');
  });

  it('shows an imported server with a Remove button (no Built-in badge)', () => {
    useSettingsStore.getState().addImportedMcpServer({
      name: 'my-mcp',
      description: 'Custom server',
      transport: 'http',
      url: 'https://example.com/mcp',
    });

    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText('my-mcp')).toBeInTheDocument();
    expect(screen.getByLabelText('Remove my-mcp')).toBeInTheDocument();
    // Bundled servers still show Built-in; imported server does not
    expect(screen.getAllByText('Built-in').length).toBe(2);
  });

  it('clicking Remove deletes an imported server from the store', async () => {
    useSettingsStore.getState().addImportedMcpServer({
      name: 'deletable-mcp',
      transport: 'http',
      url: 'https://example.com/mcp',
    });

    renderWithProviders(<McpManagerPanel />);

    await userEvent.click(screen.getByLabelText('Remove deletable-mcp'));

    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(0);
    expect(screen.queryByText('deletable-mcp')).not.toBeInTheDocument();
  });
});
