/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel with the real Zustand store.
 * Tests import and display flows without a live MCP server.
 * Note: Import operations parse files locally but no longer persist to the store
 * (plugin state is managed by the Copilot CLI via ~/.copilot/config.json).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpManagerPanel } from '@/components/McpManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: McpManagerPanel', () => {
  it('renders with bundled servers and import button', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText('workiq')).toBeInTheDocument();
    expect(screen.getByText('powerbi')).toBeInTheDocument();
    expect(screen.getAllByText('Built-in').length).toBeGreaterThanOrEqual(2);
    expect(screen.getByRole('button', { name: /Import/i })).toBeInTheDocument();
  });

  it('shows error for invalid JSON', async () => {
    renderWithProviders(<McpManagerPanel />);

    const file = new File(['not-json'], 'mcp.json', { type: 'application/json' });
    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, file);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('shows Add button that opens add server form', async () => {
    renderWithProviders(<McpManagerPanel />);

    const addButton = screen.getByRole('button', { name: /Add/i });
    await userEvent.click(addButton);

    expect(screen.getByText('Add Server', { selector: 'h4' })).toBeInTheDocument();
    expect(screen.getByPlaceholderText('my-server')).toBeInTheDocument();
  });

  it('bundled servers show Built-in badge', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getAllByText('Built-in').length).toBeGreaterThanOrEqual(2);
    expect(screen.getByText('workiq')).toBeInTheDocument();
    expect(screen.getByText('powerbi')).toBeInTheDocument();
  });
});
