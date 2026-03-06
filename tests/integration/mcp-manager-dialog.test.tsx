/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel. Bundled servers are shown with status
 * indicators and toggle buttons. Plugin servers are installed via the Copilot CLI Plugin Hub.
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
});
