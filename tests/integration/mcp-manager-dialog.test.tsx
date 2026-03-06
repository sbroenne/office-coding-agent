/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel. Servers are fetched from /api/mcp-servers.
 * Each server has an enable/disable toggle. No JSON upload; no Remove buttons.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpManagerPanel } from '@/components/McpManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

const MOCK_SERVERS = [
  { name: 'workiq', description: 'WorkIQ MCP server', transport: 'http' as const, url: 'http://localhost:3001' },
  { name: 'powerbi', description: 'Power BI MCP server', transport: 'http' as const, url: 'http://localhost:3002' },
];

beforeEach(() => {
  useSettingsStore.getState().reset();
  vi.spyOn(globalThis, 'fetch').mockResolvedValue({
    ok: true,
    json: async () => ({ servers: MOCK_SERVERS }),
  } as Response);
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe('Integration: McpManagerPanel', () => {
  it('fetches and renders servers from /api/mcp-servers', async () => {
    renderWithProviders(<McpManagerPanel />);

    await waitFor(() => {
      expect(screen.getByText('workiq')).toBeInTheDocument();
      expect(screen.getByText('powerbi')).toBeInTheDocument();
    });
  });

  it('shows empty state when API returns no servers', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({ servers: [] }),
    } as Response);

    renderWithProviders(<McpManagerPanel />);

    await waitFor(() => {
      expect(screen.getByText(/No MCP servers available/)).toBeInTheDocument();
    });
  });

  it('each server has a toggle button', async () => {
    renderWithProviders(<McpManagerPanel />);

    await waitFor(() => {
      expect(screen.getByLabelText('Toggle workiq')).toBeInTheDocument();
      expect(screen.getByLabelText('Toggle powerbi')).toBeInTheDocument();
    });
  });

  it('clicking toggle disables a server and updates the store', async () => {
    renderWithProviders(<McpManagerPanel />);

    const toggle = await screen.findByLabelText('Toggle workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'true');

    await userEvent.click(toggle);

    expect(useSettingsStore.getState().disabledMcpServerNames).toContain('workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'false');
  });

  it('clicking toggle again re-enables a server', async () => {
    renderWithProviders(<McpManagerPanel />);

    const toggle = await screen.findByLabelText('Toggle workiq');
    await userEvent.click(toggle);
    await userEvent.click(toggle);

    expect(useSettingsStore.getState().disabledMcpServerNames).not.toContain('workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'true');
  });

  it('does not show a Remove button for any server', async () => {
    renderWithProviders(<McpManagerPanel />);

    await waitFor(() => {
      expect(screen.getByText('workiq')).toBeInTheDocument();
    });

    expect(screen.queryByLabelText(/Remove/)).not.toBeInTheDocument();
  });

  it('does not show a JSON import button', async () => {
    renderWithProviders(<McpManagerPanel />);

    await waitFor(() => {
      expect(screen.getByText('workiq')).toBeInTheDocument();
    });

    expect(screen.queryByLabelText(/Import MCP/)).not.toBeInTheDocument();
  });
});
