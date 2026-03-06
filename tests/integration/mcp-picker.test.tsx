/**
 * Integration test: McpPicker component.
 *
 * Verifies the McpPicker fetches servers from /api/mcp-servers and lists
 * them in the popover with enable/disable toggles.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpPicker } from '@/components/McpPicker';
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

describe('Integration: McpPicker', () => {
  it('renders a trigger button with aria-label MCP servers', () => {
    renderWithProviders(<McpPicker />);
    expect(screen.getByRole('button', { name: 'MCP servers' })).toBeInTheDocument();
  });

  it('opens popover and shows workiq and powerbi from /api/mcp-servers', async () => {
    renderWithProviders(<McpPicker />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    await waitFor(() => {
      expect(screen.getByText('workiq')).toBeInTheDocument();
      expect(screen.getByText('powerbi')).toBeInTheDocument();
    });
  });

  it('shows "No MCP servers available" when API returns empty list', async () => {
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({ servers: [] }),
    } as Response);

    renderWithProviders(<McpPicker />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    await waitFor(() => {
      expect(screen.getByText(/No MCP servers available/)).toBeInTheDocument();
    });
  });

  it('toggling a server calls toggleMcpServer in the store', async () => {
    renderWithProviders(<McpPicker />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    const toggle = await screen.findByRole('button', { name: 'workiq' });
    expect(toggle).toHaveAttribute('aria-pressed', 'true');

    await userEvent.click(toggle);

    expect(useSettingsStore.getState().disabledMcpServerNames).toContain('workiq');
    expect(toggle).toHaveAttribute('aria-pressed', 'false');
  });

  it('shows a badge when not all servers are enabled', async () => {
    useSettingsStore.getState().toggleMcpServer('workiq'); // disable workiq

    renderWithProviders(<McpPicker />);

    // Badge shows enabled count (1 of 2 enabled)
    await waitFor(() => {
      expect(screen.getByLabelText('1 of 2 MCP servers enabled')).toBeInTheDocument();
    });
  });
});
