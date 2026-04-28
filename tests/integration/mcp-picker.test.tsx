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
import { useMcpStatusStore } from '@/stores/mcpStatusStore';

const MOCK_SERVERS = [
  {
    name: 'workiq',
    description: 'WorkIQ MCP server',
    transport: 'http' as const,
    url: 'http://localhost:3001',
  },
  {
    name: 'powerbi',
    description: 'Power BI MCP server',
    transport: 'http' as const,
    url: 'http://localhost:3002',
  },
];

beforeEach(() => {
  useSettingsStore.getState().reset();
  useMcpStatusStore.getState().clearAll();
  vi.spyOn(globalThis, 'fetch').mockResolvedValue({
    ok: true,
    json: async () => ({ servers: MOCK_SERVERS }),
  } as Response);
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe('Integration: McpPicker', () => {
  it('renders a trigger button with aria-label MCP servers', async () => {
    renderWithProviders(<McpPicker />);
    expect(screen.getByRole('button', { name: 'MCP servers' })).toBeInTheDocument();
    await waitFor(() => {
      expect(globalThis.fetch).toHaveBeenCalled();
    });
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

  it('deduplicates by endpoint: plugin server with same url as bundled is dropped', async () => {
    // A plugin declares a server at the same URL as the bundled powerbi server
    // but under a different name — the bundled one must win and only one entry appears.
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({
        servers: [
          // Bundled workiq (stdio) and powerbi (http) — kept
          {
            name: 'workiq',
            description: 'WorkIQ (bundled)',
            transport: 'stdio' as const,
            command: 'npx',
            args: ['-y', '@microsoft/workiq', 'mcp'],
          },
          {
            name: 'powerbi',
            description: 'Power BI (bundled)',
            transport: 'http' as const,
            url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
          },
          // Plugin with different name but same URL as powerbi — server already dropped by dedup
          // custom-plugin has a unique URL — kept
          {
            name: 'custom-plugin',
            description: 'From plugin: my-plugin',
            transport: 'http' as const,
            url: 'http://localhost:3003',
          },
        ],
      }),
    } as Response);

    renderWithProviders(<McpPicker />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    await waitFor(() => {
      expect(screen.getByText('workiq')).toBeInTheDocument();
      expect(screen.getByText('powerbi')).toBeInTheDocument();
      expect(screen.getByText('custom-plugin')).toBeInTheDocument();
    });

    // Only one powerbi entry (the bundled one)
    expect(screen.getAllByText('powerbi')).toHaveLength(1);
  });

  it('shows a sign-in action for auth-required remote MCP servers', async () => {
    const initiateOAuth = vi.fn().mockResolvedValue(undefined);
    useMcpStatusStore.getState().setStatus('powerbi', 'needs-auth');

    renderWithProviders(<McpPicker onInitiateOAuth={initiateOAuth} />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    await userEvent.click(await screen.findByRole('button', { name: 'Sign in' }));

    expect(initiateOAuth).toHaveBeenCalledWith('powerbi', undefined);
  });

  it('matches auth status when SDK events use normalized server keys', async () => {
    const initiateOAuth = vi.fn().mockResolvedValue(undefined);
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({
        servers: [
          {
            name: 'Power BI MCP',
            description: 'Power BI MCP server',
            transport: 'http' as const,
            url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
          },
        ],
      }),
    } as Response);
    useMcpStatusStore.getState().setStatus('power_bi_mcp', 'needs-auth');

    renderWithProviders(<McpPicker onInitiateOAuth={initiateOAuth} />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    await userEvent.click(await screen.findByRole('button', { name: 'Sign in' }));

    expect(initiateOAuth).toHaveBeenCalledWith('Power BI MCP', undefined);
  });

  it('shows retry sign-in for failed remote MCP servers', async () => {
    const initiateOAuth = vi.fn().mockResolvedValue(undefined);
    useMcpStatusStore.getState().setStatus('powerbi', 'failed', 'Authentication required');

    renderWithProviders(<McpPicker onInitiateOAuth={initiateOAuth} />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    expect(await screen.findByText('Failed — Authentication required')).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: 'Retry sign in' }));

    expect(initiateOAuth).toHaveBeenCalledWith('powerbi', undefined);
  });

  it('surfaces signed-in alias and switch-account action for connected OAuth servers', async () => {
    const openPrompt = vi.fn();
    useMcpStatusStore.getState().setOAuthState('powerbi', 'connected', 'janesmith@microsoft.com');
    useMcpStatusStore.getState().setStatus('powerbi', 'connected');

    renderWithProviders(<McpPicker onOpenOAuthPrompt={openPrompt} />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));

    expect(await screen.findByText('Signed in as janesmith@microsoft.com')).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: 'Switch account' }));

    expect(openPrompt).toHaveBeenCalledWith({
      serverName: 'powerbi',
      reason: 'switch',
      defaultLoginHint: 'janesmith@microsoft.com',
    });
  });

  it('scopes OAuth sign-in errors to the server that failed', async () => {
    const initiateOAuth = vi.fn().mockRejectedValue(new Error('OAuth failed'));
    vi.spyOn(globalThis, 'fetch').mockResolvedValue({
      ok: true,
      json: async () => ({
        servers: [
          {
            name: 'powerbi',
            description: 'Power BI MCP server',
            transport: 'http' as const,
            url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
          },
          {
            name: 'custom-api',
            description: 'Custom API server',
            transport: 'http' as const,
            url: 'https://api.example.com/mcp',
          },
        ],
      }),
    } as Response);
    useMcpStatusStore.getState().setStatus('powerbi', 'needs-auth');
    useMcpStatusStore.getState().setStatus('custom-api', 'needs-auth');

    renderWithProviders(<McpPicker onInitiateOAuth={initiateOAuth} />);
    await userEvent.click(screen.getByRole('button', { name: 'MCP servers' }));
    await userEvent.click((await screen.findAllByRole('button', { name: 'Sign in' }))[0]);

    expect(await screen.findByText('OAuth failed')).toBeInTheDocument();
    expect(screen.getAllByText('OAuth failed')).toHaveLength(1);
  });
});
