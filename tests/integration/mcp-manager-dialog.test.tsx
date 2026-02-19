/**
 * Integration test: McpManagerDialog component.
 *
 * Renders the real McpManagerDialog with the real Zustand store.
 * Tests import, toggle, and remove flows without a live MCP server.
 */
import React from 'react';
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpManagerDialog } from '@/components/McpManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

function makeJsonFile(content: unknown, name = 'mcp.json'): File {
  return new File([JSON.stringify(content)], name, { type: 'application/json' });
}

const validMcpJson = {
  mcpServers: {
    'my-server': {
      url: 'https://example.com/mcp/sse',
      type: 'sse',
      description: 'Example MCP server',
    },
    'another-server': {
      url: 'https://example.com/mcp',
      type: 'http',
    },
  },
};

const OpenDialog: React.FC = () => {
  const [open, setOpen] = React.useState(true);
  return <McpManagerDialog open={open} onOpenChange={setOpen} />;
};

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: McpManagerDialog', () => {
  it('renders empty state with import button', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByRole('dialog', { name: 'MCP Servers' })).toBeInTheDocument();
    expect(screen.getByText(/No MCP servers configured/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Import mcp\.json/i })).toBeInTheDocument();
  });

  it('imports servers from a valid mcp.json file', async () => {
    renderWithProviders(<OpenDialog />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('my-server')).toBeInTheDocument();
    });

    expect(screen.getByText('another-server')).toBeInTheDocument();
    expect(screen.getByRole('status')).toHaveTextContent('Imported 2 servers from mcp.json');
  });

  it('shows the server description when available', async () => {
    renderWithProviders(<OpenDialog />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('Example MCP server')).toBeInTheDocument();
    });
  });

  it('shows error for invalid JSON', async () => {
    renderWithProviders(<OpenDialog />);

    const file = new File(['not-json'], 'mcp.json', { type: 'application/json' });
    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, file);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('shows error when all entries are stdio (no valid HTTP/SSE servers)', async () => {
    renderWithProviders(<OpenDialog />);

    const stdioOnly = { mcpServers: { srv: { command: 'node', args: ['server.js'] } } };
    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(stdioOnly));

    await waitFor(() => {
      expect(screen.getByRole('alert')).toHaveTextContent(/No valid HTTP\/SSE/i);
    });
  });

  it('servers are active (aria-pressed=true) by default after import', async () => {
    renderWithProviders(<OpenDialog />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('my-server')).toBeInTheDocument();
    });

    // Toggle button name includes description text — use regex
    const toggleButton = screen.getByRole('button', { name: /my-server/ });
    expect(toggleButton).toHaveAttribute('aria-pressed', 'true');
  });

  it('clicking a server button toggles it off', async () => {
    renderWithProviders(<OpenDialog />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('my-server')).toBeInTheDocument();
    });

    const toggleButton = screen.getByRole('button', { name: /my-server/ });
    await userEvent.click(toggleButton);

    expect(toggleButton).toHaveAttribute('aria-pressed', 'false');
    expect(useSettingsStore.getState().activeMcpServerNames).toEqual(['another-server']);
  });

  it('clicking a disabled server toggles it back on', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'srv', url: 'https://example.com/mcp', transport: 'http' },
    ]);
    useSettingsStore.setState({ activeMcpServerNames: [] });

    renderWithProviders(<OpenDialog />);

    // Toggle button name includes URL — use regex
    const toggleButton = screen.getByRole('button', { name: /^srv/ });
    expect(toggleButton).toHaveAttribute('aria-pressed', 'false');

    await userEvent.click(toggleButton);
    expect(toggleButton).toHaveAttribute('aria-pressed', 'true');
  });

  it('Remove button removes a server from the list and store', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'to-remove', url: 'https://example.com/mcp', transport: 'http' },
      { name: 'keep', url: 'https://example.com/keep', transport: 'http' },
    ]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('to-remove')).toBeInTheDocument();

    const removeButtons = screen.getAllByRole('button', { name: 'Remove' });
    await userEvent.click(removeButtons[0]);

    await waitFor(() => {
      expect(screen.queryByText('to-remove')).not.toBeInTheDocument();
    });
    expect(screen.getByText('keep')).toBeInTheDocument();
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(1);
  });
});
