/**
 * Integration test: McpManagerPanel component.
 *
 * Renders the real McpManagerPanel with the real Zustand store.
 * Tests import, toggle, and remove flows without a live MCP server.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpManagerPanel } from '@/components/McpManagerDialog';
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

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: McpManagerPanel', () => {
  it('renders with bundled servers and import button', () => {
    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText('workiq')).toBeInTheDocument();
    expect(screen.getByText('Built-in')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Import/i })).toBeInTheDocument();
  });

  it('imports servers from a valid mcp.json file', async () => {
    renderWithProviders(<McpManagerPanel />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('my-server')).toBeInTheDocument();
    });

    expect(screen.getByText('another-server')).toBeInTheDocument();
    expect(screen.getByRole('status')).toHaveTextContent('Imported 2 servers from mcp.json');
  });

  it('shows the server description when available', async () => {
    renderWithProviders(<McpManagerPanel />);

    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(validMcpJson));

    await waitFor(() => {
      expect(screen.getByText('Example MCP server')).toBeInTheDocument();
    });
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

  it('imports stdio (npx) entries from mcp.json successfully', async () => {
    renderWithProviders(<McpManagerPanel />);

    const stdioOnly = { mcpServers: { srv: { command: 'node', args: ['server.js'] } } };
    const fileInput = screen.getByLabelText('Import mcp.json file');
    await userEvent.upload(fileInput, makeJsonFile(stdioOnly));

    await waitFor(() => {
      expect(screen.getByRole('status')).toHaveTextContent('Imported 1 server from mcp.json.');
    });
    expect(screen.getByText('srv')).toBeInTheDocument();
    expect(screen.getByText('node server.js')).toBeInTheDocument();
  });

  it('imported servers show Remove button while bundled servers do not', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'to-remove', url: 'https://example.com/mcp', transport: 'http' },
      { name: 'keep', url: 'https://example.com/keep', transport: 'http' },
    ]);

    renderWithProviders(<McpManagerPanel />);

    const removeButtons = screen.getAllByTitle('Remove');
    expect(removeButtons.length).toBe(2);
  });

  it('Remove button removes a server from the list and store', async () => {
    useSettingsStore.getState().importMcpServers([
      { name: 'to-remove', url: 'https://example.com/mcp', transport: 'http' },
      { name: 'keep', url: 'https://example.com/keep', transport: 'http' },
    ]);

    renderWithProviders(<McpManagerPanel />);

    expect(screen.getByText('to-remove')).toBeInTheDocument();

    const removeButtons = screen.getAllByTitle('Remove');
    await userEvent.click(removeButtons[0]);

    await waitFor(() => {
      expect(screen.queryByText('to-remove')).not.toBeInTheDocument();
    });
    expect(screen.getByText('keep')).toBeInTheDocument();
    expect(useSettingsStore.getState().importedMcpServers).toHaveLength(1);
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

    expect(screen.getByText('Built-in')).toBeInTheDocument();
    expect(screen.getByText('workiq')).toBeInTheDocument();
  });
});
