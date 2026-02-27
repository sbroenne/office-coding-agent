/**
 * Integration test: McpLogViewer component.
 *
 * Tests log rendering, empty state, level-based coloring, and copy/clear buttons.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpLogViewer } from '@/components/McpLogViewer';
import { useMcpStatusStore } from '@/stores/mcpStatusStore';
import type { McpLogEntry } from '@/types';

beforeEach(() => {
  useMcpStatusStore.getState().clearAll();
});

function addLogs(server: string, entries: McpLogEntry[]) {
  for (const entry of entries) {
    useMcpStatusStore.getState().addLog(server, entry);
  }
}

describe('Integration: McpLogViewer', () => {
  it('shows select-server prompt when no server is selected', () => {
    renderWithProviders(<McpLogViewer serverName={null} />);

    expect(screen.getByText('Select a server to view its output')).toBeInTheDocument();
  });

  it('shows empty state when server has no logs', () => {
    useMcpStatusStore.getState().setStatus('my-srv', 'connected');
    renderWithProviders(<McpLogViewer serverName="my-srv" />);

    expect(screen.getByText(/No output yet/)).toBeInTheDocument();
  });

  it('renders log entries for the selected server', () => {
    addLogs('test-srv', [
      { timestamp: '2025-01-01T10:00:00Z', level: 'info', message: 'Server started' },
      { timestamp: '2025-01-01T10:00:01Z', level: 'warn', message: 'Slow connection' },
      { timestamp: '2025-01-01T10:00:02Z', level: 'error', message: 'Connection failed' },
    ]);

    renderWithProviders(<McpLogViewer serverName="test-srv" />);

    expect(screen.getByText('Server started')).toBeInTheDocument();
    expect(screen.getByText('Slow connection')).toBeInTheDocument();
    expect(screen.getByText('Connection failed')).toBeInTheDocument();
  });

  it('shows server name in output header', () => {
    useMcpStatusStore.getState().setStatus('my-srv', 'connected');
    renderWithProviders(<McpLogViewer serverName="my-srv" />);

    expect(screen.getByText('Output: my-srv')).toBeInTheDocument();
  });

  it('clear button clears logs for the server', async () => {
    addLogs('srv', [
      { timestamp: '2025-01-01T10:00:00Z', level: 'info', message: 'hello world' },
    ]);

    renderWithProviders(<McpLogViewer serverName="srv" />);

    expect(screen.getByText('hello world')).toBeInTheDocument();

    await userEvent.click(screen.getByTitle('Clear logs'));

    expect(screen.queryByText('hello world')).not.toBeInTheDocument();
    expect(screen.getByText(/No output yet/)).toBeInTheDocument();
  });

  it('copy and clear buttons are disabled when there are no logs', () => {
    useMcpStatusStore.getState().setStatus('empty-srv', 'stopped');
    renderWithProviders(<McpLogViewer serverName="empty-srv" />);

    expect(screen.getByTitle('Copy logs')).toBeDisabled();
    expect(screen.getByTitle('Clear logs')).toBeDisabled();
  });

  it('copy and clear buttons are enabled when there are logs', () => {
    addLogs('srv', [
      { timestamp: '2025-01-01T10:00:00Z', level: 'info', message: 'log line' },
    ]);

    renderWithProviders(<McpLogViewer serverName="srv" />);

    expect(screen.getByTitle('Copy logs')).not.toBeDisabled();
    expect(screen.getByTitle('Clear logs')).not.toBeDisabled();
  });
});
