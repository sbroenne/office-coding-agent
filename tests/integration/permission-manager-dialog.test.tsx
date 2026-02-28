/**
 * Integration tests for PermissionManagerPanel.
 *
 * Renders the real component with the real Zustand store.
 * Tests allowAll toggle, rule display/removal, workingDirectory display,
 * and browse panel interaction.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { PermissionManagerPanel } from '@/components/PermissionManagerDialog';
import { usePermissionStore } from '@/stores/permissionStore';

// Mock fetch for /api/browse and /api/env
const mockFetchResponses: Record<string, unknown> = {};

beforeEach(() => {
  usePermissionStore.setState({
    allowAll: true,
    workingDirectory: null,
    rules: [],
  });

  // Default fetch mocks
  mockFetchResponses['/api/env'] = { cwd: '/Users/test/project', home: '/Users/test' };
  mockFetchResponses['/api/browse'] = {
    path: '/Users/test/project',
    parent: '/Users/test',
    dirs: ['src', 'tests', 'node_modules'],
  };

  vi.stubGlobal(
    'fetch',
    vi.fn((url: string) => {
      const urlStr = typeof url === 'string' ? url : String(url);
      // Match by path prefix
      const key = Object.keys(mockFetchResponses).find(k => urlStr.includes(k));
      const body = key ? mockFetchResponses[key] : {};
      return Promise.resolve({
        ok: true,
        json: () => Promise.resolve(body),
      });
    })
  );
});

describe('Integration: PermissionManagerPanel', () => {
  it('renders the Allow all toggle', () => {
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('Allow all')).toBeInTheDocument();
  });

  it('shows the current allowAll state as On', () => {
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('On')).toBeInTheDocument();
  });

  it('toggles allowAll when button is clicked', async () => {
    renderWithProviders(<PermissionManagerPanel />);
    const toggleBtn = screen.getByText('On');
    await userEvent.click(toggleBtn);
    expect(usePermissionStore.getState().allowAll).toBe(false);
    expect(screen.getByText('Off')).toBeInTheDocument();
  });

  it('shows "Not set" when workingDirectory is null', () => {
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('Not set')).toBeInTheDocument();
  });

  it('shows workingDirectory value when set', () => {
    usePermissionStore.setState({ workingDirectory: '/Users/test/project' });
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('/Users/test/project')).toBeInTheDocument();
  });

  it('shows Clear button only when workingDirectory is set', () => {
    const { unmount } = renderWithProviders(<PermissionManagerPanel />);
    expect(screen.queryByText('Clear')).not.toBeInTheDocument();
    unmount();

    usePermissionStore.setState({ workingDirectory: '/some/path' });
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('Clear')).toBeInTheDocument();
  });

  it('shows "No saved rules." when no rules exist', () => {
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('No saved rules.')).toBeInTheDocument();
  });

  it('renders rules when they exist', () => {
    usePermissionStore.setState({
      rules: [
        { id: 'write:/src', kind: 'write', pathPrefix: '/src' },
        { id: 'read:/tests', kind: 'read', pathPrefix: '/tests' },
      ],
    });
    renderWithProviders(<PermissionManagerPanel />);

    expect(screen.getByText('write')).toBeInTheDocument();
    expect(screen.getByText('/src')).toBeInTheDocument();
    expect(screen.getByText('read')).toBeInTheDocument();
    expect(screen.getByText('/tests')).toBeInTheDocument();
  });

  it('removes a rule when the Remove button is clicked', async () => {
    usePermissionStore.setState({
      rules: [{ id: 'write:/src', kind: 'write', pathPrefix: '/src' }],
    });
    renderWithProviders(<PermissionManagerPanel />);

    const removeBtn = screen.getByLabelText('Remove rule');
    await userEvent.click(removeBtn);

    expect(usePermissionStore.getState().rules).toHaveLength(0);
  });

  it('shows Clear all button when rules exist', () => {
    usePermissionStore.setState({
      rules: [{ id: 'write:/src', kind: 'write', pathPrefix: '/src' }],
    });
    renderWithProviders(<PermissionManagerPanel />);
    expect(screen.getByText('Clear all')).toBeInTheDocument();
  });

  it('clears all rules when Clear all is clicked', async () => {
    usePermissionStore.setState({
      rules: [
        { id: 'write:/src', kind: 'write', pathPrefix: '/src' },
        { id: 'read:/tests', kind: 'read', pathPrefix: '/tests' },
      ],
    });
    renderWithProviders(<PermissionManagerPanel />);

    await userEvent.click(screen.getByText('Clear all'));
    expect(usePermissionStore.getState().rules).toHaveLength(0);
  });

  it('opens the browse panel when Browse is clicked', async () => {
    renderWithProviders(<PermissionManagerPanel />);

    const browseBtn = screen.getByText('Browse');
    await userEvent.click(browseBtn);

    // Wait for the directory listing to appear
    await waitFor(() => {
      expect(screen.getByText('Select')).toBeInTheDocument();
    });
  });

  it('shows directory listing from /api/browse', async () => {
    renderWithProviders(<PermissionManagerPanel />);

    await userEvent.click(screen.getByText('Browse'));

    await waitFor(() => {
      expect(screen.getByText('src')).toBeInTheDocument();
      expect(screen.getByText('tests')).toBeInTheDocument();
      expect(screen.getByText('node_modules')).toBeInTheDocument();
    });
  });

  it('sets workingDirectory when Select is clicked in browse', async () => {
    renderWithProviders(<PermissionManagerPanel />);

    await userEvent.click(screen.getByText('Browse'));

    await waitFor(() => {
      expect(screen.getByText('Select')).toBeInTheDocument();
    });

    await userEvent.click(screen.getByText('Select'));

    expect(usePermissionStore.getState().workingDirectory).toBe('/Users/test/project');
  });
});
