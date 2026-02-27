/**
 * Integration test: McpAddServerForm component.
 *
 * Tests form validation, transport-specific fields, add and edit modes.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpAddServerForm } from '@/components/McpAddServerForm';

const defaultProps = {
  existingNames: new Set<string>(),
  onSubmit: vi.fn(),
  onCancel: vi.fn(),
};

beforeEach(() => {
  vi.clearAllMocks();
});

describe('Integration: McpAddServerForm', () => {
  it('renders add mode by default', () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    expect(screen.getByText('Add Server', { selector: 'h4' })).toBeInTheDocument();
    expect(screen.getByPlaceholderText('my-server')).toBeInTheDocument();
    expect(screen.getByText('stdio')).toBeInTheDocument();
    expect(screen.getByText('http')).toBeInTheDocument();
    expect(screen.getByText('sse')).toBeInTheDocument();
  });

  it('shows Command and Arguments fields for stdio transport', () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    expect(screen.getByPlaceholderText('npx')).toBeInTheDocument();
    expect(screen.getByPlaceholderText('-y @microsoft/workiq mcp')).toBeInTheDocument();
    expect(screen.queryByPlaceholderText('https://example.com/mcp')).not.toBeInTheDocument();
  });

  it('shows URL and Headers fields for http transport', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    await userEvent.click(screen.getByLabelText('http'));

    expect(screen.getByPlaceholderText('https://example.com/mcp')).toBeInTheDocument();
    expect(screen.queryByPlaceholderText('npx')).not.toBeInTheDocument();
  });

  it('shows URL and Headers fields for sse transport', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    await userEvent.click(screen.getByLabelText('sse'));

    expect(screen.getByPlaceholderText('https://example.com/mcp')).toBeInTheDocument();
  });

  it('shows error when name is empty', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    // Clear the name field and submit
    const nameInput = screen.getByPlaceholderText('my-server');
    await userEvent.clear(nameInput);
    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(screen.getByText('Name is required')).toBeInTheDocument();
    expect(defaultProps.onSubmit).not.toHaveBeenCalled();
  });

  it('shows error when name already exists', async () => {
    renderWithProviders(
      <McpAddServerForm
        {...defaultProps}
        existingNames={new Set(['taken-name'])}
      />
    );

    const nameInput = screen.getByPlaceholderText('my-server');
    await userEvent.clear(nameInput);
    await userEvent.type(nameInput, 'taken-name');
    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(screen.getByText('A server with this name already exists')).toBeInTheDocument();
    expect(defaultProps.onSubmit).not.toHaveBeenCalled();
  });

  it('shows error when stdio command is empty', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    const nameInput = screen.getByPlaceholderText('my-server');
    await userEvent.type(nameInput, 'my-srv');

    // Clear the command field
    const cmdInput = screen.getByPlaceholderText('npx');
    await userEvent.clear(cmdInput);

    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(screen.getByText('Command is required for stdio transport')).toBeInTheDocument();
    expect(defaultProps.onSubmit).not.toHaveBeenCalled();
  });

  it('shows error when http URL is empty', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    const nameInput = screen.getByPlaceholderText('my-server');
    await userEvent.type(nameInput, 'my-srv');
    await userEvent.click(screen.getByLabelText('http'));

    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(screen.getByText('URL is required for HTTP/SSE transport')).toBeInTheDocument();
    expect(defaultProps.onSubmit).not.toHaveBeenCalled();
  });

  it('submits valid stdio config', async () => {
    const onSubmit = vi.fn();
    renderWithProviders(
      <McpAddServerForm {...defaultProps} onSubmit={onSubmit} />
    );

    await userEvent.type(screen.getByPlaceholderText('my-server'), 'test-server');
    // Command field defaults to 'npx', add args
    await userEvent.type(screen.getByPlaceholderText('-y @microsoft/workiq mcp'), '-y my-tool mcp');

    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(onSubmit).toHaveBeenCalledWith(expect.objectContaining({
      name: 'test-server',
      transport: 'stdio',
      command: 'npx',
      args: ['-y', 'my-tool', 'mcp'],
    }));
  });

  it('submits valid http config', async () => {
    const onSubmit = vi.fn();
    renderWithProviders(
      <McpAddServerForm {...defaultProps} onSubmit={onSubmit} />
    );

    await userEvent.type(screen.getByPlaceholderText('my-server'), 'http-srv');
    await userEvent.click(screen.getByLabelText('http'));
    await userEvent.type(screen.getByPlaceholderText('https://example.com/mcp'), 'https://api.example.com/mcp');

    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(onSubmit).toHaveBeenCalledWith(expect.objectContaining({
      name: 'http-srv',
      transport: 'http',
      url: 'https://api.example.com/mcp',
    }));
  });

  it('shows error for invalid headers JSON', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    await userEvent.type(screen.getByPlaceholderText('my-server'), 'test');
    await userEvent.click(screen.getByLabelText('sse'));
    await userEvent.type(screen.getByPlaceholderText('https://example.com/mcp'), 'https://x.com/mcp');
    await userEvent.type(screen.getByPlaceholderText('{"Authorization": "Bearer ..."}'), 'not-json');

    await userEvent.click(screen.getByRole('button', { name: 'Add Server' }));

    expect(screen.getByText('Headers must be valid JSON')).toBeInTheDocument();
  });

  it('calls onCancel when Cancel button is clicked', async () => {
    renderWithProviders(<McpAddServerForm {...defaultProps} />);

    await userEvent.click(screen.getByRole('button', { name: 'Cancel' }));

    expect(defaultProps.onCancel).toHaveBeenCalled();
  });

  it('renders edit mode with pre-filled values', () => {
    renderWithProviders(
      <McpAddServerForm
        {...defaultProps}
        editMode
        initial={{
          name: 'existing-server',
          transport: 'http',
          url: 'https://example.com/mcp',
          description: 'A test server',
        }}
      />
    );

    expect(screen.getByText('Edit Server')).toBeInTheDocument();
    expect(screen.getByDisplayValue('existing-server')).toBeInTheDocument();
    expect(screen.getByDisplayValue('A test server')).toBeInTheDocument();
    expect(screen.getByDisplayValue('https://example.com/mcp')).toBeInTheDocument();
  });

  it('name field is readonly in edit mode', () => {
    renderWithProviders(
      <McpAddServerForm
        {...defaultProps}
        editMode
        initial={{
          name: 'locked-name',
          transport: 'stdio',
          command: 'npx',
        }}
      />
    );

    const nameInput = screen.getByDisplayValue('locked-name');
    expect(nameInput).toHaveAttribute('readOnly');
  });

  it('edit mode skips name uniqueness check', async () => {
    const onSubmit = vi.fn();
    renderWithProviders(
      <McpAddServerForm
        {...defaultProps}
        editMode
        existingNames={new Set(['edit-me'])}
        initial={{
          name: 'edit-me',
          transport: 'http',
          url: 'https://old.com/mcp',
        }}
        onSubmit={onSubmit}
      />
    );

    // Change the URL
    const urlInput = screen.getByDisplayValue('https://old.com/mcp');
    await userEvent.clear(urlInput);
    await userEvent.type(urlInput, 'https://new.com/mcp');

    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    // Should submit successfully even though name exists (we're editing it)
    expect(onSubmit).toHaveBeenCalledWith(expect.objectContaining({
      name: 'edit-me',
      url: 'https://new.com/mcp',
    }));
  });
});
