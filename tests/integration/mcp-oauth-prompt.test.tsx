import { describe, expect, it, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { McpOAuthPrompt } from '@/components/McpOAuthPrompt';

describe('Integration: McpOAuthPrompt', () => {
  it('opens a foreground prompt and submits a login hint', async () => {
    const onSignIn = vi.fn().mockResolvedValue('janesmith@microsoft.com');
    const onClose = vi.fn();

    renderWithProviders(
      <McpOAuthPrompt
        request={{ serverName: 'powerbi', reason: 'chat-required', blocking: true }}
        onClose={onClose}
        onSignIn={onSignIn}
      />
    );

    const dialog = screen.getByRole('dialog', { name: 'Sign in to powerbi' });
    expect(dialog).toHaveTextContent(
      'Office Coding Agent needs you to sign in to powerbi before this MCP server can be used.'
    );

    await userEvent.type(screen.getByLabelText('Username or alias'), 'janesmith');
    await userEvent.click(screen.getByRole('button', { name: 'Sign in' }));

    await waitFor(() => {
      expect(onSignIn).toHaveBeenCalledWith('powerbi', 'janesmith');
      expect(onClose).toHaveBeenCalled();
    });
  });

  it('pre-fills alias from a Microsoft login hint and supports switch account', async () => {
    const onSignIn = vi.fn().mockResolvedValue('janesmith@microsoft.com');

    renderWithProviders(
      <McpOAuthPrompt
        request={{
          serverName: 'powerbi',
          reason: 'switch',
          defaultLoginHint: 'janesmith@microsoft.com',
        }}
        onClose={() => undefined}
        onSignIn={onSignIn}
      />
    );

    expect(screen.getByRole('dialog', { name: 'Switch account for powerbi' })).toBeInTheDocument();
    expect(screen.getByLabelText('Username or alias')).toHaveValue('janesmith');
    await userEvent.click(screen.getByRole('button', { name: 'Switch account' }));

    await waitFor(() => {
      expect(onSignIn).toHaveBeenCalledWith('powerbi', 'janesmith');
    });
  });
});
