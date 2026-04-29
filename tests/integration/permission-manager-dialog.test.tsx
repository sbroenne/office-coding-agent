/**
 * Integration tests for PermissionManagerPanel.
 *
 * Renders the real component and verifies the SDK/CLI-style controls exposed by the panel.
 */
import { describe, it, expect, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { PermissionManagerPanel } from '@/components/PermissionManagerDialog';

describe('Integration: PermissionManagerPanel', () => {
  it('renders the SDK approve-all control', () => {
    renderWithProviders(
      <PermissionManagerPanel
        approveAll={false}
        onSetApproveAll={vi.fn()}
        onResetSessionApprovals={vi.fn()}
      />
    );

    expect(screen.getByText('Allow all')).toBeInTheDocument();
    expect(screen.getByText('Off')).toBeInTheDocument();
    expect(screen.getByText(/Copilot CLI session setting/i)).toBeInTheDocument();
  });

  it('calls onSetApproveAll with the next value', async () => {
    const onSetApproveAll = vi.fn();

    renderWithProviders(
      <PermissionManagerPanel
        approveAll={false}
        onSetApproveAll={onSetApproveAll}
        onResetSessionApprovals={vi.fn()}
      />
    );

    await userEvent.click(screen.getByText('Off'));
    expect(onSetApproveAll).toHaveBeenCalledWith(true);
  });

  it('renders approve-all as enabled', () => {
    renderWithProviders(
      <PermissionManagerPanel
        approveAll
        onSetApproveAll={vi.fn()}
        onResetSessionApprovals={vi.fn()}
      />
    );

    expect(screen.getByText('On')).toBeInTheDocument();
  });

  it('calls onResetSessionApprovals', async () => {
    const onResetSessionApprovals = vi.fn();

    renderWithProviders(
      <PermissionManagerPanel
        approveAll={false}
        onSetApproveAll={vi.fn()}
        onResetSessionApprovals={onResetSessionApprovals}
      />
    );

    await userEvent.click(screen.getByText('Reset session approvals'));
    expect(onResetSessionApprovals).toHaveBeenCalled();
  });
});
