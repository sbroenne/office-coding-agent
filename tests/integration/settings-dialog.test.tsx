/**
 * Integration test for SettingsDialog.
 *
 * SettingsDialog is now a simple dialog with a title and close button.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SettingsDialog } from '@/components/SettingsDialog';
import { useSettingsStore } from '@/stores/settingsStore';

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
  inferProvider: vi.fn().mockReturnValue('OpenAI'),
}));

describe('SettingsDialog', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  it('opens and shows Settings heading in controlled mode', () => {
    renderWithProviders(<SettingsDialog open={true} />);
    expect(screen.getByRole('heading', { name: 'Settings' })).toBeInTheDocument();
  });

  it('Close button calls onOpenChange(false)', async () => {
    const onOpenChange = vi.fn();
    renderWithProviders(<SettingsDialog open={true} onOpenChange={onOpenChange} />);

    await userEvent.click(screen.getByRole('button', { name: 'Close' }));
    expect(onOpenChange).toHaveBeenCalledWith(false);
  });

  it('does not show heading when closed', () => {
    renderWithProviders(<SettingsDialog open={false} />);
    expect(screen.queryByRole('heading', { name: 'Settings' })).not.toBeInTheDocument();
  });

  it('shows description text', () => {
    renderWithProviders(<SettingsDialog open={true} />);
    expect(screen.getByText('Configure runtime preferences.')).toBeInTheDocument();
  });
});
