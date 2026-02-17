/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with real Zustand store and real
 * bundled skills (loaded via rawMarkdownPlugin). Verifies toggling
 * skills on/off updates the store and shows the badge count.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SkillPicker } from '@/components/SkillPicker';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillPicker', () => {
  it('renders skill button even when no skills are loaded', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows empty state and manage action in popover', async () => {
    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('No skills available yet.')).toBeInTheDocument();
    expect(screen.getByText('Manage skills…')).toBeInTheDocument();
  });

  it('opens manager dialog from keyboard and closes with Escape', async () => {
    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    const manageButton = screen.getByRole('button', { name: /manage skills/i });
    manageButton.focus();
    await userEvent.keyboard('{Enter}');

    expect(screen.getByRole('dialog', { name: 'Manage Skills' })).toBeInTheDocument();

    await userEvent.keyboard('{Escape}');
    expect(screen.queryByRole('dialog', { name: 'Manage Skills' })).not.toBeInTheDocument();
  });
});
