/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with real Zustand store and real
 * bundled skills. Verifies toggling skills on/off updates the store
 * and the badge reflects the enabled count.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SkillPicker } from '@/components/SkillPicker';
import { useSettingsStore } from '@/stores/settingsStore';

const mockOpenPanel = vi.fn();

beforeEach(() => {
  useSettingsStore.getState().reset();
  mockOpenPanel.mockClear();
});

describe('Integration: SkillPicker', () => {
  it('renders skill button', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows skills and manage action in popover', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.getByText('Skills')).toBeInTheDocument();
    expect(screen.getByText('Manage plugins…')).toBeInTheDocument();
  });

  it('skills are enabled (aria-pressed=true) by default', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const buttons = screen.getAllByRole('button', { pressed: true });
    expect(buttons.length).toBeGreaterThan(0);
  });

  it('clicking a skill toggles it off and stores it as disabled', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skillButtons = screen.getAllByRole('button', { pressed: true });
    const firstSkill = skillButtons[0];
    const skillName = firstSkill.querySelector('.font-medium')?.textContent ?? '';

    await userEvent.click(firstSkill);

    expect(useSettingsStore.getState().disabledSkillNames).toContain(skillName);
  });

  it('clicking a disabled skill re-enables it', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skillButtons = screen.getAllByRole('button', { pressed: true });
    const firstSkill = skillButtons[0];

    // Disable then re-enable
    await userEvent.click(firstSkill);
    await userEvent.click(firstSkill);

    expect(firstSkill).toHaveAttribute('aria-pressed', 'true');
  });

  it('calls onOpenPanel when manage plugins button is clicked', async () => {
    renderWithProviders(<SkillPicker onOpenPanel={mockOpenPanel} />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const manageButton = screen.getByRole('button', { name: /manage plugins/i });
    manageButton.focus();
    await userEvent.keyboard('{Enter}');

    expect(mockOpenPanel).toHaveBeenCalledWith('plugins');
  });
});
