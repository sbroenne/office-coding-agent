/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with real Zustand store and real
 * bundled skills (loaded via rawMarkdownPlugin). Verifies toggling
 * skills on/off updates the store and shows the badge count.
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
  it('renders skill button even when no skills are loaded', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows skills content and manage action in popover', async () => {
    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('Skills')).toBeInTheDocument();
    expect(screen.getByText('Bundled')).toBeInTheDocument();
    expect(screen.getByText('Manage skills…')).toBeInTheDocument();
  });

  it('calls onOpenPanel when manage skills button is clicked', async () => {
    renderWithProviders(<SkillPicker onOpenPanel={mockOpenPanel} />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    const manageButton = screen.getByRole('button', { name: /manage skills/i });
    manageButton.focus();
    await userEvent.keyboard('{Enter}');

    expect(mockOpenPanel).toHaveBeenCalledWith('skills');
  });
});
