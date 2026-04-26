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
  it('renders the skill picker button', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows empty state because CLI-owned plugin skills are not app-managed', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.getByText('No skills available yet.')).toBeInTheDocument();
  });

  it('still persists disabled skill names for SDK disabledSkills support', () => {
    useSettingsStore.getState().toggleSkill('example-skill');
    expect(useSettingsStore.getState().disabledSkillNames).toEqual(['example-skill']);
    expect(useSettingsStore.getState().isSkillEnabled('example-skill')).toBe(false);
  });
});
