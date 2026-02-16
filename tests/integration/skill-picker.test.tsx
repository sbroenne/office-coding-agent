/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with real Zustand store and real
 * bundled skills (loaded via rawMarkdownPlugin). Verifies toggling
 * skills on/off updates the store and shows the badge count.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SkillPicker } from '@/components/SkillPicker';
import { useSettingsStore } from '@/stores/settingsStore';
import { getSkills } from '@/services/skills';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillPicker', () => {
  it('renders skill button when bundled skills exist', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows badge with count when all skills are on (default null state)', () => {
    renderWithProviders(<SkillPicker />);
    const skills = getSkills();
    const button = screen.getByLabelText('Agent skills');
    // All skills on by default — badge shows count
    expect(within(button).getByText(String(skills.length))).toBeInTheDocument();
  });

  it('clicking a skill checkbox deactivates it from the store', async () => {
    renderWithProviders(<SkillPicker />);

    // Open the menu
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skills = getSkills();
    const skillName = skills[0].metadata.name;

    // Click to deactivate (was on by default)
    await userEvent.click(screen.getByText(skillName));

    // Store should materialize explicit list without that skill
    const names = useSettingsStore.getState().activeSkillNames;
    expect(Array.isArray(names)).toBe(true);
    expect(names).not.toContain(skillName);
  });

  it('re-activating a skill adds it back', async () => {
    // Pre-deactivate a skill
    const skills = getSkills();
    const skillName = skills[0].metadata.name;
    useSettingsStore.getState().toggleSkill(skillName); // deactivate from null

    renderWithProviders(<SkillPicker />);

    // Open menu and re-activate
    await userEvent.click(screen.getByLabelText('Agent skills'));
    await userEvent.click(screen.getByText(skillName));

    expect(useSettingsStore.getState().activeSkillNames).toContain(skillName);
  });

  it('shows skill description as secondary content', async () => {
    renderWithProviders(<SkillPicker />);

    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skills = getSkills();
    const firstSentence = skills[0].metadata.description.split('.')[0];
    expect(screen.getByText(firstSentence)).toBeInTheDocument();
  });
});
