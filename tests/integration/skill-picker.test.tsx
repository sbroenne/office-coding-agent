/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with plugin skills from the Zustand store.
 * Verifies toggling skills on/off, host filtering, badge count, and
 * dynamic skill arrival via plugin.skills notifications.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SkillPicker } from '@/components/SkillPicker';
import { useSettingsStore } from '@/stores/settingsStore';

// detectOfficeHost is called by SkillPicker — mock it per test where needed
vi.mock('@/services/office/host', () => ({
  detectOfficeHost: vi.fn(() => 'excel'),
}));

import { detectOfficeHost } from '@/services/office/host';
const mockDetectOfficeHost = vi.mocked(detectOfficeHost);

beforeEach(() => {
  useSettingsStore.getState().reset();
  mockDetectOfficeHost.mockReturnValue('excel');
});

const EXCEL_SKILL = {
  metadata: { name: 'Excel Helper', description: 'Excel skill.', version: '1.0.0', tags: [], hosts: ['excel' as const] },
  content: 'Excel skill content.',
};
const PPT_SKILL = {
  metadata: { name: 'PPT Helper', description: 'PPT skill.', version: '1.0.0', tags: [], hosts: ['powerpoint' as const] },
  content: 'PPT skill content.',
};
const UNIVERSAL_SKILL = {
  metadata: { name: 'Universal', description: 'Works everywhere.', version: '1.0.0', tags: [], hosts: [] },
  content: 'Universal content.',
};

describe('Integration: SkillPicker', () => {
  it('renders skill button', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.getByLabelText('Agent skills')).toBeInTheDocument();
  });

  it('shows empty state when no plugin skills', async () => {
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.getByText('No skills available yet.')).toBeInTheDocument();
  });

  it('shows skills section when plugin skills are added to store', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL]);

    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('Skills')).toBeInTheDocument();
    expect(screen.getByText('Excel Helper')).toBeInTheDocument();
  });

  it('filters out skills that do not match the current host', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL, PPT_SKILL, UNIVERSAL_SKILL]);
    // host = excel — should see Excel Helper + Universal, NOT PPT Helper
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('Excel Helper')).toBeInTheDocument();
    expect(screen.getByText('Universal')).toBeInTheDocument();
    expect(screen.queryByText('PPT Helper')).not.toBeInTheDocument();
  });

  it('shows all skills with empty hosts on any host', async () => {
    mockDetectOfficeHost.mockReturnValue('word');
    useSettingsStore.getState().setPluginSkills([UNIVERSAL_SKILL, EXCEL_SKILL]);

    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('Universal')).toBeInTheDocument();
    expect(screen.queryByText('Excel Helper')).not.toBeInTheDocument();
  });

  it('all skills are enabled (aria-pressed=true) by default', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL, UNIVERSAL_SKILL]);
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const buttons = screen.getAllByRole('button', { pressed: true });
    expect(buttons.length).toBeGreaterThanOrEqual(2);
  });

  it('clicking a skill toggles it off and stores it as disabled', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL]);
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skillBtn = screen.getByText('Excel Helper').closest('button')!;
    await userEvent.click(skillBtn);

    expect(useSettingsStore.getState().disabledSkillNames).toContain('Excel Helper');
    expect(skillBtn).toHaveAttribute('aria-pressed', 'false');
  });

  it('clicking a disabled skill re-enables it', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL]);
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    const skillBtn = screen.getByText('Excel Helper').closest('button')!;
    await userEvent.click(skillBtn); // disable
    await userEvent.click(skillBtn); // re-enable

    expect(useSettingsStore.getState().disabledSkillNames).not.toContain('Excel Helper');
    expect(skillBtn).toHaveAttribute('aria-pressed', 'true');
  });

  it('badge shows disabled count when a skill is toggled off', async () => {
    useSettingsStore.getState().setPluginSkills([EXCEL_SKILL, UNIVERSAL_SKILL]);
    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    // Disable one skill — badge should appear showing 1 of 2 enabled
    const skillBtn = screen.getByText('Excel Helper').closest('button')!;
    await userEvent.click(skillBtn);

    expect(screen.getByLabelText('1 of 2 skills enabled')).toBeInTheDocument();
  });

  it('re-renders when store.setPluginSkills() is called after initial render', async () => {
    const { rerender } = renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.queryByText('Skills')).not.toBeInTheDocument();

    useSettingsStore.getState().setPluginSkills([
      {
        metadata: { name: 'Dynamic Plugin Skill', description: 'Arrives via notification.', version: '1.0.0', tags: [], hosts: [] },
        content: 'Content.',
      },
    ]);
    rerender(<SkillPicker />);

    expect(screen.getByText('Skills')).toBeInTheDocument();
    expect(screen.getByText('Dynamic Plugin Skill')).toBeInTheDocument();
  });
});
