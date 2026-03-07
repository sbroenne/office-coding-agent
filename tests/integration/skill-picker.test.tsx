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
    // Section header is now "Built-in" (renamed from "Skills" to support plugin sections)
    expect(screen.getByText('Built-in')).toBeInTheDocument();
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

  it('shows "Plugin" section when plugin skills are added to store', async () => {
    const store = useSettingsStore.getState();
    store.setPluginSkills([
      {
        metadata: {
          name: 'SPT IQ Preflight',
          description: 'Preflight account assessment skill.',
          version: '1.0.0',
          tags: [],
          hosts: [],
        },
        content: 'Skill body content here.',
      },
    ]);

    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    expect(screen.getByText('Plugin')).toBeInTheDocument();
    expect(screen.getByText('SPT IQ Preflight')).toBeInTheDocument();
  });

  it('plugin skill is enabled by default and can be toggled', async () => {
    const store = useSettingsStore.getState();
    store.setPluginSkills([
      {
        metadata: {
          name: 'Test Plugin Skill',
          description: 'A test plugin skill.',
          version: '1.0.0',
          tags: [],
          hosts: [],
        },
        content: 'Content.',
      },
    ]);

    renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));

    // Find the plugin skill by its visible label text
    const skillLabel = screen.getByText('Test Plugin Skill');
    const skillBtn = skillLabel.closest('button')!;
    expect(skillBtn).toHaveAttribute('aria-pressed', 'true');

    await userEvent.click(skillBtn);

    expect(useSettingsStore.getState().disabledSkillNames).toContain('Test Plugin Skill');
  });

  it('re-renders when store.setPluginSkills() is called after initial render', async () => {
    const { rerender } = renderWithProviders(<SkillPicker />);
    await userEvent.click(screen.getByLabelText('Agent skills'));
    expect(screen.queryByText('Plugin')).not.toBeInTheDocument();

    // Simulate plugin.skills notification arriving after session start
    useSettingsStore.getState().setPluginSkills([
      {
        metadata: {
          name: 'Dynamic Plugin Skill',
          description: 'Arrives via notification.',
          version: '1.0.0',
          tags: [],
          hosts: [],
        },
        content: 'Content.',
      },
    ]);
    rerender(<SkillPicker />);

    expect(screen.getByText('Plugin')).toBeInTheDocument();
    expect(screen.getByText('Dynamic Plugin Skill')).toBeInTheDocument();
  });
});
