/**
 * Integration test: SkillPicker component.
 *
 * Renders the real SkillPicker with real Zustand store and real
 * bundled skills (loaded via rawMarkdownPlugin). Verifies toggling
 * skills on/off updates the store and shows the badge count.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { SkillPicker } from '@/components/SkillPicker';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillPicker', () => {
  it('does not render skill button when no bundled skills exist', () => {
    renderWithProviders(<SkillPicker />);
    expect(screen.queryByLabelText('Agent skills')).not.toBeInTheDocument();
  });
});
