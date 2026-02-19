/**
 * Tests for stale state / hydration scenarios with the new Copilot model store.
 *
 * Verifies that setActiveModel ignores unknown model IDs, accepts valid
 * COPILOT_MODELS IDs, and that ModelPicker renders correctly with stale state.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { useSettingsStore } from '@/stores/settingsStore';
import { ModelPicker } from '@/components/ModelPicker';
import { COPILOT_MODELS } from '@/types';

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

describe('Stale state scenarios', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  describe('setActiveModel', () => {
    it('accepts a valid COPILOT_MODELS ID', () => {
      const validId = COPILOT_MODELS[1].id;
      useSettingsStore.getState().setActiveModel(validId);
      expect(useSettingsStore.getState().activeModel).toBe(validId);
    });

    it('ignores unknown model IDs (leaves activeModel unchanged)', () => {
      const before = useSettingsStore.getState().activeModel;
      useSettingsStore.getState().setActiveModel('some-stale-model-from-last-session');
      expect(useSettingsStore.getState().activeModel).toBe(before);
    });

    it('reset restores activeModel to default', () => {
      useSettingsStore.getState().setActiveModel('gpt-4.1');
      useSettingsStore.getState().reset();
      expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.5');
    });
  });

  describe('ModelPicker with stale store state', () => {
    it('shows "Select model" when activeModel is stale (not in COPILOT_MODELS)', () => {
      useSettingsStore.setState({ activeModel: 'deleted-model-from-old-session' });
      renderWithProviders(<ModelPicker />);
      expect(screen.getByText('Select model')).toBeInTheDocument();
    });

    it('shows model name when activeModel is a valid COPILOT_MODEL', () => {
      useSettingsStore.setState({ activeModel: 'gpt-4.1' });
      renderWithProviders(<ModelPicker />);
      expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    });

    it('shows default model name after reset', () => {
      useSettingsStore.getState().reset();
      renderWithProviders(<ModelPicker />);
      expect(screen.getByText('Claude Sonnet 4.5')).toBeInTheDocument();
    });
  });
});
