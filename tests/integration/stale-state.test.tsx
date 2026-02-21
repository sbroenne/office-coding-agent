/**
 * Tests for stale state / hydration scenarios with the new Copilot model store.
 *
 * Verifies that setActiveModel validates against availableModels when set,
 * and that ModelPicker renders correctly with stale state.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { useSettingsStore } from '@/stores/settingsStore';
import { ModelPicker } from '@/components/ModelPicker';
import type { CopilotModel } from '@/types';

const TEST_MODELS: CopilotModel[] = [
  { id: 'claude-sonnet-4.6', name: 'Claude Sonnet 4.6', provider: 'Anthropic' },
  { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
];

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

describe('Stale state scenarios', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    useSettingsStore.setState({ availableModels: null });
  });

  describe('setActiveModel', () => {
    it('accepts a valid model ID from availableModels', () => {
      useSettingsStore.getState().setAvailableModels(TEST_MODELS);
      const validId = TEST_MODELS[1].id;
      useSettingsStore.getState().setActiveModel(validId);
      expect(useSettingsStore.getState().activeModel).toBe(validId);
    });

    it('ignores unknown model IDs when availableModels is set', () => {
      useSettingsStore.getState().setAvailableModels(TEST_MODELS);
      const before = useSettingsStore.getState().activeModel;
      useSettingsStore.getState().setActiveModel('some-stale-model-from-last-session');
      expect(useSettingsStore.getState().activeModel).toBe(before);
    });

    it('reset restores activeModel to default', () => {
      useSettingsStore.getState().setActiveModel('gpt-4.1');
      useSettingsStore.getState().reset();
      expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.6');
    });
  });

  describe('ModelPicker with stale store state', () => {
    it('shows "Select model" when activeModel is stale (not in available models)', () => {
      useSettingsStore.setState({ activeModel: 'deleted-model-from-old-session' });
      renderWithProviders(<ModelPicker />);
      // When no models are loaded, displays formatted model ID
      expect(screen.getByText('Deleted Model From Old Session')).toBeInTheDocument();
    });

    it('shows model name when activeModel is a valid model', () => {
      useSettingsStore.setState({
        activeModel: 'gpt-4.1',
        availableModels: TEST_MODELS,
      });
      renderWithProviders(<ModelPicker />);
      expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    });

    it('shows default model ID after reset (before models load)', () => {
      useSettingsStore.getState().reset();
      renderWithProviders(<ModelPicker />);
      // Before models load, displays formatted model ID
      expect(screen.getByText('Claude Sonnet 4.6')).toBeInTheDocument();
    });
  });
});
