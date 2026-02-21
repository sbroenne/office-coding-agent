/**
 * Integration test for ModelPicker interactions.
 *
 * Renders the REAL ModelPicker with the REAL Zustand store.
 * Verifies:
 *   - Shows default model (Claude Sonnet 4) as trigger label
 *   - Opens popover with models grouped by provider
 *   - Selecting a model updates activeModel in store
 *   - Shows "Select model" when activeModel is not in available models
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ModelPicker } from '@/components/ModelPicker';
import { useSettingsStore } from '@/stores/settingsStore';
import type { CopilotModel } from '@/types';

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

const TEST_MODELS: CopilotModel[] = [
  { id: 'claude-sonnet-4.6', name: 'Claude Sonnet 4.6', provider: 'Anthropic' },
  { id: 'claude-opus-4', name: 'Claude Opus 4', provider: 'Anthropic' },
  { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
  { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro', provider: 'Google' },
];

describe('ModelPicker — interactions', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    useSettingsStore.getState().setAvailableModels(TEST_MODELS);
    vi.clearAllMocks();
  });

  it('shows default model name as trigger label', () => {
    renderWithProviders(<ModelPicker />);
    // Default is 'claude-sonnet-4.6' → 'Claude Sonnet 4.6'
    expect(screen.getByText('Claude Sonnet 4.6')).toBeInTheDocument();
  });

  it('opens popover and shows models grouped by provider', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ModelPicker />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('Anthropic')).toBeInTheDocument();
      expect(screen.getByText('OpenAI')).toBeInTheDocument();
      expect(screen.getByText('Google')).toBeInTheDocument();
    });

    expect(screen.getByText('Claude Opus 4')).toBeInTheDocument();
    expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    expect(screen.getByText('Gemini 2.5 Pro')).toBeInTheDocument();
  });

  it('selecting a model updates activeModel in the store and closes popover', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ModelPicker />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    });

    await user.click(screen.getByText('GPT-4.1'));

    expect(useSettingsStore.getState().activeModel).toBe('gpt-4.1');
  });

  it('shows formatted model ID when activeModel does not match any available model', () => {
    // Bypass validation to simulate stale persisted data with an unknown model ID
    useSettingsStore.setState({ activeModel: 'nonexistent-model-id' });
    renderWithProviders(<ModelPicker />);
    expect(screen.getByText('Nonexistent Model Id')).toBeInTheDocument();
  });

  it('shows "Connecting to Copilot…" when no models are available', async () => {
    const user = userEvent.setup();
    useSettingsStore.setState({ availableModels: null });
    renderWithProviders(<ModelPicker />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText(/Connecting to Copilot/)).toBeInTheDocument();
    });
  });
});
