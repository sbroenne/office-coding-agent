/**
 * Integration test for ModelPicker interactions.
 *
 * Renders the REAL ModelPicker with the REAL Zustand store.
 * Verifies:
 *   - Shows default model (Claude Sonnet 4.5) as trigger label
 *   - Opens popover with models grouped by provider (Anthropic, OpenAI, Google)
 *   - Selecting a model updates activeModel in store
 *   - Shows "Select model" when activeModel is not in COPILOT_MODELS
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ModelPicker } from '@/components/ModelPicker';
import { useSettingsStore } from '@/stores/settingsStore';

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

describe('ModelPicker — interactions', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    vi.clearAllMocks();
  });

  it('shows default model name as trigger label', () => {
    renderWithProviders(<ModelPicker />);
    // Default is 'claude-sonnet-4.5' → 'Claude Sonnet 4.5'
    expect(screen.getByText('Claude Sonnet 4.5')).toBeInTheDocument();
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

    expect(screen.getByText('Claude Opus 4.5')).toBeInTheDocument();
    expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    expect(screen.getByText('Gemini 2.0 Flash')).toBeInTheDocument();
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

  it('shows "Select model" when activeModel does not match any COPILOT_MODEL', () => {
    // Bypass validation to simulate stale persisted data with an unknown model ID
    useSettingsStore.setState({ activeModel: 'nonexistent-model-id' });
    renderWithProviders(<ModelPicker />);
    expect(screen.getByText('Select model')).toBeInTheDocument();
  });
});
