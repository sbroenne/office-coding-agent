/**
 * Integration test for ModelPicker interactions.
 *
 * Renders the REAL ModelPicker with the REAL Zustand store.
 * Verifies:
 *   - Shows default model (Claude Sonnet 4) as trigger label
 *   - Opens popover with models grouped by provider
 *   - Selecting a model updates activeModel in store
 *   - Mid-session model switching (calls switchModel when session active)
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

const mockSwitchModel = vi.fn();

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
    mockSwitchModel.mockClear();
    mockSwitchModel.mockResolvedValue(undefined);
    vi.clearAllMocks();
  });

  it('shows default model name as trigger label', () => {
    renderWithProviders(<ModelPicker onSwitchModel={mockSwitchModel} />);
    // Default is 'claude-sonnet-4.6' → 'Claude Sonnet 4.6'
    expect(screen.getByText('Claude Sonnet 4.6')).toBeInTheDocument();
  });

  it('opens popover and shows models grouped by provider', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ModelPicker onSwitchModel={mockSwitchModel} />);

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

  it('selecting a model with no active session updates activeModel in the store and closes popover', async () => {
    const user = userEvent.setup();
    renderWithProviders(<ModelPicker onSwitchModel={mockSwitchModel} />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    });

    await user.click(screen.getByText('GPT-4.1'));

    expect(useSettingsStore.getState().activeModel).toBe('gpt-4.1');
    expect(mockSwitchModel).not.toHaveBeenCalled();
  });

  it('selecting a model with an active session calls switchModel and updates store on success', async () => {
    const user = userEvent.setup();
    // switchModel in the real hook updates the store, so mock should too
    mockSwitchModel.mockImplementation(async (modelId: string) => {
      useSettingsStore.getState().setActiveModel(modelId);
    });

    renderWithProviders(<ModelPicker hasActiveSession onSwitchModel={mockSwitchModel} />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('GPT-4.1')).toBeInTheDocument();
    });

    await user.click(screen.getByText('GPT-4.1'));

    await waitFor(() => {
      expect(mockSwitchModel).toHaveBeenCalledWith('gpt-4.1');
      expect(useSettingsStore.getState().activeModel).toBe('gpt-4.1');
    });
  });

  it('shows (switching…) label while switching model mid-session', async () => {
    const user = userEvent.setup();
    // Make switchModel slow to capture loading state
    mockSwitchModel.mockImplementation(() => new Promise(resolve => setTimeout(resolve, 100)));

    renderWithProviders(<ModelPicker hasActiveSession onSwitchModel={mockSwitchModel} />);

    await user.click(screen.getByLabelText('Select model'));
    await waitFor(() => expect(screen.getByText('GPT-4.1')).toBeInTheDocument());
    
    await user.click(screen.getByText('GPT-4.1'));
    
    // Should show switching label
    await waitFor(() => {
      expect(screen.getByText('(switching…)')).toBeInTheDocument();
    });
  });

  it('shows error message and keeps popover open when switchModel fails', async () => {
    const user = userEvent.setup();
    mockSwitchModel.mockRejectedValue(new Error('Network error'));

    renderWithProviders(<ModelPicker hasActiveSession onSwitchModel={mockSwitchModel} />);

    await user.click(screen.getByLabelText('Select model'));
    await waitFor(() => expect(screen.getByText('GPT-4.1')).toBeInTheDocument());
    
    await user.click(screen.getByText('GPT-4.1'));

    await waitFor(() => {
      expect(screen.getByText('Network error')).toBeInTheDocument();
    });

    // Popover should still be open
    expect(screen.getByText('Anthropic')).toBeInTheDocument();
    
    // Store should NOT be updated on failure
    expect(useSettingsStore.getState().activeModel).toBe('claude-sonnet-4.6');
  });

  it('shows formatted model ID when activeModel does not match any available model', () => {
    // Bypass validation to simulate stale persisted data with an unknown model ID
    useSettingsStore.setState({ activeModel: 'nonexistent-model-id' });
    renderWithProviders(<ModelPicker onSwitchModel={mockSwitchModel} />);
    expect(screen.getByText('Nonexistent Model Id')).toBeInTheDocument();
  });

  it('shows "Connecting to Copilot…" when no models are available', async () => {
    const user = userEvent.setup();
    useSettingsStore.setState({ availableModels: null });
    renderWithProviders(<ModelPicker onSwitchModel={mockSwitchModel} />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText(/Connecting to Copilot/)).toBeInTheDocument();
    });
  });
});
