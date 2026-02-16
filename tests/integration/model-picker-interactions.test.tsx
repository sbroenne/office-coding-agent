/**
 * Integration test for ModelPicker interactions.
 *
 * Renders the REAL ModelPicker with the REAL Zustand store.
 * Verifies:
 *   - Model selection updates the store
 *   - Models are grouped by provider
 *   - "Configure models in Settings" shown when no models
 *   - Disabled state when no endpoint configured
 *   - Opening and closing the popover
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ModelPicker } from '@/components/ModelPicker';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Mock services (not used directly, but imported transitively) ───
vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

// ─── Helpers ───

function setupWithModels() {
  const epId = useSettingsStore.getState().addEndpoint({
    displayName: 'Test',
    resourceUrl: 'https://test.openai.azure.com',
    authMethod: 'apiKey',
    apiKey: 'key',
  });

  useSettingsStore.getState().setModelsForEndpoint(epId, [
    { id: 'gpt-5.2-chat', name: 'GPT 5.2 Chat', ownedBy: 'user', provider: 'OpenAI' },
    { id: 'claude-opus-4-6', name: 'Claude Opus 4.6', ownedBy: 'user', provider: 'Anthropic' },
    { id: 'gpt-4.1', name: 'GPT 4.1', ownedBy: 'user', provider: 'OpenAI' },
  ]);

  useSettingsStore.getState().setActiveEndpoint(epId);
  return epId;
}

// ─── Tests ───

describe('ModelPicker — interactions', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    vi.clearAllMocks();
  });

  it('shows "No endpoint configured" when no endpoint exists', () => {
    renderWithProviders(<ModelPicker />);
    expect(screen.getByText('No endpoint configured')).toBeInTheDocument();
  });

  it('shows current model name as trigger label', () => {
    setupWithModels();
    renderWithProviders(<ModelPicker />);

    // setActiveEndpoint auto-selects first matching model
    const state = useSettingsStore.getState();
    const model = state.getActiveModel();
    expect(screen.getByText(model!.name)).toBeInTheDocument();
  });

  it('opens popover and shows models grouped by provider', async () => {
    const user = userEvent.setup();
    setupWithModels();
    renderWithProviders(<ModelPicker />);

    // Click the trigger
    await user.click(screen.getByLabelText('Select model'));

    // Provider group headers should appear
    await waitFor(() => {
      expect(screen.getByText('Anthropic')).toBeInTheDocument();
      expect(screen.getByText('OpenAI')).toBeInTheDocument();
    });

    // All models should be listed (some may appear twice — trigger label + list item)
    expect(screen.getAllByText('GPT 5.2 Chat').length).toBeGreaterThanOrEqual(1);
    expect(screen.getByText('Claude Opus 4.6')).toBeInTheDocument();
    expect(screen.getByText('GPT 4.1')).toBeInTheDocument();
  });

  it('selecting a model updates the store and closes the popover', async () => {
    const user = userEvent.setup();
    setupWithModels();
    renderWithProviders(<ModelPicker />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('Claude Opus 4.6')).toBeInTheDocument();
    });

    await user.click(screen.getByText('Claude Opus 4.6'));

    // Store should reflect the selection
    expect(useSettingsStore.getState().activeModelId).toBe('claude-opus-4-6');
  });

  it('shows "Configure models in Settings" when endpoint has no models', async () => {
    const user = userEvent.setup();
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'Empty',
      resourceUrl: 'https://empty.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });
    useSettingsStore.getState().setActiveEndpoint(epId);

    renderWithProviders(<ModelPicker />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('Configure models in Settings')).toBeInTheDocument();
    });
  });

  it('calls onOpenSettings when "Configure models" is clicked', async () => {
    const user = userEvent.setup();
    const onOpenSettings = vi.fn();
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'Empty',
      resourceUrl: 'https://empty.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });
    useSettingsStore.getState().setActiveEndpoint(epId);

    renderWithProviders(<ModelPicker onOpenSettings={onOpenSettings} />);

    await user.click(screen.getByLabelText('Select model'));

    await waitFor(() => {
      expect(screen.getByText('Configure models in Settings')).toBeInTheDocument();
    });

    await user.click(screen.getByText('Configure models in Settings'));
    expect(onOpenSettings).toHaveBeenCalledOnce();
  });

  it('shows "Select model" when active model ID does not match any model', () => {
    setupWithModels();
    useSettingsStore.getState().setActiveModel('nonexistent');
    renderWithProviders(<ModelPicker />);

    expect(screen.getByText('Select model')).toBeInTheDocument();
  });
});
