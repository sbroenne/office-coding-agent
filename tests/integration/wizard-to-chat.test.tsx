/**
 * Integration test: SetupWizard → App → ChatHeader + ModelPicker
 *
 * Renders the REAL App with REAL child components (SetupWizard, ChatHeader,
 * ChatPanel, ModelPicker). Only mocks external services (AI SDK, chat service).
 *
 * Verifies the critical user flow: completing the wizard actually results in
 * a usable chat UI with a model selected — not the broken "Select model" state.
 *
 * These tests would have caught the bugs:
 * - "No model selected after wizard completion"
 * - "Select model dropdown shows blank after wizard"
 * - "activeModelId is null despite wizard finishing"
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Mock ONLY external services, NOT components ───
const { mockDiscoverModels, mockValidateModel } = vi.hoisted(() => ({
  mockDiscoverModels: vi.fn(),
  mockValidateModel: vi.fn(),
}));

vi.mock('@/services/ai', () => ({
  discoverModels: mockDiscoverModels,
  validateModelDeployment: mockValidateModel,
  invalidateClient: vi.fn(),
}));

vi.mock('@/services/ai/aiClientFactory', () => ({
  getAIClient: vi.fn(),
  getAzureProvider: vi.fn(() => ({})),
  invalidateClient: vi.fn(),
}));

// Mock useOfficeChat — App calls this hook with provider/modelId/host
vi.mock('@/hooks/useOfficeChat', () => ({
  useOfficeChat: () => ({
    messages: [],
    sendMessage: vi.fn(),
    stop: vi.fn(),
    status: 'ready',
    setMessages: vi.fn(),
    error: undefined,
    clearError: vi.fn(),
    id: 'test',
  }),
}));

// Import App AFTER mocks
const { App } = await import('@/taskpane/App');

// ─── Helpers ───

/** Walk through the wizard: endpoint → auth → connect */
async function completeWizardSteps() {
  // Step 1: Endpoint
  await userEvent.type(
    screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
    'https://my-resource.openai.azure.com'
  );
  await userEvent.click(screen.getByRole('button', { name: 'Next' }));

  // Step 2: Auth
  await userEvent.type(screen.getByPlaceholderText('Enter your API key'), 'test-key-123');
  await userEvent.click(screen.getByRole('button', { name: 'Connect' }));
}

// ─── Tests ───

describe('Wizard → Chat integration', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    vi.clearAllMocks();
  });

  it('auto-discovery: wizard completion results in chat UI with model selected', async () => {
    mockDiscoverModels.mockResolvedValue({
      models: [
        {
          id: 'gpt-5.2-chat',
          name: 'gpt-5.2-chat',
          ownedBy: 'system',
          provider: 'OpenAI' as const,
        },
        { id: 'gpt-4.1', name: 'gpt-4.1', ownedBy: 'system', provider: 'OpenAI' as const },
      ],
      method: 'deployments',
    });

    renderWithProviders(<App />);

    // Wait for wizard to appear
    await waitFor(() => {
      expect(screen.getByText('Connect to Azure AI Foundry')).toBeInTheDocument();
    });

    // Walk through wizard
    await completeWizardSteps();

    // Should reach "done" step
    await waitFor(() => {
      expect(screen.getByText("You're all set!")).toBeInTheDocument();
    });

    // Click "Start Chatting"
    await userEvent.click(screen.getByRole('button', { name: 'Start Chatting' }));

    // App should now show the chat UI (NOT the wizard)
    await waitFor(() => {
      expect(screen.getByText('Excel')).toBeInTheDocument();
    });

    // ─── CRITICAL ASSERTIONS ───
    // The model picker should show a model name, NOT "Select model"
    expect(screen.queryByText('Select model')).not.toBeInTheDocument();
    expect(screen.getByText('gpt-5.2-chat')).toBeInTheDocument();

    // Store should have activeModelId set
    const state = useSettingsStore.getState();
    expect(state.activeEndpointId).not.toBeNull();
    expect(state.activeModelId).toBe('gpt-5.2-chat');
    expect(state.endpoints).toHaveLength(1);

    const models = state.getModelsForActiveEndpoint();
    expect(models).toHaveLength(2);
  });

  it('manual path: wizard completion results in chat UI with model selected', async () => {
    mockDiscoverModels.mockResolvedValue({
      models: [],
      method: 'manual',
    });
    // First call: default model validation succeeds
    // Second call: manually-added model validation succeeds
    mockValidateModel.mockResolvedValue(true);

    renderWithProviders(<App />);

    // Wait for wizard
    await waitFor(() => {
      expect(screen.getByText('Connect to Azure AI Foundry')).toBeInTheDocument();
    });

    await completeWizardSteps();

    // Manual step — default model should be pre-added
    await waitFor(() => {
      expect(screen.getByText('Add Your Models')).toBeInTheDocument();
    });
    expect(screen.getByText('gpt-5.2-chat')).toBeInTheDocument();

    // Click Finish (the pre-added default model is enough)
    await userEvent.click(screen.getByRole('button', { name: 'Finish' }));

    // Should reach done
    await waitFor(() => {
      expect(screen.getByText("You're all set!")).toBeInTheDocument();
    });

    await userEvent.click(screen.getByRole('button', { name: 'Start Chatting' }));

    // Chat UI should show with model selected
    await waitFor(() => {
      expect(screen.getByText('Excel')).toBeInTheDocument();
    });

    expect(screen.queryByText('Select model')).not.toBeInTheDocument();

    const state = useSettingsStore.getState();
    expect(state.activeModelId).not.toBeNull();
    expect(state.activeEndpointId).not.toBeNull();
  });

  it('manual path with custom model: model is selected after wizard', async () => {
    mockDiscoverModels.mockResolvedValue({
      models: [],
      method: 'manual',
    });
    // First call: default model validation FAILS
    // Second call: user's manually entered model validation succeeds
    mockValidateModel.mockResolvedValueOnce(false).mockResolvedValueOnce(true);

    renderWithProviders(<App />);

    await waitFor(() => {
      expect(screen.getByText('Connect to Azure AI Foundry')).toBeInTheDocument();
    });

    await completeWizardSteps();

    // Manual step — no pre-added model (default failed)
    await waitFor(() => {
      expect(screen.getByText('Add Your Models')).toBeInTheDocument();
    });

    // Type and add a custom model
    await userEvent.type(screen.getByPlaceholderText('e.g., gpt-4.1'), 'my-custom-model');
    await userEvent.click(screen.getByRole('button', { name: /Add/ }));

    await waitFor(() => {
      expect(screen.getByText('my-custom-model')).toBeInTheDocument();
    });

    await userEvent.click(screen.getByRole('button', { name: 'Finish' }));

    await waitFor(() => {
      expect(screen.getByText("You're all set!")).toBeInTheDocument();
    });

    await userEvent.click(screen.getByRole('button', { name: 'Start Chatting' }));

    await waitFor(() => {
      expect(screen.getByText('Excel')).toBeInTheDocument();
    });

    // The custom model should be selected
    expect(screen.queryByText('Select model')).not.toBeInTheDocument();

    const state = useSettingsStore.getState();
    expect(state.activeModelId).toBe('my-custom-model');
  });
});
