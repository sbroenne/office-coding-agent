/**
 * Integration test for ModelManager component.
 *
 * Renders the REAL ModelManager with the REAL Zustand store.
 * Only mocks validateModelDeployment (network call).
 *
 * Verifies:
 *   - Empty state with "Add a model" prompt
 *   - Add model flow: type name → validate → appears in list
 *   - Validation error display when model can't be reached
 *   - Duplicate prevention
 *   - Rename model (inline edit)
 *   - Remove model
 *   - Provider badge display
 *   - Cancel add form
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { ModelManager } from '@/components/ModelManager';
import { useSettingsStore } from '@/stores/settingsStore';

// ─── Mock only the network call ───
const mockValidate = vi.fn();
vi.mock('@/services/ai', () => ({
  validateModelDeployment: (...args: unknown[]) => mockValidate(...args),
  inferProvider: vi.fn().mockReturnValue('OpenAI'),
}));

// ─── Helpers ───

function setupEndpoint() {
  const epId = useSettingsStore.getState().addEndpoint({
    displayName: 'Test Endpoint',
    resourceUrl: 'https://test.openai.azure.com',
    authMethod: 'apiKey',
    apiKey: 'test-key',
  });
  return epId;
}

function setupEndpointWithModel() {
  const epId = setupEndpoint();
  useSettingsStore.getState().addModel(epId, {
    id: 'gpt-4.1',
    name: 'gpt-4.1',
    ownedBy: 'user',
    provider: 'OpenAI',
  });
  return epId;
}

// ─── Tests ───

describe('ModelManager', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    vi.clearAllMocks();
  });

  it('shows empty state when no models are configured', () => {
    const epId = setupEndpoint();
    renderWithProviders(<ModelManager endpointId={epId} />);

    expect(screen.getByText('No models configured.')).toBeInTheDocument();
    expect(screen.getByText('Add a model')).toBeInTheDocument();
  });

  it('shows the add model form when clicking the + button', async () => {
    const user = userEvent.setup();
    const epId = setupEndpointWithModel();
    renderWithProviders(<ModelManager endpointId={epId} />);

    // The + button is the icon button next to "Models" header
    const header = screen.getByText('Models').closest('div')!;
    const addBtn = header.querySelector('button')!;
    await user.click(addBtn);

    expect(screen.getByLabelText('Deployment name')).toBeInTheDocument();
    expect(screen.getByPlaceholderText('e.g., gpt-5.2-chat')).toBeInTheDocument();
  });

  it('validates and adds a model successfully', async () => {
    const user = userEvent.setup();
    mockValidate.mockResolvedValue(true);
    const epId = setupEndpoint();
    renderWithProviders(<ModelManager endpointId={epId} />);

    // Click "Add a model" link in empty state
    await user.click(screen.getByText('Add a model'));

    // Type model name and submit
    await user.type(screen.getByPlaceholderText('e.g., gpt-5.2-chat'), 'gpt-5.2-chat');
    await user.click(screen.getByRole('button', { name: /Add/ }));

    // Model should appear in the list
    await waitFor(() => {
      expect(screen.getByText('gpt-5.2-chat')).toBeInTheDocument();
    });

    // Validation was called
    expect(mockValidate).toHaveBeenCalledOnce();
  });

  it('shows validation error when model cannot be reached', async () => {
    const user = userEvent.setup();
    mockValidate.mockResolvedValue(false);
    const epId = setupEndpoint();
    renderWithProviders(<ModelManager endpointId={epId} />);

    await user.click(screen.getByText('Add a model'));
    await user.type(screen.getByPlaceholderText('e.g., gpt-5.2-chat'), 'bad-model');
    await user.click(screen.getByRole('button', { name: /Add/ }));

    await waitFor(() => {
      expect(screen.getByText(/Could not reach "bad-model"/)).toBeInTheDocument();
    });
  });

  it('prevents adding duplicate models', async () => {
    const user = userEvent.setup();
    mockValidate.mockResolvedValue(true);
    const epId = setupEndpointWithModel(); // already has 'gpt-4.1'
    renderWithProviders(<ModelManager endpointId={epId} />);

    // Click the + button
    const header = screen.getByText('Models').closest('div')!;
    await user.click(header.querySelector('button')!);

    await user.type(screen.getByPlaceholderText('e.g., gpt-5.2-chat'), 'gpt-4.1');
    await user.click(screen.getByRole('button', { name: /Add/ }));

    await waitFor(() => {
      expect(screen.getByText(/"gpt-4.1" is already added/)).toBeInTheDocument();
    });

    // Validation should NOT have been called (duplicate check happens first)
    expect(mockValidate).not.toHaveBeenCalled();
  });

  it('cancels the add form without saving', async () => {
    const user = userEvent.setup();
    const epId = setupEndpointWithModel();
    renderWithProviders(<ModelManager endpointId={epId} />);

    const header = screen.getByText('Models').closest('div')!;
    await user.click(header.querySelector('button')!);
    expect(screen.getByPlaceholderText('e.g., gpt-5.2-chat')).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: 'Cancel' }));
    expect(screen.queryByPlaceholderText('e.g., gpt-5.2-chat')).not.toBeInTheDocument();
  });

  it('removes a model from the list', async () => {
    const user = userEvent.setup();
    const epId = setupEndpointWithModel();
    renderWithProviders(<ModelManager endpointId={epId} />);

    expect(screen.getByText('gpt-4.1')).toBeInTheDocument();

    // The remove button is the second icon button (after rename) in the model row
    const buttons = screen.getAllByRole('button');
    const removeBtn = buttons.find(btn => btn.querySelector('[class*="lucide-trash"]') !== null)
      ?? buttons[buttons.length - 1]; // fallback to last button in the model row
    await user.click(removeBtn);

    await waitFor(() => {
      expect(screen.queryByText('gpt-4.1')).not.toBeInTheDocument();
    });
    expect(screen.getByText('No models configured.')).toBeInTheDocument();
  });

  it('renames a model via inline edit', async () => {
    const user = userEvent.setup();
    const epId = setupEndpointWithModel();
    renderWithProviders(<ModelManager endpointId={epId} />);

    // Find the rename (pencil) button — it's the first icon button in the model row
    const modelRow = screen.getByText('gpt-4.1').closest('div[class*="rounded-md"]')!;
    const editButtons = modelRow.querySelectorAll('button');
    // First icon button is rename, second is remove
    await user.click(editButtons[0]);

    // Input should appear with current name
    const input = screen.getByDisplayValue('gpt-4.1');
    expect(input).toBeInTheDocument();

    // Clear and type new name
    await user.clear(input);
    await user.type(input, 'GPT 4.1 Turbo');

    // Save with the first button in the edit row (check icon)
    const editRow = input.closest('div[class*="rounded-md"]')!;
    const saveBtn = editRow.querySelectorAll('button')[0];
    await user.click(saveBtn);

    // New name should be displayed
    await waitFor(() => {
      expect(screen.getByText('GPT 4.1 Turbo')).toBeInTheDocument();
    });

    // Store should reflect the update
    const models = useSettingsStore.getState().getModelsForEndpoint(epId);
    expect(models[0].name).toBe('GPT 4.1 Turbo');
    expect(models[0].id).toBe('gpt-4.1'); // ID unchanged
  });

  it('shows provider badge for each model', () => {
    const epId = setupEndpointWithModel();
    renderWithProviders(<ModelManager endpointId={epId} />);

    expect(screen.getByText('OpenAI')).toBeInTheDocument();
  });

  it('submits model on Enter key press', async () => {
    const user = userEvent.setup();
    mockValidate.mockResolvedValue(true);
    const epId = setupEndpoint();
    renderWithProviders(<ModelManager endpointId={epId} />);

    await user.click(screen.getByText('Add a model'));
    const input = screen.getByPlaceholderText('e.g., gpt-5.2-chat');
    await user.type(input, 'gpt-5.2-chat{enter}');

    await waitFor(() => {
      expect(screen.getByText('gpt-5.2-chat')).toBeInTheDocument();
    });
  });
});
