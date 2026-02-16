/**
 * Integration test for SettingsDialog CRUD operations.
 *
 * Renders the REAL SettingsDialog (with real ModelManager) — only mocks
 * external services. Verifies the complete endpoint add/delete/activate
 * flow works, including auto-validation on save.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import { SettingsDialog } from '@/components/SettingsDialog';
import { useSettingsStore } from '@/stores/settingsStore';

// Mock only external services
const { mockValidateModel } = vi.hoisted(() => ({
  mockValidateModel: vi.fn().mockResolvedValue(true),
}));

vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: mockValidateModel,
  inferProvider: vi.fn().mockReturnValue('OpenAI'),
}));

describe('SettingsDialog CRUD', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
    mockValidateModel.mockResolvedValue(true);
  });

  // ── Opening / Closing ──────────────────────────────────

  it('opens and closes in controlled mode', async () => {
    const onOpenChange = vi.fn();
    const { rerender } = renderWithProviders(
      <SettingsDialog open={true} onOpenChange={onOpenChange} />
    );

    expect(screen.getByText('Settings')).toBeInTheDocument();

    // Click Close
    await userEvent.click(screen.getByRole('button', { name: 'Close' }));
    expect(onOpenChange).toHaveBeenCalledWith(false);

    // Re-render closed
    rerender(<SettingsDialog open={false} onOpenChange={onOpenChange} />);
  });

  // ── Empty state ────────────────────────────────────────

  it('shows empty state when no endpoints exist', () => {
    renderWithProviders(<SettingsDialog open={true} />);

    expect(screen.getByText(/No endpoints configured/)).toBeInTheDocument();
  });

  // ── Add Endpoint ────────────────────────────────────────

  it('adds an endpoint via the form (auto-tests on save)', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    // Click "Add Endpoint" button (in empty state)
    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));

    // Form should appear
    expect(screen.getByText('Add Endpoint')).toBeInTheDocument();

    // Fill out the form
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Test Foundry');
    await userEvent.type(
      screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
      'https://test-resource.openai.azure.com'
    );
    await userEvent.type(screen.getByPlaceholderText('Enter API key'), 'test-key-abc');

    // Click Save — auto-tests connection
    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    // Wait for save to complete (async validation)
    await waitFor(() => {
      expect(useSettingsStore.getState().endpoints).toHaveLength(1);
    });

    // Connection summary should show hostname
    expect(screen.getByText(/test-resource\.openai\.azure\.com/)).toBeInTheDocument();

    // Store should have the endpoint
    const state = useSettingsStore.getState();
    expect(state.endpoints[0].displayName).toBe('Test Foundry');
    expect(state.endpoints[0].apiKey).toBe('test-key-abc');
  });

  it('shows error when save fails validation', async () => {
    mockValidateModel.mockResolvedValueOnce(false);

    renderWithProviders(<SettingsDialog open={true} />);

    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));

    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Bad EP');
    await userEvent.type(
      screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
      'https://bad.openai.azure.com'
    );
    await userEvent.type(screen.getByPlaceholderText('Enter API key'), 'wrong-key');

    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    // Should show error, NOT save
    expect(await screen.findByText(/Could not connect/)).toBeInTheDocument();
    expect(useSettingsStore.getState().endpoints).toHaveLength(0);
  });

  it('disables Save when required fields are empty', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));

    // Save should be disabled with empty fields
    expect(screen.getByRole('button', { name: 'Save' })).toBeDisabled();

    // Fill only display name — still disabled
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Test');
    expect(screen.getByRole('button', { name: 'Save' })).toBeDisabled();
  });

  it('Cancel closes the form without saving', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Should Not Save');
    await userEvent.click(screen.getByRole('button', { name: 'Cancel' }));

    // Form should be gone, endpoint should not exist
    expect(screen.queryByPlaceholderText('My AI Foundry Resource')).not.toBeInTheDocument();
    expect(useSettingsStore.getState().endpoints).toHaveLength(0);
  });

  // ── Delete Endpoint ──────────────────────────────────────

  it('deletes an endpoint from edit mode', async () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'To Delete',
      resourceUrl: 'https://delete-me.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(<SettingsDialog open={true} />);

    expect(screen.getByText('To Delete')).toBeInTheDocument();

    // Enter edit mode
    await userEvent.click(screen.getByRole('button', { name: 'Edit endpoint' }));

    // Click delete
    await userEvent.click(screen.getByRole('button', { name: 'Delete endpoint' }));

    // Confirm the deletion
    expect(screen.getByText(/Delete this endpoint/)).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: 'Confirm delete' }));

    // Endpoint should be gone
    expect(screen.queryByText('To Delete')).not.toBeInTheDocument();
    expect(useSettingsStore.getState().endpoints).toHaveLength(0);
  });

  // ── Multiple endpoints ────────────────────────────────────

  it('handles adding multiple endpoints', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    // Add first endpoint
    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Endpoint A');
    await userEvent.type(
      screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
      'https://a.openai.azure.com'
    );
    await userEvent.type(screen.getByPlaceholderText('Enter API key'), 'key-a');
    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    await waitFor(() => {
      expect(useSettingsStore.getState().endpoints).toHaveLength(1);
    });

    // Add second endpoint
    await userEvent.click(screen.getByRole('button', { name: /Add another endpoint/ }));
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Endpoint B');
    await userEvent.type(
      screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
      'https://b.openai.azure.com'
    );
    await userEvent.type(screen.getByPlaceholderText('Enter API key'), 'key-b');
    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    await waitFor(() => {
      expect(useSettingsStore.getState().endpoints).toHaveLength(2);
    });
  });

  // ── Endpoint URL normalization ─────────────────────────

  it('strips trailing slashes from endpoint URL', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));
    await userEvent.type(screen.getByPlaceholderText('My AI Foundry Resource'), 'Trimmed');
    await userEvent.type(
      screen.getByPlaceholderText('https://your-resource.openai.azure.com'),
      'https://trimmed.openai.azure.com///'
    );
    await userEvent.type(screen.getByPlaceholderText('Enter API key'), 'key');
    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    await waitFor(() => {
      expect(useSettingsStore.getState().endpoints).toHaveLength(1);
    });

    const ep = useSettingsStore.getState().endpoints[0];
    expect(ep.resourceUrl).toBe('https://trimmed.openai.azure.com');
  });

  // ── Delete cascade ──────────────────────────────────────

  it('deleting the active endpoint cascades: removes models and falls back to remaining endpoint', async () => {
    const ep1 = useSettingsStore.getState().addEndpoint({
      displayName: 'Primary',
      resourceUrl: 'https://primary.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'k1',
    });
    const ep2 = useSettingsStore.getState().addEndpoint({
      displayName: 'Fallback',
      resourceUrl: 'https://fallback.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'k2',
    });

    useSettingsStore
      .getState()
      .setModelsForEndpoint(ep1, [
        { id: 'gpt-4.1', name: 'gpt-4.1', ownedBy: 'user', provider: 'OpenAI' },
      ]);

    expect(useSettingsStore.getState().activeEndpointId).toBe(ep1);
    expect(useSettingsStore.getState().getModelsForEndpoint(ep1)).toHaveLength(1);

    renderWithProviders(<SettingsDialog open={true} />);

    // Enter edit mode, then delete
    await userEvent.click(screen.getByRole('button', { name: 'Edit endpoint' }));
    await userEvent.click(screen.getByRole('button', { name: 'Delete endpoint' }));

    expect(screen.getByText(/Delete this endpoint/)).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: 'Confirm delete' }));

    // Store should have only Fallback, which should now be active
    const state = useSettingsStore.getState();
    expect(state.endpoints).toHaveLength(1);
    expect(state.endpoints[0].displayName).toBe('Fallback');
    expect(state.activeEndpointId).toBe(ep2);
    expect(state.getModelsForEndpoint(ep1)).toEqual([]);
  });

  it('deleting the only endpoint returns to empty state', async () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'Solo',
      resourceUrl: 'https://solo.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(<SettingsDialog open={true} />);
    expect(screen.getByText('Solo')).toBeInTheDocument();

    // Enter edit mode, delete
    await userEvent.click(screen.getByRole('button', { name: 'Edit endpoint' }));
    await userEvent.click(screen.getByRole('button', { name: 'Delete endpoint' }));

    expect(screen.getByText(/Delete this endpoint/)).toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: 'Confirm delete' }));

    expect(screen.queryByText('Solo')).not.toBeInTheDocument();
    expect(screen.getByText(/No endpoints configured/)).toBeInTheDocument();
    expect(useSettingsStore.getState().endpoints).toHaveLength(0);
    expect(useSettingsStore.getState().activeEndpointId).toBeNull();
  });

  // ── API Key Eye Toggle ──────────────────────────────────

  it('toggles API key visibility with the eye button', async () => {
    renderWithProviders(<SettingsDialog open={true} />);

    await userEvent.click(screen.getByRole('button', { name: /Add Endpoint/ }));

    // API key field should start as password
    const apiKeyInput = screen.getByPlaceholderText('Enter API key');
    expect(apiKeyInput).toHaveAttribute('type', 'password');

    // Click the eye button to reveal
    await userEvent.click(screen.getByRole('button', { name: 'Show API key' }));
    expect(apiKeyInput).toHaveAttribute('type', 'text');

    // Click again to hide
    await userEvent.click(screen.getByRole('button', { name: 'Hide API key' }));
    expect(apiKeyInput).toHaveAttribute('type', 'password');
  });

  // ── Delete Confirmation Cancel ──────────────────────────

  it('cancelling delete confirmation keeps the endpoint', async () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'Keep Me',
      resourceUrl: 'https://keep.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    renderWithProviders(<SettingsDialog open={true} />);

    // Enter edit mode, click delete
    await userEvent.click(screen.getByRole('button', { name: 'Edit endpoint' }));
    await userEvent.click(screen.getByRole('button', { name: 'Delete endpoint' }));
    expect(screen.getByText(/Delete this endpoint/)).toBeInTheDocument();

    // Cancel the delete
    await userEvent.click(screen.getByRole('button', { name: 'Cancel delete' }));

    // Endpoint should still be there
    expect(useSettingsStore.getState().endpoints).toHaveLength(1);
  });

  // ── Edit Endpoint ──────────────────────────────────────

  it('edits an existing endpoint', async () => {
    useSettingsStore.getState().addEndpoint({
      displayName: 'Original',
      resourceUrl: 'https://original.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'old-key',
    });

    renderWithProviders(<SettingsDialog open={true} />);

    // Click edit button
    await userEvent.click(screen.getByRole('button', { name: 'Edit endpoint' }));

    // Form should be pre-filled
    expect(screen.getByText('Edit Connection')).toBeInTheDocument();
    const nameInput = screen.getByDisplayValue('Original');
    expect(nameInput).toBeInTheDocument();

    // Change the display name
    await userEvent.clear(nameInput);
    await userEvent.type(nameInput, 'Updated');

    // Save (auto-tests)
    await userEvent.click(screen.getByRole('button', { name: 'Save' }));

    // Wait for save
    await waitFor(() => {
      expect(useSettingsStore.getState().endpoints[0].displayName).toBe('Updated');
    });

    // Connection summary should show updated name
    expect(screen.getByText('Updated')).toBeInTheDocument();
    expect(screen.queryByText('Original')).not.toBeInTheDocument();
  });

  // ── Models visible by default ──────────────────────────

  it('shows ModelManager for the active endpoint without clicking Edit', () => {
    const epId = useSettingsStore.getState().addEndpoint({
      displayName: 'My EP',
      resourceUrl: 'https://my-ep.openai.azure.com',
      authMethod: 'apiKey',
      apiKey: 'key',
    });

    useSettingsStore
      .getState()
      .setModelsForEndpoint(epId, [
        { id: 'gpt-4.1', name: 'gpt-4.1', ownedBy: 'user', provider: 'OpenAI' },
      ]);

    renderWithProviders(<SettingsDialog open={true} />);

    // Models section should be visible immediately
    expect(screen.getByText('gpt-4.1')).toBeInTheDocument();
  });
});
