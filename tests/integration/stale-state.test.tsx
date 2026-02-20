/**
 * Tests for stale localStorage / state hydration scenarios.
 *
 * The settings store persists to localStorage via Zustand's `persist` middleware.
 * When a user returns to the add-in, they may have stale data from a previous
 * session (deleted endpoints, old model IDs, outdated defaultModelId).
 *
 * These tests verify the store and UI handle those cases gracefully,
 * not with silent failures or blank "Select model" dropdowns.
 */
import { describe, it, expect, beforeEach, vi } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { useSettingsStore } from '@/stores/settingsStore';
import { ModelPicker } from '@/components/ModelPicker';

// ─── Mock services ───
vi.mock('@/services/ai', () => ({
  discoverModels: vi.fn(),
  validateModelDeployment: vi.fn(),
}));

describe('Stale state scenarios', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  // ─── Store-level hydration ───

  describe('store getters with stale references', () => {
    it('getActiveModel returns undefined when activeModelId references a non-existent model', () => {
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });

      // Set a model, then manually set activeModelId to something that doesn't exist
      useSettingsStore
        .getState()
        .setModelsForEndpoint(epId, [
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
        ]);
      useSettingsStore.getState().setActiveModel('deleted-model-from-old-session');

      // getActiveModel should return undefined — it should NOT crash
      expect(useSettingsStore.getState().getActiveModel()).toBeUndefined();
    });

    it('getActiveEndpoint returns undefined when activeEndpointId references a deleted endpoint', () => {
      // Add and then remove an endpoint
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Will Be Deleted',
        resourceUrl: 'https://gone.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });
      useSettingsStore.getState().removeEndpoint(epId);

      // Manually set a stale activeEndpointId (simulating localStorage hydration)
      useSettingsStore.setState({ activeEndpointId: 'stale-endpoint-id-from-old-session' });

      expect(useSettingsStore.getState().getActiveEndpoint()).toBeUndefined();
    });

    it('getModelsForActiveEndpoint returns empty array when active endpoint has no models', () => {
      useSettingsStore.getState().addEndpoint({
        displayName: 'Empty',
        resourceUrl: 'https://empty.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });

      // Endpoint exists but has no models
      expect(useSettingsStore.getState().getModelsForActiveEndpoint()).toEqual([]);
    });
  });

  // ─── UI-level stale state ───

  describe('ModelPicker with stale store state', () => {
    it('shows "Select model" when activeModelId is stale (does not match any model)', () => {
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });

      useSettingsStore
        .getState()
        .setModelsForEndpoint(epId, [
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
        ]);

      // Simulate stale activeModelId from a previous session
      useSettingsStore.getState().setActiveModel('claude-opus-4-6-from-last-week');

      renderWithProviders(<ModelPicker />);

      // Should show "Select model" because the active model doesn't exist
      expect(screen.getByText('Select model')).toBeInTheDocument();
    });

    it('shows model name when store is correctly configured (wizard sequence)', () => {
      // Replicate the EXACT sequence the wizard performs:
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });

      useSettingsStore
        .getState()
        .setModelsForEndpoint(epId, [
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
        ]);

      useSettingsStore.getState().setActiveEndpoint(epId);

      renderWithProviders(<ModelPicker />);

      // Should show actual model name, NOT "Select model"
      expect(screen.queryByText('Select model')).not.toBeInTheDocument();
      expect(screen.getByText('gpt-5.2-chat')).toBeInTheDocument();
    });

    it('shows "No endpoint configured" when activeEndpointId is stale', () => {
      // Set a stale activeEndpointId with no matching endpoint
      useSettingsStore.setState({ activeEndpointId: 'nonexistent-ep' });

      renderWithProviders(<ModelPicker />);

      expect(screen.getByText('No endpoint configured')).toBeInTheDocument();
    });
  });

  // ─── Stale defaultModelId across sessions ───

  describe('defaultModelId persistence', () => {
    it('defaultModelId is gpt-5.2-chat after reset', () => {
      useSettingsStore.getState().reset();
      expect(useSettingsStore.getState().defaultModelId).toBe('gpt-5.2-chat');
    });

    it('setActiveEndpoint picks first model when defaultModelId does not match any model', () => {
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });

      // Models that DON'T include the default
      useSettingsStore.getState().setModelsForEndpoint(epId, [
        { id: 'kimi-k2', name: 'kimi-k2', ownedBy: 'user', provider: 'Other' },
        { id: 'claude-opus-4-6', name: 'claude-opus-4-6', ownedBy: 'user', provider: 'Anthropic' },
      ]);

      useSettingsStore.getState().setActiveModel('');
      useSettingsStore.getState().setActiveEndpoint(epId);

      // Should fall back to first model, not leave null
      expect(useSettingsStore.getState().activeModelId).toBe('kimi-k2');
    });

    it('reset clears all state back to defaults', () => {
      // Set up some state
      const epId = useSettingsStore.getState().addEndpoint({
        displayName: 'Test',
        resourceUrl: 'https://test.openai.azure.com',
        authMethod: 'apiKey',
        apiKey: 'key',
      });
      useSettingsStore
        .getState()
        .setModelsForEndpoint(epId, [
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'user', provider: 'OpenAI' },
        ]);

      // Reset
      useSettingsStore.getState().reset();

      const state = useSettingsStore.getState();
      expect(state.endpoints).toEqual([]);
      expect(state.activeEndpointId).toBeNull();
      expect(state.activeModelId).toBeNull();
      expect(state.endpointModels).toEqual({});
      expect(state.defaultModelId).toBe('gpt-5.2-chat');
    });
  });
});
