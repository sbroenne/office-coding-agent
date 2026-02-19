/**
 * Integration test for ModelManager component.
 * ModelManager is now a stub (renders null) — verify it renders without errors.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { render } from '@testing-library/react';
import { ModelManager } from '@/components/ModelManager';
import { useSettingsStore } from '@/stores/settingsStore';

describe('ModelManager', () => {
  beforeEach(() => {
    useSettingsStore.getState().reset();
  });

  it('renders without errors given an endpointId', () => {
    const { container } = render(<ModelManager endpointId="some-id" />);
    // Stub renders null — container should be empty
    expect(container.firstChild).toBeNull();
  });

  it('renders without errors with an empty endpointId', () => {
    const { container } = render(<ModelManager endpointId="" />);
    expect(container.firstChild).toBeNull();
  });
});
