/**
 * Integration tests for SkillManagerPanel.
 *
 * Renders the real SkillManagerPanel. Bundled skills are shown read-only with
 * download-as-template buttons. Custom skills are installed via plugins.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { SkillManagerPanel } from '@/components/SkillManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillManagerPanel', () => {
  it('shows bundled skills in the read-only section', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    expect(screen.getByText('excel')).toBeInTheDocument();
  });

  it('shows plugin hub message in the "From plugins" section', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText('From plugins')).toBeInTheDocument();
    expect(screen.getByText(/installed via the Copilot CLI Plugin Hub/i)).toBeInTheDocument();
  });

  it('bundled skills each have a "Download as template" button', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(
      screen.getByRole('button', { name: 'Download excel as template' })
    ).toBeInTheDocument();
  });

  it('does not show import upload buttons', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.queryByLabelText('Import skills ZIP file')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Import skill Markdown file')).not.toBeInTheDocument();
  });
});
