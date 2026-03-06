/**
 * Integration tests for SkillManagerPanel.
 *
 * Renders the real SkillManagerPanel with the real Zustand store.
 * Tests ZIP import, .md single-file import, and UI state flows.
 * Note: Import operations parse files locally but no longer persist to the store
 * (plugin state is managed by the Copilot CLI via ~/.copilot/config.json).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import JSZip from 'jszip';
import { renderWithProviders } from '../test-utils';
import { SkillManagerPanel } from '@/components/SkillManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

const validSkillMarkdown = `---
name: My Custom Skill
description: A custom skill for testing
version: 1.0.0
---
Custom skill instructions.`;

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillManagerPanel', () => {
  it('renders the panel with import buttons', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText('Custom Skills')).toBeInTheDocument();
    expect(screen.getByLabelText('Import skills ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import skill Markdown file')).toBeInTheDocument();
  });

  it('shows bundled skills in the read-only section', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    // The bundled 'excel' skill must appear
    expect(screen.getByText('excel')).toBeInTheDocument();
  });

  it('shows "No imported skills." when nothing has been imported', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText('No imported skills.')).toBeInTheDocument();
  });

  it('shows error alert when ZIP contains no skills/ markdown files', async () => {
    renderWithProviders(<SkillManagerPanel />);

    const zip = new JSZip();
    zip.file('notes/readme.txt', 'not a skill');
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    const badFile = new File([buffer], 'bad.zip', { type: 'application/zip' });

    await userEvent.upload(screen.getByLabelText('Import skills ZIP file'), badFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('shows error alert when .md file has no name', async () => {
    renderWithProviders(<SkillManagerPanel />);

    const badMd = `---
description: no name here
version: 1.0.0
---
Content`;

    const mdFile = new File([badMd], 'no-name.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import skill Markdown file'), mdFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('bundled skills each have a "Download as template" button', () => {
    renderWithProviders(<SkillManagerPanel />);

    // 'excel' is a known bundled skill
    expect(
      screen.getByRole('button', { name: 'Download excel as template' })
    ).toBeInTheDocument();
  });

  it('"Download all" button is not shown when no imported skills exist', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.queryByText('Download all')).not.toBeInTheDocument();
  });

  it('clears error when a new import is attempted', async () => {
    renderWithProviders(<SkillManagerPanel />);

    // First: trigger an error
    const badMd = `---
description: no name
version: 1.0.0
---
Content`;
    const badFile = new File([badMd], 'bad.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import skill Markdown file'), badFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });

    // Then: import a valid .md — error should be cleared
    const goodFile = new File([validSkillMarkdown], 'good.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import skill Markdown file'), goodFile);

    await waitFor(() => {
      expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    });
    expect(screen.getByRole('status')).toBeInTheDocument();
  });
});
