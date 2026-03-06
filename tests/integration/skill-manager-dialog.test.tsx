/**
 * Integration tests for SkillManagerPanel.
 *
 * Renders the real SkillManagerPanel. Bundled skills are shown read-only with
 * download-as-template buttons. Users can upload custom skills via ZIP or .md.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
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

  it('shows the Uploaded section with count', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText(/Uploaded \(0\)/i)).toBeInTheDocument();
  });

  it('bundled skills each have a "Download as template" button', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(
      screen.getByRole('button', { name: 'Download excel as template' })
    ).toBeInTheDocument();
  });

  it('shows upload buttons for ZIP and .md', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByLabelText('Import skills ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import skill Markdown file')).toBeInTheDocument();
  });

  it('shows empty-state message when no skills are uploaded', () => {
    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText(/No uploaded skills/i)).toBeInTheDocument();
  });

  it('shows uploaded skill with delete button when store has an imported skill', () => {
    const skill = {
      metadata: {
        name: 'my-skill',
        description: 'A test skill',
        version: '1.0.0',
        tags: [],
        hosts: [],
      },
      content: 'Do something useful.',
    };
    useSettingsStore.getState().addImportedSkill(skill);

    renderWithProviders(<SkillManagerPanel />);

    expect(screen.getByText('my-skill')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Remove my-skill' })).toBeInTheDocument();
    expect(screen.getByText(/Uploaded \(1\)/i)).toBeInTheDocument();
  });

  it('clicking Remove deletes the skill from the store', async () => {
    const skill = {
      metadata: {
        name: 'removable-skill',
        description: '',
        version: '1.0.0',
        tags: [],
        hosts: [],
      },
      content: 'content',
    };
    useSettingsStore.getState().addImportedSkill(skill);

    renderWithProviders(<SkillManagerPanel />);

    await userEvent.click(screen.getByRole('button', { name: 'Remove removable-skill' }));

    expect(useSettingsStore.getState().importedSkills).toHaveLength(0);
    expect(screen.queryByText('removable-skill')).not.toBeInTheDocument();
  });
});
