/**
 * Integration tests for SkillManagerDialog.
 *
 * Renders the real SkillManagerDialog with the real Zustand store.
 * Tests ZIP import, .md single-file import, remove, and UI state flows.
 */
import React from 'react';
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import JSZip from 'jszip';
import { renderWithProviders } from '../test-utils';
import { SkillManagerDialog } from '@/components/SkillManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';
import type { AgentSkill } from '@/types';

async function createSkillsZipFile(entries: Record<string, string>): Promise<File> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  return new File([buffer], 'skills.zip', { type: 'application/zip' });
}

const validSkillMarkdown = `---
name: My Custom Skill
description: A custom skill for testing
version: 1.0.0
---
Custom skill instructions.`;

function makeSkill(name: string): AgentSkill {
  return {
    metadata: { name, description: `${name} description`, version: '1.0.0', tags: [], hosts: [] },
    content: `${name} content.`,
  };
}

const OpenDialog: React.FC = () => {
  const [open, setOpen] = React.useState(true);
  return <SkillManagerDialog open={open} onOpenChange={setOpen} />;
};

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: SkillManagerDialog', () => {
  it('renders the dialog with title and both import buttons', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByRole('dialog', { name: 'Manage Skills' })).toBeInTheDocument();
    expect(screen.getByLabelText('Import skills ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import skill Markdown file')).toBeInTheDocument();
  });

  it('shows bundled skills in the read-only section', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    // The bundled 'excel' skill must appear
    expect(screen.getByText('excel')).toBeInTheDocument();
  });

  it('shows "No imported skills." when nothing has been imported', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('No imported skills.')).toBeInTheDocument();
  });

  it('imports skills from a valid ZIP file and shows success status', async () => {
    renderWithProviders(<OpenDialog />);

    const zipFile = await createSkillsZipFile({
      'skills/my-skill.md': validSkillMarkdown,
    });

    await userEvent.upload(screen.getByLabelText('Import skills ZIP file'), zipFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Skill')).toBeInTheDocument();
    });

    expect(screen.getByRole('status')).toHaveTextContent('Imported 1 skill from skills.zip.');
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
  });

  it('imports multiple skills from a ZIP with multiple files', async () => {
    renderWithProviders(<OpenDialog />);

    const skillB = `---
name: Skill B
description: Second skill
version: 1.0.0
---
Skill B content.`;

    const zipFile = await createSkillsZipFile({
      'skills/skill-a.md': validSkillMarkdown,
      'skills/skill-b.md': skillB,
    });

    await userEvent.upload(screen.getByLabelText('Import skills ZIP file'), zipFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Skill')).toBeInTheDocument();
    });

    expect(screen.getByText('Skill B')).toBeInTheDocument();
    expect(screen.getByRole('status')).toHaveTextContent('Imported 2 skills from skills.zip.');
  });

  it('shows error alert when ZIP contains no skills/ markdown files', async () => {
    renderWithProviders(<OpenDialog />);

    const zip = new JSZip();
    zip.file('notes/readme.txt', 'not a skill');
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    const badFile = new File([buffer], 'bad.zip', { type: 'application/zip' });

    await userEvent.upload(screen.getByLabelText('Import skills ZIP file'), badFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('imports a single skill from a .md file and shows success status', async () => {
    renderWithProviders(<OpenDialog />);

    const mdFile = new File([validSkillMarkdown], 'skill.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import skill Markdown file'), mdFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Skill')).toBeInTheDocument();
    });

    expect(screen.getByRole('status')).toHaveTextContent(
      'Imported skill "My Custom Skill" from skill.md.'
    );
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
  });

  it('shows error alert when .md file has no name', async () => {
    renderWithProviders(<OpenDialog />);

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
    renderWithProviders(<OpenDialog />);

    // 'excel' is a known bundled skill
    expect(
      screen.getByRole('button', { name: 'Download excel as template' })
    ).toBeInTheDocument();
  });

  it('"Download all" button is not shown when no imported skills exist', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.queryByText('Download all')).not.toBeInTheDocument();
  });

  it('"Download all" button appears once imported skills exist', () => {
    useSettingsStore.getState().importSkills([makeSkill('My Skill')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('Download all')).toBeInTheDocument();
  });

  it('imported skill shows individual Download and Remove buttons', () => {
    useSettingsStore.getState().importSkills([makeSkill('My Import')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByRole('button', { name: 'Download My Import' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Remove My Import' })).toBeInTheDocument();
  });

  it('Remove button removes the skill from the list and store', async () => {
    useSettingsStore.getState().importSkills([makeSkill('To Remove')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('To Remove')).toBeInTheDocument();

    await userEvent.click(screen.getByRole('button', { name: 'Remove To Remove' }));

    await waitFor(() => {
      expect(screen.queryByText('To Remove')).not.toBeInTheDocument();
    });

    expect(useSettingsStore.getState().importedSkills).toHaveLength(0);
    expect(screen.getByText('No imported skills.')).toBeInTheDocument();
  });

  it('removing one of many imported skills only removes the target', async () => {
    useSettingsStore.getState().importSkills([makeSkill('Keep Me'), makeSkill('Remove Me')]);

    renderWithProviders(<OpenDialog />);

    await userEvent.click(screen.getByRole('button', { name: 'Remove Remove Me' }));

    await waitFor(() => {
      expect(screen.queryByText('Remove Me')).not.toBeInTheDocument();
    });

    expect(screen.getByText('Keep Me')).toBeInTheDocument();
    expect(useSettingsStore.getState().importedSkills).toHaveLength(1);
  });

  it('clears error when a new import is attempted', async () => {
    renderWithProviders(<OpenDialog />);

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
