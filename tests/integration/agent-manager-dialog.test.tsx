/**
 * Integration tests for AgentManagerDialog.
 *
 * Renders the real AgentManagerDialog with the real Zustand store.
 * Tests ZIP import, .md single-file import, remove, and UI state flows.
 */
import React from 'react';
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import JSZip from 'jszip';
import { renderWithProviders } from '../test-utils';
import { AgentManagerDialog } from '@/components/AgentManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';
import type { AgentConfig } from '@/types';

async function createAgentsZipFile(entries: Record<string, string>): Promise<File> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  return new File([buffer], 'agents.zip', { type: 'application/zip' });
}

const validAgentMarkdown = `---
name: My Custom Agent
description: A custom agent for testing
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Custom agent instructions.`;

function makeAgent(name: string): AgentConfig {
  return {
    metadata: {
      name,
      description: `${name} description`,
      version: '1.0.0',
      hosts: ['excel'],
      defaultForHosts: [],
    },
    instructions: `${name} instructions.`,
  };
}

const OpenDialog: React.FC = () => {
  const [open, setOpen] = React.useState(true);
  return <AgentManagerDialog open={open} onOpenChange={setOpen} />;
};

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: AgentManagerDialog', () => {
  it('renders the dialog with title and both import buttons', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByRole('dialog', { name: 'Manage Agents' })).toBeInTheDocument();
    expect(screen.getByLabelText('Import agents ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import agent Markdown file')).toBeInTheDocument();
  });

  it('shows bundled agents in the read-only section', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    // The Excel bundled agent must appear
    expect(screen.getAllByText('Excel').length).toBeGreaterThanOrEqual(1);
  });

  it('shows "No imported agents." when no custom agents have been imported', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('No imported agents.')).toBeInTheDocument();
  });

  it('imports agents from a valid ZIP file and shows success status', async () => {
    renderWithProviders(<OpenDialog />);

    const zipFile = await createAgentsZipFile({
      'agents/my-agent.md': validAgentMarkdown,
    });

    await userEvent.upload(screen.getByLabelText('Import agents ZIP file'), zipFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Agent')).toBeInTheDocument();
    });

    expect(screen.getByRole('status')).toHaveTextContent('Imported 1 agent from agents.zip.');
    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
  });

  it('imports multiple agents from a ZIP with multiple files', async () => {
    renderWithProviders(<OpenDialog />);

    const agentB = `---
name: Agent B
description: Second agent
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Agent B instructions.`;

    const zipFile = await createAgentsZipFile({
      'agents/agent-a.md': validAgentMarkdown,
      'agents/agent-b.md': agentB,
    });

    await userEvent.upload(screen.getByLabelText('Import agents ZIP file'), zipFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Agent')).toBeInTheDocument();
    });

    expect(screen.getByText('Agent B')).toBeInTheDocument();
    expect(screen.getByRole('status')).toHaveTextContent('Imported 2 agents from agents.zip.');
  });

  it('shows error alert when ZIP contains no agents/ markdown files', async () => {
    renderWithProviders(<OpenDialog />);

    const zip = new JSZip();
    zip.file('notes/readme.txt', 'not an agent');
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    const badFile = new File([buffer], 'bad.zip', { type: 'application/zip' });

    await userEvent.upload(screen.getByLabelText('Import agents ZIP file'), badFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('shows error alert when ZIP agent has no valid hosts', async () => {
    renderWithProviders(<OpenDialog />);

    const noHostsMd = `---
name: No Hosts Agent
description: desc
version: 1.0.0
---
Instructions`;

    const zipFile = await createAgentsZipFile({
      'agents/no-hosts.md': noHostsMd,
    });

    await userEvent.upload(screen.getByLabelText('Import agents ZIP file'), zipFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('imports a single agent from a .md file and shows success status', async () => {
    renderWithProviders(<OpenDialog />);

    const mdFile = new File([validAgentMarkdown], 'agent.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import agent Markdown file'), mdFile);

    await waitFor(() => {
      expect(screen.getByText('My Custom Agent')).toBeInTheDocument();
    });

    expect(screen.getByRole('status')).toHaveTextContent(
      'Imported agent "My Custom Agent" from agent.md.'
    );
    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
  });

  it('shows error alert when .md file has no valid hosts', async () => {
    renderWithProviders(<OpenDialog />);

    const badMd = `---
name: No Hosts Agent
description: desc
version: 1.0.0
---
Instructions`;

    const mdFile = new File([badMd], 'no-hosts.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import agent Markdown file'), mdFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });
  });

  it('bundled agents each have a "Download as template" button', () => {
    renderWithProviders(<OpenDialog />);

    // Excel is a known bundled agent
    expect(
      screen.getByRole('button', { name: 'Download Excel as template' })
    ).toBeInTheDocument();
  });

  it('"Download all" button is not shown when no imported agents exist', () => {
    renderWithProviders(<OpenDialog />);

    expect(screen.queryByText('Download all')).not.toBeInTheDocument();
  });

  it('"Download all" button appears once imported agents exist', () => {
    useSettingsStore.getState().importAgents([makeAgent('Custom Agent')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('Download all')).toBeInTheDocument();
  });

  it('imported agent shows individual Download and Remove buttons', () => {
    useSettingsStore.getState().importAgents([makeAgent('My Import')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByRole('button', { name: 'Download My Import' })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Remove My Import' })).toBeInTheDocument();
  });

  it('Remove button removes the agent from the list and store', async () => {
    useSettingsStore.getState().importAgents([makeAgent('To Remove')]);

    renderWithProviders(<OpenDialog />);

    expect(screen.getByText('To Remove')).toBeInTheDocument();

    await userEvent.click(screen.getByRole('button', { name: 'Remove To Remove' }));

    await waitFor(() => {
      expect(screen.queryByText('To Remove')).not.toBeInTheDocument();
    });

    expect(useSettingsStore.getState().importedAgents).toHaveLength(0);
    expect(screen.getByText('No imported agents.')).toBeInTheDocument();
  });

  it('removing one of many imported agents only removes the target', async () => {
    useSettingsStore.getState().importAgents([makeAgent('Keep Me'), makeAgent('Remove Me')]);

    renderWithProviders(<OpenDialog />);

    await userEvent.click(screen.getByRole('button', { name: 'Remove Remove Me' }));

    await waitFor(() => {
      expect(screen.queryByText('Remove Me')).not.toBeInTheDocument();
    });

    expect(screen.getByText('Keep Me')).toBeInTheDocument();
    expect(useSettingsStore.getState().importedAgents).toHaveLength(1);
  });

  it('clears error when a new import is attempted', async () => {
    renderWithProviders(<OpenDialog />);

    // First: trigger an error
    const badMd = `---
name: No Hosts
description: desc
version: 1.0.0
---
Instructions`;
    const badFile = new File([badMd], 'bad.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import agent Markdown file'), badFile);

    await waitFor(() => {
      expect(screen.getByRole('alert')).toBeInTheDocument();
    });

    // Then: import a valid .md — error should be cleared
    const goodFile = new File([validAgentMarkdown], 'good.md', { type: 'text/markdown' });
    await userEvent.upload(screen.getByLabelText('Import agent Markdown file'), goodFile);

    await waitFor(() => {
      expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    });
    expect(screen.getByRole('status')).toBeInTheDocument();
  });
});
