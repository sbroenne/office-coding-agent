/**
 * Integration tests for AgentManagerPanel.
 *
 * Renders the real AgentManagerPanel with the real Zustand store.
 * Tests ZIP import, .md single-file import, and UI state flows.
 * Note: Import operations parse files locally but no longer persist to the store
 * (plugin state is managed by the Copilot CLI via ~/.copilot/config.json).
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import JSZip from 'jszip';
import { renderWithProviders } from '../test-utils';
import { AgentManagerPanel } from '@/components/AgentManagerDialog';
import { useSettingsStore } from '@/stores/settingsStore';

async function createAgentsZipFile(entries: Record<string, string>): Promise<File> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  return new File([buffer], 'agents.zip', { type: 'application/zip' });
}

beforeEach(() => {
  useSettingsStore.getState().reset();
});

describe('Integration: AgentManagerPanel', () => {
  it('renders the panel with import buttons', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText('Custom Agents')).toBeInTheDocument();
    expect(screen.getByLabelText('Import agents ZIP file')).toBeInTheDocument();
    expect(screen.getByLabelText('Import agent Markdown file')).toBeInTheDocument();
  });

  it('shows bundled agents in the read-only section', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText(/Bundled \(read-only\)/i)).toBeInTheDocument();
    // The Excel bundled agent must appear
    expect(screen.getAllByText('Excel').length).toBeGreaterThanOrEqual(1);
  });

  it('shows "No imported agents." when no custom agents have been imported', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.getByText('No imported agents.')).toBeInTheDocument();
  });

  it('shows error alert when ZIP contains no agents/ markdown files', async () => {
    renderWithProviders(<AgentManagerPanel />);

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
    renderWithProviders(<AgentManagerPanel />);

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

  it('shows error alert when .md file has no valid hosts', async () => {
    renderWithProviders(<AgentManagerPanel />);

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
    renderWithProviders(<AgentManagerPanel />);

    // Excel is a known bundled agent
    expect(
      screen.getByRole('button', { name: 'Download Excel as template' })
    ).toBeInTheDocument();
  });

  it('"Download all" button is not shown when no imported agents exist', () => {
    renderWithProviders(<AgentManagerPanel />);

    expect(screen.queryByText('Download all')).not.toBeInTheDocument();
  });
});
