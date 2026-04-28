import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatComposer } from '@/components/chat/ChatComposer';

describe('ChatComposer slash commands', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('shows installed skills and prompt files as direct slash suggestions', async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        skills: [
          {
            type: 'skill',
            name: 'excel',
            description: 'Work with Excel workbooks',
            plugin: 'office-excel@office-coding-agent',
          },
        ],
        prompts: [
          {
            type: 'prompt',
            name: 'explain-code',
            description: 'Explain selected code',
            source: 'workspace',
          },
        ],
      }),
    });
    vi.stubGlobal('fetch', fetchMock);
    const onSend = vi.fn();

    renderWithProviders(<ChatComposer onSend={onSend} onCancel={vi.fn()} isRunning={false} />);

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/exc' },
    });

    const option = await screen.findByRole('option', { name: /\/excel/i });
    fireEvent.click(option);

    expect(onSend).not.toHaveBeenCalled();
    expect(screen.getByRole('textbox', { name: 'Message input' })).toHaveValue('/excel ');

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/explain' },
    });

    const promptOption = await screen.findByRole('option', { name: /\/explain-code/i });
    fireEvent.click(promptOption);

    expect(onSend).not.toHaveBeenCalled();
    expect(screen.getByRole('textbox', { name: 'Message input' })).toHaveValue('/explain-code ');
  });

  it('does not show skills management commands in slash suggestions', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ skills: [], prompts: [] }),
      })
    );

    renderWithProviders(<ChatComposer onSend={vi.fn()} onCancel={vi.fn()} isRunning={false} />);

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/skills' },
    });

    expect(await screen.findByRole('listbox', { name: 'slash suggestions' })).toBeVisible();
    expect(screen.queryByRole('option', { name: /\/skills list/i })).not.toBeInTheDocument();
  });
});
