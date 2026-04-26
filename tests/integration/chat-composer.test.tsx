import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, screen, waitFor } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatComposer } from '@/components/chat/ChatComposer';

describe('ChatComposer slash commands', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('shows installed CLI skills as direct slash suggestions and inserts the selected skill', async () => {
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
  });

  it('shows Copilot CLI skills management command suggestions for /skills', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ skills: [] }),
      })
    );

    renderWithProviders(<ChatComposer onSend={vi.fn()} onCancel={vi.fn()} isRunning={false} />);

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/skills' },
    });

    await waitFor(() => {
      expect(screen.getByRole('listbox', { name: 'skills command suggestions' })).toBeVisible();
    });
    expect(screen.getByRole('option', { name: /\/skills list/i })).toBeVisible();
  });
});
