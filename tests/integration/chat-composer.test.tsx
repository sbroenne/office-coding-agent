import { beforeEach, describe, expect, it, vi } from 'vitest';
import { fireEvent, screen, waitFor } from '@testing-library/react';
import { renderWithProviders } from '../test-utils';
import { ChatComposer } from '@/components/chat/ChatComposer';

describe('ChatComposer slash commands', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('shows installed CLI skills and sends the selected skill invocation', async () => {
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
        prompts: [],
      }),
    });
    vi.stubGlobal('fetch', fetchMock);
    const onSend = vi.fn();

    renderWithProviders(<ChatComposer onSend={onSend} onCancel={vi.fn()} isRunning={false} />);

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/skills' },
    });

    const option = await screen.findByRole('option', { name: /excel/i });
    fireEvent.click(option);

    expect(onSend).toHaveBeenCalledWith('Use the excel skill.');
  });

  it('shows an empty prompt state when no CLI prompts are installed', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ skills: [], prompts: [] }),
      })
    );

    renderWithProviders(<ChatComposer onSend={vi.fn()} onCancel={vi.fn()} isRunning={false} />);

    fireEvent.change(screen.getByRole('textbox', { name: 'Message input' }), {
      target: { value: '/prompts' },
    });

    await waitFor(() => {
      expect(screen.getByText('No prompts found from installed Copilot CLI plugins.')).toBeVisible();
    });
  });
});
