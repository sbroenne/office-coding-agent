import { describe, it, expect } from 'vitest';
import { messagesToCoreMessages } from '@/services/ai/chatService';
import type { ChatMessage } from '@/types';

describe('messagesToCoreMessages', () => {
  it('converts a simple user message', () => {
    const messages: ChatMessage[] = [{ role: 'user', content: 'Hello' }];
    const result = messagesToCoreMessages(messages);

    expect(result).toEqual([{ role: 'user', content: 'Hello' }]);
  });

  it('converts a simple assistant message', () => {
    const messages: ChatMessage[] = [{ role: 'assistant', content: 'Hi there!' }];
    const result = messagesToCoreMessages(messages);

    expect(result).toEqual([{ role: 'assistant', content: 'Hi there!' }]);
  });

  it('skips streaming messages', () => {
    const messages: ChatMessage[] = [
      { role: 'user', content: 'Hello' },
      { role: 'assistant', content: 'partial...', isStreaming: true },
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({ role: 'user', content: 'Hello' });
  });

  it('converts assistant message with tool calls', () => {
    const messages: ChatMessage[] = [
      {
        role: 'assistant',
        content: '',
        toolCalls: [
          {
            id: 'call-1',
            functionName: 'get_range_values',
            arguments: '{"address":"A1:C5"}',
            parsedArguments: { address: 'A1:C5' },
          },
        ],
      },
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('assistant');

    const content = result[0].content as { type: string }[];
    expect(content).toHaveLength(1);
    expect(content[0]).toMatchObject({
      type: 'tool-call',
      toolCallId: 'call-1',
      toolName: 'get_range_values',
      input: { address: 'A1:C5' },
    });
  });

  it('includes text content alongside tool calls', () => {
    const messages: ChatMessage[] = [
      {
        role: 'assistant',
        content: 'Let me check that.',
        toolCalls: [
          {
            id: 'call-1',
            functionName: 'get_used_range',
            arguments: '{}',
            parsedArguments: {},
          },
        ],
      },
    ];
    const result = messagesToCoreMessages(messages);

    const content = result[0].content as { type: string; text?: string }[];
    expect(content).toHaveLength(2);
    expect(content[0]).toEqual({ type: 'text', text: 'Let me check that.' });
    expect(content[1]).toMatchObject({ type: 'tool-call', toolCallId: 'call-1' });
  });

  it('falls back to JSON.parse when parsedArguments is missing', () => {
    const messages: ChatMessage[] = [
      {
        role: 'assistant',
        content: '',
        toolCalls: [
          {
            id: 'call-1',
            functionName: 'set_range_values',
            arguments: '{"address":"B2","values":[[42]]}',
          },
        ],
      },
    ];
    const result = messagesToCoreMessages(messages);

    const content = result[0].content as { type: string; input?: unknown }[];
    expect(content[0]).toMatchObject({
      type: 'tool-call',
      input: { address: 'B2', values: [[42]] },
    });
  });

  it('converts a tool result message', () => {
    const messages: ChatMessage[] = [
      {
        role: 'tool',
        content: '{"success":true,"data":{"address":"A1:C5"}}',
        toolCallId: 'call-1',
      },
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('tool');

    const content = result[0].content as { type: string; toolCallId: string }[];
    expect(content).toHaveLength(1);
    expect(content[0]).toMatchObject({
      type: 'tool-result',
      toolCallId: 'call-1',
    });
  });

  it('merges consecutive tool result messages', () => {
    const messages: ChatMessage[] = [
      { role: 'tool', content: 'result-1', toolCallId: 'call-1' },
      { role: 'tool', content: 'result-2', toolCallId: 'call-2' },
    ];
    const result = messagesToCoreMessages(messages);

    // Both should be merged into a single tool message
    expect(result).toHaveLength(1);
    expect(result[0].role).toBe('tool');

    const content = result[0].content as { type: string; toolCallId: string }[];
    expect(content).toHaveLength(2);
    expect(content[0].toolCallId).toBe('call-1');
    expect(content[1].toolCallId).toBe('call-2');
  });

  it('does not merge tool results separated by other messages', () => {
    const messages: ChatMessage[] = [
      { role: 'tool', content: 'result-1', toolCallId: 'call-1' },
      { role: 'assistant', content: 'Thinking...' },
      { role: 'tool', content: 'result-2', toolCallId: 'call-2' },
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(3);
    expect(result[0].role).toBe('tool');
    expect(result[1].role).toBe('assistant');
    expect(result[2].role).toBe('tool');
  });

  it('ignores tool messages without toolCallId', () => {
    const messages: ChatMessage[] = [
      { role: 'tool', content: 'orphan result' }, // no toolCallId
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(0);
  });

  it('converts a full multi-turn conversation', () => {
    const messages: ChatMessage[] = [
      { role: 'user', content: 'What data is in the spreadsheet?' },
      {
        role: 'assistant',
        content: '',
        toolCalls: [
          {
            id: 'call-1',
            functionName: 'get_used_range',
            arguments: '{}',
            parsedArguments: {},
          },
        ],
      },
      {
        role: 'tool',
        content: '{"address":"A1:C5","rowCount":5}',
        toolCallId: 'call-1',
      },
      { role: 'assistant', content: 'Your spreadsheet has data in A1:C5.' },
    ];
    const result = messagesToCoreMessages(messages);

    expect(result).toHaveLength(4);
    expect(result[0]).toEqual({ role: 'user', content: 'What data is in the spreadsheet?' });
    expect(result[1].role).toBe('assistant');
    expect(result[2].role).toBe('tool');
    expect(result[3]).toEqual({
      role: 'assistant',
      content: 'Your spreadsheet has data in A1:C5.',
    });
  });

  it('handles empty message array', () => {
    const result = messagesToCoreMessages([]);
    expect(result).toEqual([]);
  });

  it('handles assistant with empty toolCalls array', () => {
    const messages: ChatMessage[] = [{ role: 'assistant', content: 'Just text.', toolCalls: [] }];
    const result = messagesToCoreMessages(messages);

    // Empty toolCalls → treated as plain text message
    expect(result).toEqual([{ role: 'assistant', content: 'Just text.' }]);
  });
});
