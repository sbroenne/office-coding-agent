import { beforeEach, describe, expect, it } from 'vitest';
import { useChatStore } from '@/stores/chatStore';

beforeEach(() => {
  localStorage.removeItem('office-coding-agent-chat');
  useChatStore.getState().reset();
});

describe('chatStore', () => {
  it('starts with no messages', () => {
    expect(useChatStore.getState().messages).toEqual([]);
  });

  it('stores serializable message payloads', () => {
    const messages = [{ role: 'user', content: 'hello', fn: () => 'x' }];

    useChatStore.getState().setMessages(messages);

    expect(useChatStore.getState().messages).toEqual([{ role: 'user', content: 'hello' }]);
  });

  it('clears and resets messages', () => {
    useChatStore.getState().setMessages([{ role: 'assistant', content: 'hi' }]);
    useChatStore.getState().clearMessages();
    expect(useChatStore.getState().messages).toEqual([]);

    useChatStore.getState().setMessages([{ role: 'assistant', content: 'again' }]);
    useChatStore.getState().reset();
    expect(useChatStore.getState().messages).toEqual([]);
  });
});
