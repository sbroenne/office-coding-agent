import { createContext, useContext } from 'react';

export interface ChatActions {
  send: (text: string) => void | Promise<void>;
  enqueue: (text: string) => void;
}

export const ChatActionsContext = createContext<ChatActions>({
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  send: () => {},
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  enqueue: () => {},
});

export function useChatActions(): ChatActions {
  return useContext(ChatActionsContext);
}
