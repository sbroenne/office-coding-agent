import { createContext, useContext } from 'react';

export interface ChatActions {
  send: (text: string) => void | Promise<void>;
}

export const ChatActionsContext = createContext<ChatActions>({
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  send: () => {},
});

export function useChatActions(): ChatActions {
  return useContext(ChatActionsContext);
}
