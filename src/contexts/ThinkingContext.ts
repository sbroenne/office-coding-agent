import { createContext, useContext } from 'react';

export const ThinkingContext = createContext<string | null>(null);

export function useThinkingText(): string | null {
  return useContext(ThinkingContext);
}
