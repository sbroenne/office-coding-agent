import BASE_PROMPT from './BASE_PROMPT.md';
import type { OfficeHostApp } from '@/services/office/host';

export { BASE_PROMPT };

interface SessionSystemPromptOptions {
  memoryContext?: string;
}

export function getAppPromptForHost(host: OfficeHostApp): string {
  void host;
  return '';
}

export function buildSystemPrompt(host: OfficeHostApp): string {
  const hostPrompt = getAppPromptForHost(host).trim();
  return hostPrompt ? `${BASE_PROMPT}\n\n${hostPrompt}` : BASE_PROMPT;
}

export function buildSessionSystemPrompt(
  host: OfficeHostApp,
  options: SessionSystemPromptOptions = {}
): string {
  const sections = [buildSystemPrompt(host)];
  const memoryContext = options.memoryContext?.trim();

  if (memoryContext) {
    sections.push(memoryContext);
  }

  return sections.join('\n\n');
}
