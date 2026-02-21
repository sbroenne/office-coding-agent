import BASE_PROMPT from './BASE_PROMPT.md';
import EXCEL_APP_PROMPT from './prompts/EXCEL_APP_PROMPT.md';
import POWERPOINT_APP_PROMPT from './prompts/POWERPOINT_APP_PROMPT.md';
import WORD_APP_PROMPT from './prompts/WORD_APP_PROMPT.md';
import type { OfficeHostApp } from '@/services/office/host';

export { BASE_PROMPT };

export function getAppPromptForHost(host: OfficeHostApp): string {
  switch (host) {
    case 'excel':
      return EXCEL_APP_PROMPT;
    case 'powerpoint':
      return POWERPOINT_APP_PROMPT;
    case 'word':
      return WORD_APP_PROMPT;
    default:
      return 'You are an AI assistant running inside a Microsoft Office add-in.';
  }
}

export function buildSystemPrompt(host: OfficeHostApp): string {
  return `${BASE_PROMPT}\n\n${getAppPromptForHost(host)}`;
}
