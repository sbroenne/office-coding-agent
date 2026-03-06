/**
 * Maps a tool ID to a codicon name for the chain-of-thought line,
 * matching VS Code's `getToolInvocationIcon()` in chatThinkingContentPart.ts.
 */
export function getToolIcon(toolId: string): string {
  const lower = toolId.toLowerCase();

  if (lower === 'task_complete') return 'pass';

  if (
    lower.includes('search') ||
    lower.includes('grep') ||
    lower.includes('find') ||
    lower.includes('list') ||
    lower.includes('semantic') ||
    lower.includes('changes') ||
    lower.includes('codebase')
  ) {
    return 'search';
  }

  if (lower.includes('read') || lower.includes('get_file') || lower.includes('problems')) {
    return 'book';
  }

  if (
    lower.includes('edit') ||
    lower.includes('create') ||
    lower.includes('replace') ||
    lower.includes('update') ||
    lower.includes('set') ||
    lower.includes('insert') ||
    lower.includes('delete') ||
    lower.includes('add') ||
    lower.includes('remove')
  ) {
    return 'pencil';
  }

  if (lower.includes('terminal') || lower.includes('run')) {
    return 'terminal';
  }

  return 'tools';
}
