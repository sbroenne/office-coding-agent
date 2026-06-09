import type { AgentSkill, SkillMetadata } from '@/types/skill';

/**
 * Parse YAML frontmatter from a skill markdown file.
 * Handles the `---` delimited block at the top of the file.
 */
export function parseFrontmatter(raw: string): { metadata: SkillMetadata; content: string } {
  const trimmed = raw.trimStart();

  if (!trimmed.startsWith('---')) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [], hosts: [] },
      content: trimmed,
    };
  }

  const endIndex = trimmed.indexOf('---', 3);
  if (endIndex === -1) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [], hosts: [] },
      content: trimmed,
    };
  }

  const yamlBlock = trimmed.slice(3, endIndex).trim();
  const content = trimmed.slice(endIndex + 3).trim();

  // Simple YAML parser for flat key-value and tag arrays
  const metadata: SkillMetadata = {
    name: '',
    description: '',
    version: '0.0.0',
    tags: [],
    hosts: [],
  };

  let currentKey = '';
  let isMultilineValue = false;
  let multilineValue = '';

  for (const line of yamlBlock.split('\n')) {
    const trimmedLine = line.trim();

    // Array items: "  - value"
    if (trimmedLine.startsWith('- ') && (currentKey === 'tags' || currentKey === 'hosts')) {
      const itemValue = trimmedLine.slice(2).trim();
      if (currentKey === 'tags') {
        metadata.tags.push(itemValue);
      } else {
        const lower = itemValue.toLowerCase();
        if (
          lower === 'excel' ||
          lower === 'powerpoint' ||
          lower === 'word' ||
          lower === 'outlook'
        ) {
          metadata.hosts.push(lower);
        }
      }
      continue;
    }

    // Multiline continuation (indented lines after "key: >")
    if (isMultilineValue && (line.startsWith('  ') || line.startsWith('\t'))) {
      multilineValue += (multilineValue ? ' ' : '') + trimmedLine;
      continue;
    }

    // Flush multiline value
    if (isMultilineValue) {
      setMetadataField(metadata, currentKey, multilineValue);
      isMultilineValue = false;
      multilineValue = '';
    }

    // Key-value pairs
    const colonIndex = trimmedLine.indexOf(':');
    if (colonIndex === -1) continue;

    currentKey = trimmedLine.slice(0, colonIndex).trim();
    const value = trimmedLine.slice(colonIndex + 1).trim();

    if (value === '>' || value === '|') {
      // Multiline scalar
      isMultilineValue = true;
      multilineValue = '';
    } else if (value === '') {
      // Could be start of array (tags: / hosts:) — handled by "- " check above
      continue;
    } else if (currentKey === 'hosts') {
      metadata.hosts = parseInlineArray(value)
        .map(h => h.toLowerCase())
        .filter(h => h === 'excel' || h === 'powerpoint' || h === 'word' || h === 'outlook');
    } else {
      setMetadataField(metadata, currentKey, value);
    }
  }

  // Flush any trailing multiline value
  if (isMultilineValue && multilineValue) {
    setMetadataField(metadata, currentKey, multilineValue);
  }

  return { metadata, content };
}

function parseInlineArray(value: string): string[] {
  const trimmed = value.trim();
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    return trimmed
      .slice(1, -1)
      .split(',')
      .map(item => item.trim())
      .filter(Boolean);
  }
  return [trimmed];
}

function setMetadataField(metadata: SkillMetadata, key: string, value: string): void {
  switch (key) {
    case 'name':
      metadata.name = value;
      break;
    case 'description':
      metadata.description = value;
      break;
    case 'version':
      metadata.version = value;
      break;
    case 'license':
      metadata.license = value;
      break;
    case 'repository':
      metadata.repository = value;
      break;
    case 'documentation':
      metadata.documentation = value;
      break;
  }
}

/** Serialize a skill to its YAML-frontmatter markdown format. */
export function skillToMarkdown(skill: AgentSkill): string {
  const { metadata, content } = skill;
  const lines: string[] = ['---'];
  lines.push(`name: ${metadata.name}`);
  lines.push(`description: ${metadata.description}`);
  lines.push(`version: ${metadata.version}`);
  lines.push('---');
  lines.push('');
  lines.push(content);
  return lines.join('\n');
}
