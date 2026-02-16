import type { AgentSkill, SkillMetadata } from '@/types/skill';

/**
 * Parse YAML frontmatter from a skill markdown file.
 * Handles the `---` delimited block at the top of the file.
 */
export function parseFrontmatter(raw: string): { metadata: SkillMetadata; content: string } {
  const trimmed = raw.trimStart();

  if (!trimmed.startsWith('---')) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [] },
      content: trimmed,
    };
  }

  const endIndex = trimmed.indexOf('---', 3);
  if (endIndex === -1) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [] },
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
  };

  let currentKey = '';
  let isMultilineValue = false;
  let multilineValue = '';

  for (const line of yamlBlock.split('\n')) {
    const trimmedLine = line.trim();

    // Tag array items: "  - value"
    if (trimmedLine.startsWith('- ') && currentKey === 'tags') {
      metadata.tags.push(trimmedLine.slice(2).trim());
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
      // Could be start of array (tags:) — handled by "- " check above
      continue;
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

/** All bundled skills, parsed at module load time. */
const bundledSkills: AgentSkill[] = [];

/**
 * Get all loaded agent skills.
 */
export function getSkills(): AgentSkill[] {
  return bundledSkills;
}

/**
 * Get a single skill by name.
 */
export function getSkill(name: string): AgentSkill | undefined {
  return bundledSkills.find(s => s.metadata.name === name);
}

/**
 * Build the combined skill context string for injection into the system prompt.
 * @param activeNames — if provided, only include skills whose names are in this list.
 *                       If omitted or empty, all bundled skills are included.
 * Returns an empty string if no skills match.
 */
export function buildSkillContext(activeNames?: string[]): string {
  let skills = getSkills();

  // Filter to active names if provided (empty array = none active)
  if (activeNames !== undefined) {
    skills = skills.filter(s => activeNames.includes(s.metadata.name));
  }

  if (skills.length === 0) return '';

  const sections = skills.map(
    skill =>
      `\n\n---\n## Agent Skill: ${skill.metadata.name}\n${skill.metadata.description}\n\n${skill.content}`
  );

  return `\n\n# Agent Skills\nThe following agent skills provide domain-specific knowledge. Use them to help the user with specialized tasks.${sections.join('')}`;
}
