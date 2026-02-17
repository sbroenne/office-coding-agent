import type { AgentSkill, SkillMetadata } from '@/types/skill';
import { getAllAgents } from '@/services/agents';
import excelSkillRaw from '@/skills/excel/SKILL.md';

/**
 * Parse YAML frontmatter from a skill markdown file.
 * Handles the `---` delimited block at the top of the file.
 */
export function parseFrontmatter(raw: string): { metadata: SkillMetadata; content: string } {
  const trimmed = raw.trimStart();

  if (!trimmed.startsWith('---')) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [], references: [] },
      content: trimmed,
    };
  }

  const endIndex = trimmed.indexOf('---', 3);
  if (endIndex === -1) {
    return {
      metadata: { name: 'unknown', description: '', version: '0.0.0', tags: [], references: [] },
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
    references: [],
  };

  let currentKey = '';
  let isMultilineValue = false;
  let multilineValue = '';

  for (const line of yamlBlock.split('\n')) {
    const trimmedLine = line.trim();

    // Array items: "  - value"
    if (trimmedLine.startsWith('- ') && (currentKey === 'tags' || currentKey === 'references')) {
      const itemValue = trimmedLine.slice(2).trim();
      if (currentKey === 'tags') {
        metadata.tags.push(itemValue);
      } else {
        metadata.references?.push(itemValue);
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
      // Could be start of array (tags:, references:) — handled by "- " check above
      continue;
    } else {
      if (currentKey === 'references') {
        const inlineReferences = parseInlineArray(value);
        metadata.references = [...(metadata.references ?? []), ...inlineReferences];
        continue;
      }
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

  return [trimmed].filter(Boolean);
}

interface InlineReferenceExtraction {
  cleanedContent: string;
  references: string[];
}

function extractInlineReferences(content: string): InlineReferenceExtraction {
  const references: string[] = [];
  const cleanedLines: string[] = [];

  for (const rawLine of content.split('\n')) {
    const trimmedLine = rawLine.trim();
    const directiveMatch = trimmedLine.match(/^@references?\s+(.+)$/i);

    if (directiveMatch) {
      references.push(
        ...directiveMatch[1]
          .split(',')
          .map(item => item.trim())
          .filter(Boolean),
      );
      continue;
    }

    const inlineMatches = [...rawLine.matchAll(/@references?\(([^)]+)\)/gi)];
    if (inlineMatches.length > 0) {
      for (const match of inlineMatches) {
        references.push(
          ...match[1]
            .split(',')
            .map(item => item.trim())
            .filter(Boolean),
        );
      }

      cleanedLines.push(rawLine.replace(/@references?\(([^)]+)\)/gi, '').trimEnd());
      continue;
    }

    cleanedLines.push(rawLine);
  }

  return {
    cleanedContent: cleanedLines.join('\n').trim(),
    references,
  };
}

function normalizeReferenceKey(value: string): string {
  return value.trim().toLowerCase();
}

function resolveSkillContentWithReferences(
  skill: AgentSkill,
  skillsByName: Map<string, AgentSkill>,
  agentsByName: Map<string, string>,
  visitedSkills: Set<string> = new Set(),
): string {
  const skillKey = normalizeReferenceKey(skill.metadata.name);
  if (visitedSkills.has(skillKey)) {
    return extractInlineReferences(skill.content).cleanedContent;
  }

  const nextVisited = new Set(visitedSkills);
  nextVisited.add(skillKey);

  const { cleanedContent, references: inlineReferences } = extractInlineReferences(skill.content);
  const declaredReferences = skill.metadata.references ?? [];
  const allReferences = Array.from(new Set([...declaredReferences, ...inlineReferences]));

  if (allReferences.length === 0) {
    return cleanedContent;
  }

  const resolvedBlocks: string[] = [];

  for (const reference of allReferences) {
    const trimmedReference = reference.trim();
    if (!trimmedReference) continue;

    if (trimmedReference.toLowerCase().startsWith('agent:')) {
      const agentName = trimmedReference.slice('agent:'.length).trim();
      const agentInstructions = agentsByName.get(normalizeReferenceKey(agentName));
      if (agentInstructions) {
        resolvedBlocks.push(`#### Reference: agent:${agentName}\n${agentInstructions}`);
      }
      continue;
    }

    const skillName = trimmedReference.toLowerCase().startsWith('skill:')
      ? trimmedReference.slice('skill:'.length).trim()
      : trimmedReference;

    const referencedSkill = skillsByName.get(normalizeReferenceKey(skillName));
    if (referencedSkill && normalizeReferenceKey(referencedSkill.metadata.name) !== skillKey) {
      const referencedContent = resolveSkillContentWithReferences(
        referencedSkill,
        skillsByName,
        agentsByName,
        nextVisited,
      );

      resolvedBlocks.push(`#### Reference: skill:${referencedSkill.metadata.name}\n${referencedContent}`);
    }
  }

  if (resolvedBlocks.length === 0) {
    return cleanedContent;
  }

  return `${cleanedContent}\n\n### @references\n${resolvedBlocks.join('\n\n')}`;
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
function loadBundledSkills(): AgentSkill[] {
  const loaded: AgentSkill[] = [];

  const webpackRequire =
    typeof require === 'function' ? (require as NodeRequire & { context?: Function }) : undefined;

  if (webpackRequire?.context) {
    const context = webpackRequire.context('../../skills', true, /SKILL\.md$/);
    for (const key of context.keys() as string[]) {
      const raw = context(key) as string;
      const parsed = parseFrontmatter(raw);
      loaded.push({ metadata: parsed.metadata, content: parsed.content });
    }

    return loaded.sort((left, right) => left.metadata.name.localeCompare(right.metadata.name));
  }

  const parsedExcel = parseFrontmatter(excelSkillRaw);
  loaded.push({ metadata: parsedExcel.metadata, content: parsedExcel.content });

  return loaded;
}

const bundledSkills: AgentSkill[] = loadBundledSkills();
let importedSkills: AgentSkill[] = [];

export function getBundledSkills(): AgentSkill[] {
  return bundledSkills;
}

export function getImportedSkills(): AgentSkill[] {
  return importedSkills;
}

export function setImportedSkills(skills: AgentSkill[]): void {
  importedSkills = skills;
}

/**
 * Get all loaded agent skills.
 */
export function getSkills(): AgentSkill[] {
  return [...bundledSkills, ...importedSkills];
}

/**
 * Get a single skill by name.
 */
export function getSkill(name: string): AgentSkill | undefined {
  return getSkills().find(s => s.metadata.name === name);
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

  const allSkills = getSkills();
  const skillsByName = new Map(allSkills.map(skill => [normalizeReferenceKey(skill.metadata.name), skill]));
  const agentsByName = new Map(
    getAllAgents().map(agent => [normalizeReferenceKey(agent.metadata.name), agent.instructions]),
  );

  const sections = skills.map(
    skill =>
      `\n\n---\n## Agent Skill: ${skill.metadata.name}\n${skill.metadata.description}\n\n${resolveSkillContentWithReferences(skill, skillsByName, agentsByName)}`
  );

  return `\n\n# Agent Skills\nThe following agent skills provide domain-specific knowledge. Use them to help the user with specialized tasks.${sections.join('')}`;
}
