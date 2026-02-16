import { describe, it, expect } from 'vitest';
import { parseFrontmatter } from '@/services/skills/skillService';

describe('parseFrontmatter', () => {
  it('parses a complete frontmatter block', () => {
    const raw = `---
name: my-skill
description: A test skill
version: 1.2.3
license: MIT
tags:
  - excel
  - cloud
---
# Body content here`;

    const { metadata, content } = parseFrontmatter(raw);

    expect(metadata.name).toBe('my-skill');
    expect(metadata.description).toBe('A test skill');
    expect(metadata.version).toBe('1.2.3');
    expect(metadata.license).toBe('MIT');
    expect(metadata.tags).toEqual(['excel', 'cloud']);
    expect(content).toBe('# Body content here');
  });

  it('returns defaults when no frontmatter delimiters exist', () => {
    const raw = '# Just a markdown file';
    const { metadata, content } = parseFrontmatter(raw);

    expect(metadata.name).toBe('unknown');
    expect(metadata.description).toBe('');
    expect(metadata.version).toBe('0.0.0');
    expect(metadata.tags).toEqual([]);
    expect(content).toBe('# Just a markdown file');
  });

  it('returns defaults when closing delimiter is missing', () => {
    const raw = `---
name: broken
This never closes`;

    const { metadata } = parseFrontmatter(raw);
    expect(metadata.name).toBe('unknown');
  });

  it('handles multiline description with > scalar', () => {
    const raw = `---
name: multi
description: >
  This is a long
  description that spans
  multiple lines
version: 1.0.0
---
body`;

    const { metadata, content } = parseFrontmatter(raw);
    expect(metadata.name).toBe('multi');
    expect(metadata.description).toBe('This is a long description that spans multiple lines');
    expect(metadata.version).toBe('1.0.0');
    expect(content).toBe('body');
  });

  it('handles empty tags array gracefully', () => {
    const raw = `---
name: no-tags
description: No tags
version: 0.1.0
---
content`;

    const { metadata } = parseFrontmatter(raw);
    expect(metadata.tags).toEqual([]);
  });

  it('trims leading whitespace before parsing', () => {
    const raw = `

---
name: indented
version: 1.0.0
---
body`;

    const { metadata } = parseFrontmatter(raw);
    expect(metadata.name).toBe('indented');
  });

  it('handles optional fields (repository, documentation)', () => {
    const raw = `---
name: full
description: Full metadata
version: 2.0.0
repository: https://github.com/example/repo
documentation: https://docs.example.com
---
body`;

    const { metadata } = parseFrontmatter(raw);
    expect(metadata.repository).toBe('https://github.com/example/repo');
    expect(metadata.documentation).toBe('https://docs.example.com');
  });
});
