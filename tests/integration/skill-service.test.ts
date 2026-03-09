/**
 * Integration tests for parseFrontmatter and skillToMarkdown.
 */

import { describe, it, expect } from 'vitest';
import { parseFrontmatter } from '@/services/skills/skillService';

describe('parseFrontmatter edge cases', () => {
  it('returns defaults when no frontmatter delimiters', () => {
    const { metadata, content } = parseFrontmatter('Just plain text');
    expect(metadata.name).toBe('unknown');
    expect(content).toBe('Just plain text');
  });

  it('returns defaults when closing delimiter is missing', () => {
    const { metadata } = parseFrontmatter('---\nname: test\n');
    expect(metadata.name).toBe('unknown');
  });

  it('parses simple key-value frontmatter', () => {
    const raw = `---
name: my-skill
description: A test skill
version: 1.0.0
---
Body content here`;
    const { metadata, content } = parseFrontmatter(raw);
    expect(metadata.name).toBe('my-skill');
    expect(metadata.description).toBe('A test skill');
    expect(metadata.version).toBe('1.0.0');
    expect(content).toBe('Body content here');
  });

  it('parses tags array', () => {
    const raw = `---
name: tagged
tags:
  - azure
  - excel
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.tags).toEqual(['azure', 'excel']);
  });

  it('ignores non-standard references field', () => {
    const raw = `---
name: with-refs
references:
  - skill:excel
  - agent:Excel
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata).not.toHaveProperty('references');
  });

  it('parses hosts nested under metadata: block', () => {
    const raw = `---
name: nested-host-skill
description: Test
version: 1.0.0
license: MIT
metadata:
  hosts: [excel, word]
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.hosts).toEqual(['excel', 'word']);
  });

  it('normalizes capitalized host names to lowercase in inline array', () => {
    const raw = `---
name: caps-host-skill
description: Test
version: 1.0.0
hosts: [Excel, PowerPoint]
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.hosts).toEqual(['excel', 'powerpoint']);
  });

  it('normalizes capitalized host names to lowercase in list format', () => {
    const raw = `---
name: caps-host-skill
description: Test
version: 1.0.0
hosts:
  - Excel
  - Word
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.hosts).toEqual(['excel', 'word']);
  });

  it('filters out invalid host values during parsing', () => {
    const raw = `---
name: bad-host-skill
description: Test
version: 1.0.0
hosts: [excel, invalidhost, word]
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.hosts).toEqual(['excel', 'word']);
  });
});

