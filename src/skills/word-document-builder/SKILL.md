---
name: word-document-builder
description: >
  Specialized skill for creating and restructuring Word documents. Covers
  multi-section document creation, content planning, and the deep-mode
  planner → worker orchestration workflow.
version: 1.0.0
license: MIT
hosts: [word]
---

# Document Builder Skill

Activate this skill when creating new documents, restructuring existing ones, or building multi-section content.

## Deep Mode — Planner → Worker Orchestration

When triggered by keywords like **"gründlich"**, **"deep"**, **"think"**, **"go deep"**, **"detail"**, **"ausführlich"**, **"thoroughly"**, use the deep planning workflow:

### Phase 1: Plan
1. Read existing document with `get_document_overview` + `get_document_content`
2. Create a structured plan:
   - Document title and purpose
   - Section outline (headings, subheadings)
   - Content summary for each section
   - Formatting approach

### Phase 2: Execute Section by Section
For each planned section:
1. Create heading with `insert_paragraph` + named style
2. Insert content with `insert_content_at_selection` (HTML for rich content)
3. Add tables, lists, images as needed
4. **Verify** with `get_document_section` — read back what was written
5. **Refine** if content is incomplete or formatting is off
6. Move to next section

### Phase 3: Polish
1. `get_document_overview` → verify final structure
2. Check heading hierarchy is consistent
3. Verify spacing and formatting
4. Add headers/footers if appropriate
5. Final summary of what was created

## Create → Verify → Fix Loop (MANDATORY)

For EVERY section you create:
```
1. Write content
2. get_document_section → read back
3. Compare to plan — is it complete? Well-formatted?
4. If issues → fix → read again
5. Only move to next section when current one is right
```

## Document Structure Planning

### Content Types
- **Report**: Title → Executive Summary → Sections → Conclusion
- **Proposal**: Title → Overview → Approach → Timeline → Budget
- **Memo**: To/From/Date → Subject → Body → Action Items
- **Meeting Notes**: Date/Attendees → Agenda → Discussion → Action Items
- **Technical Doc**: Title → Overview → Details → Examples → References

### Layout Variety
Mix content types within sections:
- Paragraphs for narrative
- Bullet lists for key points
- Numbered lists for steps/processes
- Tables for comparisons/data
- Quotes/callouts for emphasis

## Fast Mode (Default)

Without deep triggers, work in a single pass:
1. Read document state
2. Make changes
3. Verify
4. Summarize

## Always-On Rules

- **Plan before writing** — even in fast mode, outline what you'll create
- **Section by section** — never try to write an entire document in one tool call
- **Verify everything** — read back every section after writing
- **Match existing style** — if the document has content, match its tone and formatting
