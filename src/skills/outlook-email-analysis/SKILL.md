---
name: outlook-email-analysis
description: >
  Specialized skill for analyzing, summarizing, and extracting insights from emails
  and attachments. Covers triage, key-point extraction, and thread analysis.
version: 1.0.0
license: MIT
hosts: [outlook]
---

# Email Analysis Skill

Activate this skill when the user wants to understand, summarize, or extract information from emails.

## Analysis Workflow

1. `get_mail_item` → understand sender, subject, recipients, date
2. `get_mail_body` → read full content
3. `get_mail_attachments` → check for attachments
4. `get_attachment_content` → read attachment content if relevant
5. `get_mail_headers` → inspect headers for routing/threading info if needed
6. Deliver analysis to the user

## Analysis Types

### Quick Summary
- Who sent it, when, to whom
- 2–3 sentence summary of the content
- Action items or requests mentioned

### Deep Analysis
- Full thread context (if forwarded/replied)
- Key decisions, deadlines, or commitments
- Sentiment and tone assessment
- Open questions that need a response

### Attachment Analysis
1. `get_mail_attachments` → list all attachments with names and types
2. `get_attachment_content` → read content (first 2000 chars)
3. Summarize what each attachment contains
4. Flag any attachments that need action

### Email Triage
- Classify urgency: urgent / action needed / FYI / low priority
- Identify who is asking for what
- Suggest next steps

## Output Guidelines

- **Be concise** — bullet points over paragraphs
- **Highlight action items** — bold any tasks or deadlines
- **Quote key phrases** — reference exact words when important
- **Flag missing info** — note if context seems incomplete
