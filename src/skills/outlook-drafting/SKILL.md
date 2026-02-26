---
name: outlook-drafting
description: >
  Specialized skill for composing, replying, and forwarding emails.
  Covers tone, formatting, recipient management, and attachment handling.
version: 1.0.0
license: MIT
hosts: [outlook]
---

# Email Drafting Skill

Activate this skill when the user wants to write, reply to, or forward emails.

## Drafting Workflow

### Reply or Forward
1. `get_mail_item` → understand context (sender, subject, recipients)
2. `get_mail_body` → read what was said
3. Compose response content
4. `reply_to_mail` or `forward_mail` → send it
5. Confirm what was done

### New Email
1. `display_new_message` → open compose window with recipients, subject, and body
2. Confirm to user

### Compose Mode (editing a draft)
1. `get_mail_item` → check current state
2. `set_mail_subject` → set or update subject
3. `set_mail_body` → set body content (HTML or plain text)
4. `add_mail_recipient` → add To/CC/BCC recipients
5. `add_file_attachment` → attach files if needed
6. `save_draft` → save without sending

## Tone Guidelines

| Context | Tone |
|---------|------|
| Reply to manager | Professional, concise, clear action items |
| Reply to colleague | Friendly but focused, collaborative |
| Reply to external | Formal, complete sentences, sign-off |
| Quick acknowledgment | Brief — "Thanks, will do." is fine |
| Escalation / complaint | Measured, factual, solution-oriented |

**Always match the user's requested tone.** If unspecified, mirror the tone of the original email.

## Formatting Rules

- **Use HTML format** for replies with structure (lists, bold, links)
- **Use plain text** for short, simple responses
- **Keep paragraphs short** — 2–3 sentences max
- **Bold key info** — dates, deadlines, action items
- **Include greeting and sign-off** unless the user says otherwise

## Recipient Management

- `add_mail_recipient` supports `to`, `cc`, and `bcc` fields
- Always confirm recipients before sending
- Warn if replying-all to a large group

## Common Patterns

### Professional reply
```
Hi [Name],

Thanks for your email. [Response to their points.]

[Action item or next step.]

Best regards,
[User name]
```

### Forward with context
```
Hi [Name],

Forwarding this for your review. [Brief context of why.]

[Any specific ask.]

Thanks,
[User name]
```
