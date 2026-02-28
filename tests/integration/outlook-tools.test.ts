/**
 * Integration tests for Outlook tool schemas.
 *
 * Validates that every Outlook tool has correct JSON Schema parameters,
 * name, description, and handler. Does NOT execute handlers (requires
 * real Office.Mailbox) — execution is covered by E2E tests.
 */
import { describe, it, expect } from 'vitest';
import Ajv from 'ajv';
import { outlookTools } from '@/tools/outlook';

const ajv = new Ajv({ allErrors: true });

function validate(schema: unknown, data: unknown): boolean {
  return Boolean(ajv.compile(schema as object)(data));
}

const toolsByName = Object.fromEntries(outlookTools.map(t => [t.name, t]));

const ALL_TOOL_NAMES = [
  'get_mail_item',
  'get_mail_body',
  'get_mail_attachments',
  'get_attachment_content',
  'set_mail_body',
  'set_mail_subject',
  'add_mail_recipient',
  'reply_to_mail',
  'forward_mail',
  'get_user_profile',
  'add_file_attachment',
  'remove_attachment',
  'get_mail_categories',
  'set_mail_categories',
  'remove_mail_categories',
  'add_notification',
  'remove_notification',
  'save_draft',
  'get_mail_headers',
  'display_new_message',
  'display_new_appointment',
  'get_diagnostics',
] as const;

// ─── Structural ───────────────────────────────────────────────────────────────

describe('Integration: Outlook tools — structural', () => {
  it('outlookTools array contains exactly the expected tools', () => {
    const actual = outlookTools.map(t => t.name).sort();
    const expected = [...ALL_TOOL_NAMES].sort();
    expect(actual).toEqual(expected);
  });

  it('every tool has a non-empty name, description, parameters, and handler', () => {
    for (const tool of outlookTools) {
      expect(tool.name.length).toBeGreaterThan(0);
      expect(tool.description!.length).toBeGreaterThan(0);
      expect(tool.parameters).toBeDefined();
      expect(typeof tool.handler).toBe('function');
    }
  });

  it('every tool parameters schema is a valid JSON Schema object', () => {
    for (const tool of outlookTools) {
      const params = tool.parameters as Record<string, unknown>;
      expect(params.type).toBe('object');
      expect(params.properties).toBeDefined();
    }
  });

  it('no duplicate tool names', () => {
    const names = outlookTools.map(t => t.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

// ─── No-args tools ────────────────────────────────────────────────────────────

describe('Integration: Outlook schema — no-args tools', () => {
  const noArgTools = [
    'get_mail_item',
    'get_mail_attachments',
    'get_user_profile',
    'get_mail_categories',
    'save_draft',
    'get_mail_headers',
    'get_diagnostics',
  ];

  for (const name of noArgTools) {
    it(`${name} accepts empty args`, () => {
      const schema = toolsByName[name].parameters;
      expect(validate(schema, {})).toBe(true);
    });
  }
});

// ─── get_mail_body ────────────────────────────────────────────────────────────

describe('Integration: Outlook schema — get_mail_body', () => {
  const schema = toolsByName.get_mail_body.parameters;

  it('accepts empty args (format defaults)', () => {
    expect(validate(schema, {})).toBe(true);
  });

  it('accepts format: html', () => {
    expect(validate(schema, { format: 'html' })).toBe(true);
  });

  it('accepts format: text', () => {
    expect(validate(schema, { format: 'text' })).toBe(true);
  });

  it('rejects invalid format', () => {
    expect(validate(schema, { format: 'markdown' })).toBe(false);
  });
});

// ─── set_mail_body ────────────────────────────────────────────────────────────

describe('Integration: Outlook schema — set_mail_body', () => {
  const schema = toolsByName.set_mail_body.parameters;

  it('requires content', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { content: '<p>Hello</p>' })).toBe(true);
  });

  it('accepts optional format', () => {
    expect(validate(schema, { content: 'Hello', format: 'text' })).toBe(true);
    expect(validate(schema, { content: '<b>Hi</b>', format: 'html' })).toBe(true);
  });
});

// ─── set_mail_subject ─────────────────────────────────────────────────────────

describe('Integration: Outlook schema — set_mail_subject', () => {
  const schema = toolsByName.set_mail_subject.parameters;

  it('requires subject', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { subject: 'Meeting Notes' })).toBe(true);
  });
});

// ─── add_mail_recipient ───────────────────────────────────────────────────────

describe('Integration: Outlook schema — add_mail_recipient', () => {
  const schema = toolsByName.add_mail_recipient.parameters;

  it('requires field and recipients', () => {
    expect(validate(schema, {})).toBe(false);
    expect(
      validate(schema, {
        field: 'to',
        recipients: [{ emailAddress: 'user@example.com' }],
      })
    ).toBe(true);
  });

  it('accepts displayName in recipients', () => {
    expect(
      validate(schema, {
        field: 'cc',
        recipients: [{ emailAddress: 'user@example.com', displayName: 'User' }],
      })
    ).toBe(true);
  });

  it('accepts valid field values', () => {
    for (const field of ['to', 'cc', 'bcc']) {
      expect(
        validate(schema, {
          field,
          recipients: [{ emailAddress: 'a@b.com' }],
        })
      ).toBe(true);
    }
  });
});

// ─── reply_to_mail ────────────────────────────────────────────────────────────

describe('Integration: Outlook schema — reply_to_mail', () => {
  const schema = toolsByName.reply_to_mail.parameters;

  it('requires htmlBody', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { htmlBody: '<p>Thanks!</p>' })).toBe(true);
  });

  it('accepts optional replyAll', () => {
    expect(validate(schema, { htmlBody: 'Thanks', replyAll: true })).toBe(true);
  });
});

// ─── forward_mail ─────────────────────────────────────────────────────────────

describe('Integration: Outlook schema — forward_mail', () => {
  const schema = toolsByName.forward_mail.parameters;

  it('accepts empty args', () => {
    expect(validate(schema, {})).toBe(true);
  });

  it('accepts optional htmlBody', () => {
    expect(validate(schema, { htmlBody: '<p>FYI</p>' })).toBe(true);
  });
});

// ─── get_attachment_content ───────────────────────────────────────────────────

describe('Integration: Outlook schema — get_attachment_content', () => {
  const schema = toolsByName.get_attachment_content.parameters;

  it('requires index', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { index: 0 })).toBe(true);
  });

  it('rejects non-number index', () => {
    expect(validate(schema, { index: 'first' })).toBe(false);
  });
});

// ─── add_file_attachment ──────────────────────────────────────────────────────

describe('Integration: Outlook schema — add_file_attachment', () => {
  const schema = toolsByName.add_file_attachment.parameters;

  it('requires uri and attachmentName', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { uri: 'https://example.com/file.pdf' })).toBe(false);
    expect(
      validate(schema, { uri: 'https://example.com/file.pdf', attachmentName: 'Report.pdf' })
    ).toBe(true);
  });
});

// ─── remove_attachment ────────────────────────────────────────────────────────

describe('Integration: Outlook schema — remove_attachment', () => {
  const schema = toolsByName.remove_attachment.parameters;

  it('requires attachmentId', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { attachmentId: 'AAMkAGI1AAABwwNAAA=' })).toBe(true);
  });
});

// ─── set_mail_categories ──────────────────────────────────────────────────────

describe('Integration: Outlook schema — set_mail_categories', () => {
  const schema = toolsByName.set_mail_categories.parameters;

  it('requires categories', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { categories: ['Red Category', 'Blue Category'] })).toBe(true);
  });
});

// ─── remove_mail_categories ───────────────────────────────────────────────────

describe('Integration: Outlook schema — remove_mail_categories', () => {
  const schema = toolsByName.remove_mail_categories.parameters;

  it('requires categories', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { categories: ['Red Category'] })).toBe(true);
  });
});

// ─── display_new_message ──────────────────────────────────────────────────────

describe('Integration: Outlook schema — display_new_message', () => {
  const schema = toolsByName.display_new_message.parameters;

  it('accepts empty args', () => {
    expect(validate(schema, {})).toBe(true);
  });

  it('accepts optional fields', () => {
    expect(
      validate(schema, {
        toRecipients: ['user@example.com'],
        subject: 'Hello',
        htmlBody: '<p>Hi!</p>',
      })
    ).toBe(true);
  });
});

// ─── display_new_appointment ──────────────────────────────────────────────────

describe('Integration: Outlook schema — display_new_appointment', () => {
  const schema = toolsByName.display_new_appointment.parameters;

  it('accepts empty args', () => {
    expect(validate(schema, {})).toBe(true);
  });

  it('accepts appointment fields', () => {
    expect(
      validate(schema, {
        subject: 'Team Meeting',
        start: '2025-01-15T10:00:00',
        end: '2025-01-15T11:00:00',
        location: 'Room 101',
      })
    ).toBe(true);
  });
});

// ─── add_notification ─────────────────────────────────────────────────────────

describe('Integration: Outlook schema — add_notification', () => {
  const schema = toolsByName.add_notification.parameters;

  it('requires key and message', () => {
    expect(validate(schema, {})).toBe(false);
    expect(validate(schema, { message: 'Processing complete' })).toBe(false);
    expect(validate(schema, { key: 'notif-1', message: 'Processing complete' })).toBe(true);
  });
});
