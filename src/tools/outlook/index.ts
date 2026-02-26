import type { Tool, ToolResultObject } from '@github/copilot-sdk';

function fail(e: unknown): ToolResultObject {
  const msg = e instanceof Error ? e.message : String(e);
  return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
}

function getMailbox(): Office.Mailbox {
  if (typeof Office === 'undefined' || !Office.context?.mailbox) {
    throw new Error('Outlook mailbox API is not available.');
  }
  return Office.context.mailbox;
}

const getMailItem: Tool = {
  name: 'get_mail_item',
  description:
    'Get an overview of the current mail item: subject, sender, recipients (to/cc/bcc), date, and item type (read or compose).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const itemType = item.itemType;
      const isRead = itemType === (Office.MailboxEnums.ItemType.Message as string);

      const lines: string[] = [`Mail Item Overview`, `${'='.repeat(40)}`];
      lines.push(`Type: ${String(itemType)}`);

      if (item.subject) {
        if (typeof item.subject === 'string') {
          lines.push(`Subject: ${item.subject}`);
        } else {
          const subject = await new Promise<string>((resolve, reject) => {
            (item.subject as Office.Subject).getAsync(result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
              else reject(new Error(result.error?.message ?? 'Failed to get subject'));
            });
          });
          lines.push(`Subject: ${subject}`);
        }
      }

      if (isRead && item.from) {
        const from = item.from as Office.EmailAddressDetails;
        lines.push(`From: ${from.displayName} <${from.emailAddress}>`);
      }

      if (item.to) {
        if (Array.isArray(item.to)) {
          const recipients = (item.to as Office.EmailAddressDetails[])
            .map(r => `${r.displayName} <${r.emailAddress}>`)
            .join(', ');
          lines.push(`To: ${recipients}`);
        } else {
          const recipients = await new Promise<Office.EmailAddressDetails[]>((resolve, reject) => {
            (item.to as Office.Recipients).getAsync(result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
              else reject(new Error(result.error?.message ?? 'Failed to get recipients'));
            });
          });
          lines.push(
            `To: ${recipients.map(r => `${r.displayName} <${r.emailAddress}>`).join(', ')}`
          );
        }
      }

      if (isRead && item.dateTimeCreated) {
        lines.push(`Date: ${String(item.dateTimeCreated)}`);
      }

      lines.push(`Mode: ${isRead ? 'read' : 'compose'}`);

      return lines.join('\n');
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailBody: Tool = {
  name: 'get_mail_body',
  description: 'Get the full body content of the current mail item as HTML or plain text.',
  parameters: {
    type: 'object',
    properties: {
      format: {
        type: 'string',
        enum: ['html', 'text'],
        description: 'Body format to retrieve. Defaults to "text".',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const format = a.format === 'html' ? Office.CoercionType.Html : Office.CoercionType.Text;

      const body = await new Promise<string>((resolve, reject) => {
        item.body.getAsync(format, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message ?? 'Failed to get body'));
        });
      });

      return body;
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailAttachments: Tool = {
  name: 'get_mail_attachments',
  description:
    'List all attachments on the current mail item with name, size, and content type. Does not download attachment content.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      if (!item.attachments || item.attachments.length === 0) {
        return Promise.resolve('No attachments on this mail item.');
      }

      const attachments = item.attachments;
      const lines = [`Attachments (${String(attachments.length)})`, `${'='.repeat(40)}`];
      for (const att of attachments) {
        lines.push(`- ${att.name} (${att.contentType}, ${String(att.size)} bytes)`);
      }
      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const setMailBody: Tool = {
  name: 'set_mail_body',
  description:
    'Set or replace the body content of the current mail item. Only works in compose mode. Supports HTML or plain text.',
  parameters: {
    type: 'object',
    properties: {
      content: {
        type: 'string',
        description: 'The body content to set.',
      },
      format: {
        type: 'string',
        enum: ['html', 'text'],
        description: 'Content format. Defaults to "html".',
      },
    },
    required: ['content'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const content = a.content ?? '';
      const format = a.format === 'text' ? Office.CoercionType.Text : Office.CoercionType.Html;

      await new Promise<void>((resolve, reject) => {
        item.body.setAsync(content, { coercionType: format }, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to set body'));
        });
      });

      return 'Mail body updated successfully.';
    } catch (e) {
      return fail(e);
    }
  },
};

const setMailSubject: Tool = {
  name: 'set_mail_subject',
  description: 'Set the subject of the current mail item. Only works in compose mode.',
  parameters: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'The subject text to set.',
      },
    },
    required: ['subject'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');
      if (!item.subject || typeof item.subject === 'string') {
        throw new Error('Cannot set subject — item is in read mode.');
      }

      const subject = a.subject ?? '';
      await new Promise<void>((resolve, reject) => {
        (item.subject as Office.Subject).setAsync(subject, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to set subject'));
        });
      });

      return `Subject set to: "${subject}"`;
    } catch (e) {
      return fail(e);
    }
  },
};

const addMailRecipient: Tool = {
  name: 'add_mail_recipient',
  description:
    'Add one or more recipients to the current mail item. Only works in compose mode. Specify the field (to, cc, or bcc).',
  parameters: {
    type: 'object',
    properties: {
      field: {
        type: 'string',
        enum: ['to', 'cc', 'bcc'],
        description: 'Recipient field to add to.',
      },
      recipients: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            displayName: { type: 'string', description: 'Display name of the recipient.' },
            emailAddress: { type: 'string', description: 'Email address of the recipient.' },
          },
          required: ['emailAddress'],
        },
        description: 'Array of recipients to add.',
      },
    },
    required: ['field', 'recipients'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as {
        field?: string;
        recipients?: { displayName?: string; emailAddress: string }[];
      };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const field = a.field ?? 'to';
      const recipients = a.recipients;
      if (!recipients || recipients.length === 0) throw new Error('No recipients provided.');

      const recipientField = item[field as 'to' | 'cc' | 'bcc'];
      if (!recipientField || Array.isArray(recipientField)) {
        throw new Error(
          `Cannot add recipients — item is in read mode or field "${field}" is not available.`
        );
      }

      const formatted = recipients.map(r => ({
        displayName: r.displayName ?? r.emailAddress,
        emailAddress: r.emailAddress,
      }));

      await new Promise<void>((resolve, reject) => {
        recipientField.addAsync(formatted, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to add recipients'));
        });
      });

      return `Added ${String(formatted.length)} recipient(s) to ${field}.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const replyToMail: Tool = {
  name: 'reply_to_mail',
  description:
    'Create a reply to the current mail item with the given HTML content. Only works in read mode. Use replyAll to reply to all recipients.',
  parameters: {
    type: 'object',
    properties: {
      htmlBody: {
        type: 'string',
        description: 'HTML content for the reply body.',
      },
      replyAll: {
        type: 'boolean',
        description: 'If true, reply to all recipients. Defaults to false.',
      },
    },
    required: ['htmlBody'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { htmlBody?: string; replyAll?: boolean };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const htmlBody = a.htmlBody ?? '';
      const replyAll = Boolean(a.replyAll);

      await new Promise<void>((resolve, reject) => {
        const displayCall = replyAll ? item.displayReplyAllForm : item.displayReplyForm;
        if (!displayCall) {
          reject(new Error('Reply is not supported on this item.'));
          return;
        }
        displayCall.call(item, {
          htmlBody,
          callback: (result: Office.AsyncResult<void>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message ?? 'Failed to create reply'));
          },
        });
      });

      return `Reply${replyAll ? ' all' : ''} form opened with provided content.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const forwardMail: Tool = {
  name: 'forward_mail',
  description:
    'Open a forward form for the current mail item with optional HTML content pre-filled. Only works in read mode.',
  parameters: {
    type: 'object',
    properties: {
      htmlBody: {
        type: 'string',
        description: 'Optional HTML content to prepend to the forwarded message.',
      },
    },
    required: [],
  },
  handler: (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { htmlBody?: string };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const htmlBody = a.htmlBody;

      // displayForwardForm is available on MessageRead but not typed in all @types/office-js versions
      const itemRead = item as unknown as {
        displayForwardForm?: (formData: { htmlBody?: string }) => void;
      };
      if (!itemRead.displayForwardForm) {
        throw new Error('Forward is not supported on this item.');
      }

      itemRead.displayForwardForm(htmlBody ? { htmlBody } : {});
      return Promise.resolve('Forward form opened.');
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const getUserProfile: Tool = {
  name: 'get_user_profile',
  description: 'Get the current mailbox user profile information (name, email, time zone).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const userProfile = mailbox.userProfile;

      const lines = [
        `User Profile`,
        `${'='.repeat(40)}`,
        `Name: ${userProfile.displayName}`,
        `Email: ${userProfile.emailAddress}`,
        `Time Zone: ${userProfile.timeZone}`,
      ];
      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const getAttachmentContent: Tool = {
  name: 'get_attachment_content',
  description:
    'Get the content of a specific attachment by index. Returns base64-encoded data for file attachments or item details for item attachments.',
  parameters: {
    type: 'object',
    properties: {
      index: {
        type: 'number',
        description: 'Zero-based index of the attachment to retrieve.',
      },
    },
    required: ['index'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { index?: number };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const attachments = item.attachments;
      if (!attachments || attachments.length === 0) {
        throw new Error('No attachments on this mail item.');
      }
      const idx = a.index ?? 0;
      if (idx < 0 || idx >= attachments.length) {
        throw new Error(
          `Attachment index ${String(idx)} out of range (0-${String(attachments.length - 1)}).`
        );
      }

      const att = attachments[idx];
      const content = await new Promise<string>((resolve, reject) => {
        item.getAttachmentContentAsync(att.id, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const ac = result.value;
            resolve(
              `Name: ${att.name}\nType: ${att.contentType}\nFormat: ${String(ac.format)}\nContent (first 2000 chars):\n${ac.content.slice(0, 2000)}`
            );
          } else {
            reject(new Error(result.error?.message ?? 'Failed to get attachment content'));
          }
        });
      });

      return content;
    } catch (e) {
      return fail(e);
    }
  },
};

const addFileAttachment: Tool = {
  name: 'add_file_attachment',
  description:
    'Add a file attachment to the current mail item by URL. Only works in compose mode. The file is downloaded from the URL by Outlook.',
  parameters: {
    type: 'object',
    properties: {
      uri: {
        type: 'string',
        description: 'The URL of the file to attach.',
      },
      attachmentName: {
        type: 'string',
        description: 'Display name for the attachment.',
      },
    },
    required: ['uri', 'attachmentName'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { uri?: string; attachmentName?: string };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const uri = a.uri ?? '';
      const name = a.attachmentName ?? 'attachment';

      const attachmentId = await new Promise<string>((resolve, reject) => {
        item.addFileAttachmentAsync(uri, name, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message ?? 'Failed to add attachment'));
        });
      });

      return `Attachment "${name}" added (id: ${attachmentId}).`;
    } catch (e) {
      return fail(e);
    }
  },
};

const removeAttachment: Tool = {
  name: 'remove_attachment',
  description:
    'Remove an attachment from the current mail item by its attachment ID. Only works in compose mode.',
  parameters: {
    type: 'object',
    properties: {
      attachmentId: {
        type: 'string',
        description: 'The ID of the attachment to remove.',
      },
    },
    required: ['attachmentId'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { attachmentId?: string };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const id = a.attachmentId ?? '';
      await new Promise<void>((resolve, reject) => {
        item.removeAttachmentAsync(id, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to remove attachment'));
        });
      });

      return `Attachment removed.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailCategories: Tool = {
  name: 'get_mail_categories',
  description: 'Get the categories (color labels) assigned to the current mail item.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const categories = await new Promise<Office.CategoryDetails[]>((resolve, reject) => {
        item.categories.getAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message ?? 'Failed to get categories'));
        });
      });

      if (categories.length === 0) return 'No categories assigned.';

      const lines = [`Categories (${String(categories.length)})`, `${'='.repeat(40)}`];
      for (const cat of categories) {
        lines.push(`- ${cat.displayName} (color: ${cat.color})`);
      }
      return lines.join('\n');
    } catch (e) {
      return fail(e);
    }
  },
};

const setMailCategories: Tool = {
  name: 'set_mail_categories',
  description:
    'Add categories (color labels) to the current mail item. Categories must be in the master category list.',
  parameters: {
    type: 'object',
    properties: {
      categories: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of category names to add.',
      },
    },
    required: ['categories'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { categories?: string[] };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const cats = a.categories ?? [];
      if (cats.length === 0) throw new Error('No categories provided.');

      await new Promise<void>((resolve, reject) => {
        item.categories.addAsync(cats, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to set categories'));
        });
      });

      return `Added ${String(cats.length)} category/categories: ${cats.join(', ')}`;
    } catch (e) {
      return fail(e);
    }
  },
};

const removeMailCategories: Tool = {
  name: 'remove_mail_categories',
  description: 'Remove categories from the current mail item.',
  parameters: {
    type: 'object',
    properties: {
      categories: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of category names to remove.',
      },
    },
    required: ['categories'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { categories?: string[] };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const cats = a.categories ?? [];
      if (cats.length === 0) throw new Error('No categories provided.');

      await new Promise<void>((resolve, reject) => {
        item.categories.removeAsync(cats, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to remove categories'));
        });
      });

      return `Removed ${String(cats.length)} category/categories.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const addNotification: Tool = {
  name: 'add_notification',
  description:
    'Show an informational notification banner on the current mail item. The notification appears at the top of the reading pane or compose form.',
  parameters: {
    type: 'object',
    properties: {
      key: {
        type: 'string',
        description: 'Unique key to identify this notification (for later removal).',
      },
      message: {
        type: 'string',
        description: 'Notification message text.',
      },
      type: {
        type: 'string',
        enum: ['informationalMessage', 'errorMessage', 'insightMessage'],
        description: 'Notification type. Defaults to "informationalMessage".',
      },
      persistent: {
        type: 'boolean',
        description: 'If true, notification persists until removed. Defaults to false.',
      },
    },
    required: ['key', 'message'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { key?: string; message?: string; type?: string; persistent?: boolean };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const key = a.key ?? 'agent-notification';
      const message = a.message ?? '';
      const persistent = a.persistent ?? false;

      const typeMap: Record<string, Office.MailboxEnums.ItemNotificationMessageType> = {
        informationalMessage: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        errorMessage: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        insightMessage: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      };
      const notificationType =
        typeMap[a.type ?? 'informationalMessage'] ??
        Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

      await new Promise<void>((resolve, reject) => {
        item.notificationMessages.addAsync(
          key,
          {
            type: notificationType,
            message,
            persistent,
            icon: 'Icon.16x16',
          },
          result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message ?? 'Failed to add notification'));
          }
        );
      });

      return `Notification "${key}" shown.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const removeNotification: Tool = {
  name: 'remove_notification',
  description: 'Remove a previously shown notification banner by its key.',
  parameters: {
    type: 'object',
    properties: {
      key: {
        type: 'string',
        description: 'The key of the notification to remove.',
      },
    },
    required: ['key'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { key?: string };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      await new Promise<void>((resolve, reject) => {
        item.notificationMessages.removeAsync(a.key ?? '', result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to remove notification'));
        });
      });

      return `Notification removed.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const saveDraft: Tool = {
  name: 'save_draft',
  description: 'Save the current compose item as a draft. Only works in compose mode.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const itemId = await new Promise<string>((resolve, reject) => {
        item.saveAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message ?? 'Failed to save draft'));
        });
      });

      return `Draft saved (id: ${itemId}).`;
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailHeaders: Tool = {
  name: 'get_mail_headers',
  description:
    'Get internet message headers of the current mail item including Message-ID, conversation ID, and item ID.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const lines = [`Mail Headers`, `${'='.repeat(40)}`];

      if (item.itemId) lines.push(`Item ID: ${item.itemId}`);
      if (item.conversationId) lines.push(`Conversation ID: ${item.conversationId}`);
      if (item.internetMessageId) lines.push(`Internet Message ID: ${item.internetMessageId}`);
      if (item.dateTimeCreated) lines.push(`Created: ${String(item.dateTimeCreated)}`);
      if (item.dateTimeModified) lines.push(`Modified: ${String(item.dateTimeModified)}`);

      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const displayNewMessage: Tool = {
  name: 'display_new_message',
  description:
    'Open a new message compose form with optional pre-filled fields (to, cc, subject, body).',
  parameters: {
    type: 'object',
    properties: {
      toRecipients: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of email addresses for the To field.',
      },
      ccRecipients: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of email addresses for the CC field.',
      },
      subject: {
        type: 'string',
        description: 'Subject line for the new message.',
      },
      htmlBody: {
        type: 'string',
        description: 'HTML body content for the new message.',
      },
    },
    required: [],
  },
  handler: (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as {
        toRecipients?: string[];
        ccRecipients?: string[];
        subject?: string;
        htmlBody?: string;
      };
      const mailbox = getMailbox();

      const formData: Record<string, unknown> = {};
      if (a.toRecipients) formData.toRecipients = a.toRecipients;
      if (a.ccRecipients) formData.ccRecipients = a.ccRecipients;
      if (a.subject) formData.subject = a.subject;
      if (a.htmlBody) formData.htmlBody = a.htmlBody;

      mailbox.displayNewMessageForm(formData);
      return Promise.resolve('New message form opened.');
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const displayNewAppointment: Tool = {
  name: 'display_new_appointment',
  description:
    'Open a new appointment/meeting form with optional pre-filled fields (attendees, subject, body, start/end times, location).',
  parameters: {
    type: 'object',
    properties: {
      requiredAttendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of email addresses for required attendees.',
      },
      optionalAttendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of email addresses for optional attendees.',
      },
      subject: {
        type: 'string',
        description: 'Subject/title for the appointment.',
      },
      htmlBody: {
        type: 'string',
        description: 'HTML body/description for the appointment.',
      },
      location: {
        type: 'string',
        description: 'Location of the appointment.',
      },
      start: {
        type: 'string',
        description: 'Start date/time as ISO 8601 string.',
      },
      end: {
        type: 'string',
        description: 'End date/time as ISO 8601 string.',
      },
    },
    required: [],
  },
  handler: (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as {
        requiredAttendees?: string[];
        optionalAttendees?: string[];
        subject?: string;
        htmlBody?: string;
        location?: string;
        start?: string;
        end?: string;
      };
      const mailbox = getMailbox();

      const formData: Record<string, unknown> = {};
      if (a.requiredAttendees) formData.requiredAttendees = a.requiredAttendees;
      if (a.optionalAttendees) formData.optionalAttendees = a.optionalAttendees;
      if (a.subject) formData.subject = a.subject;
      if (a.htmlBody) formData.body = a.htmlBody;
      if (a.location) formData.location = a.location;
      if (a.start) formData.start = new Date(a.start);
      if (a.end) formData.end = new Date(a.end);

      mailbox.displayNewAppointmentForm(formData);
      return Promise.resolve('New appointment form opened.');
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const getDiagnostics: Tool = {
  name: 'get_diagnostics',
  description: 'Get diagnostic information about the Outlook host (host name, version, platform).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const diag = mailbox.diagnostics;

      const lines = [
        `Diagnostics`,
        `${'='.repeat(40)}`,
        `Host Name: ${diag.hostName}`,
        `Host Version: ${diag.hostVersion}`,
      ];
      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

export const outlookTools: Tool[] = [
  getMailItem,
  getMailBody,
  getMailAttachments,
  getAttachmentContent,
  setMailBody,
  setMailSubject,
  addMailRecipient,
  replyToMail,
  forwardMail,
  getUserProfile,
  addFileAttachment,
  removeAttachment,
  getMailCategories,
  setMailCategories,
  removeMailCategories,
  addNotification,
  removeNotification,
  saveDraft,
  getMailHeaders,
  displayNewMessage,
  displayNewAppointment,
  getDiagnostics,
];
