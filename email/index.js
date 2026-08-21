/**
 * Email module for iCloud MCP
 * Provides email tools via IMAP/SMTP
 */

const { z } = require('zod');
const cloudClient = require('./imap-client');
const smtpClient = require('./smtp-client');
const localClient = require('./local-client');
const { isLocalMode } = require('../mode');
const { formatSuccess, formatError, withErrorHandler } = require('../utils/error-handler');
const { listOutput, listResult } = require('../utils/schemas');
const { formatDate, formatRelative } = require('../utils/date-utils');
const config = require('../config');

/**
 * Normalize a message summary to one shape regardless of mode.
 *
 * Cloud (IMAP) returns a uid plus a raw `flags` array; local (Mail.app)
 * returns an id plus a boolean `read`. Downstream sees `ref` (the handle you
 * pass back to read/mark) and `unread`.
 */
function normalizeEmail(email, local) {
  if (local) {
    return { ...email, ref: email.id, unread: email.read === false };
  }
  return { ...email, ref: email.uid, unread: !(email.flags || []).includes('\\Seen') };
}

/**
 * Handler: List emails
 */
async function handleListEmails(args) {
  const folder = args.folder || 'inbox';
  const count = Math.min(args.count || 25, config.DEFAULTS.MAX_RESULTS);

  const local = isLocalMode();
  const client = local ? localClient : cloudClient;

  const emails = (await client.listEmails(folder, count)).map(e => normalizeEmail(e, local));

  if (emails.length === 0) {
    return formatSuccess(`No emails found in ${folder}.`, listResult([]));
  }

  const lines = emails.map((email, i) => {
    const unread = email.unread ? '[UNREAD] ' : '';
    const date = formatRelative(new Date(email.date));
    return `${i + 1}. ${unread}${email.subject}\n   From: ${email.from}\n   Date: ${date}\n   ${local ? 'ID' : 'UID'}: ${email.ref}`;
  });

  return formatSuccess(`Emails in ${folder} (${emails.length}):\n\n${lines.join('\n\n')}`, listResult(emails));
}

/**
 * Handler: Read email
 */
async function handleReadEmail(args) {
  if (!args.uid) {
    return formatError(new Error('Email UID is required'));
  }

  const folder = args.folder || 'inbox';
  const local = isLocalMode();

  const email = local
    ? await localClient.readEmail(args.uid)
    : await cloudClient.readEmail(args.uid, folder);

  // Mail.app returns the body as `body`; mailparser returns `text`.
  let body = email.text || email.body || '';
  if (body.length > config.DEFAULTS.EMAIL_BODY_MAX_LENGTH) {
    body = body.substring(0, config.DEFAULTS.EMAIL_BODY_MAX_LENGTH) + '\n... (truncated)';
  }

  const attachments = email.attachments || [];
  const attachmentInfo = attachments.length > 0
    ? `\n\nAttachments (${attachments.length}):\n${attachments.map(a => `- ${a.filename}${a.size ? ` (${Math.round(a.size / 1024)}KB)` : ''}`).join('\n')}`
    : '';

  return formatSuccess(
    `Subject: ${email.subject}
From: ${email.from}
To: ${email.to}${email.cc ? `\nCC: ${email.cc}` : ''}
Date: ${formatDate(email.date)}
${local ? 'ID' : 'UID'}: ${email.uid || email.id}

---

${body}${attachmentInfo}`
  );
}

/**
 * Handler: Send email
 */
async function handleSendEmail(args) {
  if (!args.to && !args.inReplyTo) {
    return formatError(new Error('Recipient (to) is required, unless inReplyTo is set'));
  }
  if (!args.subject && !args.inReplyTo) {
    return formatError(new Error('Subject is required, unless inReplyTo is set'));
  }
  if (!args.body) {
    return formatError(new Error('Body is required'));
  }
  if (args.inReplyTo && !isLocalMode()) {
    return formatError(new Error('inReplyTo is only supported in LOCAL mode (Mail.app)'));
  }

  const result = await (isLocalMode() ? localClient : smtpClient).sendEmail({
    to: args.to,
    cc: args.cc,
    bcc: args.bcc,
    subject: args.subject,
    body: args.body,
    isHtml: args.isHtml || false,
    inReplyTo: args.inReplyTo,
    replyToAll: args.replyToAll || false
  });

  if (result.success) {
    if (result.reply) {
      // inReplyTo path: report what Mail.app actually set, not the (mostly
      // ignored) args the caller passed in - verifiable instead of assumed.
      return formatSuccess(
        `Email sent successfully!\n\nTo: ${result.reply.to || '(unknown)'}\nSubject: ${result.reply.subject || '(unknown)'}\n\nQuoted content preview:\n${result.reply.contentPreview}`
      );
    }
    return formatSuccess(
      `Email sent successfully!\n\nTo: ${args.to}${args.cc ? `\nCC: ${args.cc}` : ''}\nSubject: ${args.subject}${result.messageId ? `\nMessage ID: ${result.messageId}` : ''}`
    );
  } else {
    return formatError(new Error('Failed to send email'));
  }
}

/**
 * Handler: Save draft
 */
async function handleSaveDraft(args) {
  if (!isLocalMode()) {
    return formatError(new Error('save-draft is only available in LOCAL mode (Mail.app). Cloud mode has no IMAP APPEND support for drafts in this server.'));
  }

  const result = await localClient.saveDraft({
    to: args.to,
    cc: args.cc,
    bcc: args.bcc,
    subject: args.subject,
    body: args.body,
    inReplyTo: args.inReplyTo,
    replyToAll: args.replyToAll || false
  });

  if (result.success) {
    if (result.reply) {
      // inReplyTo path: report what Mail.app actually set, not the (mostly
      // ignored) args the caller passed in - verifiable instead of assumed.
      return formatSuccess(
        `Draft saved successfully!\n\nTo: ${result.reply.to || '(unknown)'}\nSubject: ${result.reply.subject || '(unknown)'}\n\nQuoted content preview:\n${result.reply.contentPreview}`
      );
    }
    return formatSuccess(
      `Draft saved successfully!\n\n${args.to ? `To: ${args.to}\n` : ''}${args.cc ? `CC: ${args.cc}\n` : ''}Subject: ${args.subject || '(no subject)'}`
    );
  } else {
    return formatError(new Error('Failed to save draft'));
  }
}

/**
 * Handler: Search emails
 */
async function handleSearchEmails(args) {
  const folder = args.folder || 'inbox';
  const count = Math.min(args.count || 25, config.DEFAULTS.MAX_RESULTS);

  const criteria = {};
  if (args.from) criteria.from = args.from;
  if (args.subject) criteria.subject = args.subject;
  if (args.query) criteria.text = args.query;
  if (args.unreadOnly) criteria.unseen = true;

  const local = isLocalMode();
  const raw = local
    ? await localClient.searchEmails({ query: args.query, from: args.from, subject: args.subject, folder, count })
    : await cloudClient.searchEmails(criteria, folder, count);
  const emails = raw.map(e => normalizeEmail(e, local));

  if (emails.length === 0) {
    return formatSuccess(`No emails found matching your search criteria in ${folder}.`, listResult([]));
  }

  const lines = emails.map((email, i) => {
    const unread = email.unread ? '[UNREAD] ' : '';
    const date = formatRelative(new Date(email.date));
    return `${i + 1}. ${unread}${email.subject}\n   From: ${email.from}\n   Date: ${date}\n   ${local ? 'ID' : 'UID'}: ${email.ref}`;
  });

  return formatSuccess(`Search results in ${folder} (${emails.length}):\n\n${lines.join('\n\n')}`, listResult(emails));
}

/**
 * Handler: Mark as read
 */
async function handleMarkAsRead(args) {
  if (!args.uid) {
    return formatError(new Error('Email UID is required'));
  }

  const folder = args.folder || 'inbox';
  const isRead = args.isRead !== false;

  if (isLocalMode()) {
    await localClient.markAsRead(args.uid, isRead);
  } else {
    await cloudClient.markAsRead(args.uid, folder, isRead);
  }

  return formatSuccess(`Email ${args.uid} marked as ${isRead ? 'read' : 'unread'}.`);
}

/**
 * Handler: List folders
 */
async function handleListFolders() {
  const folders = isLocalMode()
    ? await localClient.listFolders()
    : await cloudClient.listFolders();

  const lines = folders.map(f => `- ${f.name}`);

  return formatSuccess(`Email folders:\n\n${lines.join('\n')}`, listResult(folders));
}

// Tool definitions
const emailTools = [
  {
    name: 'list-emails',
    outputSchema: listOutput('Emails'),
    title: 'List Emails',
    description: 'Lists emails from a folder (default: inbox)',
    inputSchema: {
      folder: z.string().optional().describe('Email folder (inbox, sent, drafts, trash, archive, junk)'),
      count: z.number().int().min(1).max(50).optional().describe('Number of emails to retrieve (default: 25, max: 50)')
    },
    annotations: {"readOnlyHint":true,"idempotentHint":true,"openWorldHint":true},
    handler: withErrorHandler(handleListEmails, 'list-emails')
  },
  {
    name: 'read-email',
    title: 'Read Email',
    description: 'Reads the full content of an email by UID',
    inputSchema: {
      uid: z.string().describe('Email handle: the UID from list-emails in cloud mode, or the message ID in local mode'),
      folder: z.string().optional().describe('Email folder (default: inbox)')
    },
    annotations: {"readOnlyHint":true,"idempotentHint":true,"openWorldHint":true},
    handler: withErrorHandler(handleReadEmail, 'read-email')
  },
  {
    name: 'send-email',
    title: 'Send Email',
    description: 'Sends an email immediately from the user\'s account to one or more comma-separated recipients, with optional CC, BCC and HTML body. Cloud mode delivers over SMTP and returns the new message ID; local mode sends through Mail.app, which only sends plain text and returns no ID. There is no draft step, so the message goes out as soon as the tool runs. To reply to an existing message with correct threading (In-Reply-To/References headers, quoted body, matching recipient), pass inReplyTo instead of to/subject - LOCAL mode only.',
    inputSchema: {
      to: z.string().optional().describe('Recipient email address(es), comma-separated. Required unless inReplyTo is set (a reply\'s recipient comes from the original message).'),
      cc: z.string().optional().describe('CC recipient(s), comma-separated'),
      bcc: z.string().optional().describe('BCC recipient(s), comma-separated'),
      subject: z.string().optional().describe('Email subject. Required unless inReplyTo is set (a reply\'s subject is derived automatically and this is ignored).'),
      body: z.string().describe('Email body content'),
      isHtml: z.boolean().optional().describe('Whether the body is HTML (default: false)'),
      inReplyTo: z.string().optional().describe('UID of the message to reply to (from list-emails/read-email/search-emails). Uses Mail.app\'s native reply, which sets In-Reply-To/References, quotes the original body beneath yours, and fills in the recipient - to and subject are ignored when this is set. LOCAL mode only.'),
      replyToAll: z.boolean().optional().describe('When inReplyTo is set, also CC every other recipient of the original message, not just the sender. Default: false.')
    },
    annotations: {"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false,"openWorldHint":true},
    handler: withErrorHandler(handleSendEmail, 'send-email')
  },
  {
    name: 'save-draft',
    title: 'Save Email Draft',
    description: 'Saves an email as a draft in Mail.app without sending it. LOCAL mode only - cloud mode has no way to write a draft over IMAP in this server. Same fields as send-email, but nothing is transmitted; the message lands in the Drafts mailbox for later editing/sending from Mail.app. To draft a reply with correct threading (In-Reply-To/References headers, quoted body, matching recipient), pass inReplyTo instead of to/subject.',
    inputSchema: {
      to: z.string().optional().describe('Recipient email address(es), comma-separated (optional for a draft; ignored if inReplyTo is set)'),
      cc: z.string().optional().describe('CC recipient(s), comma-separated'),
      bcc: z.string().optional().describe('BCC recipient(s), comma-separated'),
      subject: z.string().optional().describe('Email subject (ignored if inReplyTo is set - a reply\'s subject is derived automatically)'),
      body: z.string().optional().describe('Email body content'),
      inReplyTo: z.string().optional().describe('UID of the message to reply to (from list-emails/read-email/search-emails). Uses Mail.app\'s native reply, which sets In-Reply-To/References, quotes the original body beneath yours, and fills in the recipient - to and subject are ignored when this is set.'),
      replyToAll: z.boolean().optional().describe('When inReplyTo is set, also CC every other recipient of the original message, not just the sender. Default: false.')
    },
    annotations: {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false,"openWorldHint":false},
    handler: withErrorHandler(handleSaveDraft, 'save-draft')
  },
  {
    name: 'search-emails',
    outputSchema: listOutput('Matching emails'),
    title: 'Search Emails',
    description: 'Searches one mail folder (inbox by default) by free text, sender, subject and unread state; the filters combine as AND. Results carry the same summary fields as list-emails, including the ref handle that read-email and mark-as-read take.',
    inputSchema: {
      query: z.string().optional().describe('Text to search in email content'),
      from: z.string().optional().describe('Filter by sender'),
      subject: z.string().optional().describe('Filter by subject'),
      folder: z.string().optional().describe('Email folder to search (default: inbox)'),
      unreadOnly: z.boolean().optional().describe('Only show unread emails'),
      count: z.number().int().min(1).max(50).optional().describe('Max results (default: 25, max: 50)')
    },
    annotations: {"readOnlyHint":true,"idempotentHint":true,"openWorldHint":true},
    handler: withErrorHandler(handleSearchEmails, 'search-emails')
  },
  {
    name: 'mark-as-read',
    title: 'Mark Email Read/Unread',
    description: 'Marks an email as read or unread',
    inputSchema: {
      uid: z.string().describe('Email handle: the UID from list-emails in cloud mode, or the message ID in local mode'),
      folder: z.string().optional().describe('Email folder (default: inbox)'),
      isRead: z.boolean().optional().describe('Mark as read (true) or unread (false). Default: true')
    },
    annotations: {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":true,"openWorldHint":true},
    handler: withErrorHandler(handleMarkAsRead, 'mark-as-read')
  },
  {
    name: 'list-folders',
    outputSchema: listOutput('Mail folders'),
    title: 'List Mail Folders',
    description: 'Lists all email folders',
    inputSchema: {},
    annotations: {"readOnlyHint":true,"idempotentHint":true,"openWorldHint":true},
    handler: withErrorHandler(handleListFolders, 'list-folders')
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleReadEmail,
  handleSendEmail,
  handleSaveDraft,
  handleSearchEmails,
  handleMarkAsRead,
  handleListFolders
};
