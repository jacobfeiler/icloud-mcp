/**
 * Email module for iCloud MCP
 * Provides email tools via IMAP/SMTP
 */

const { z } = require('zod');
const cloudClient = require('./imap-client');
const smtpClient = require('./smtp-client');
const localClient = require('./local-client');
const { isLocalMode } = require('../mode');
const { hasCredentials } = require('../auth');
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

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Plain text -> HTML fragment, one <br> per newline (line breaks survive). */
function textToBrHtml(s) {
  return escapeHtml(s).replace(/\r\n/g, '\n').replace(/\n/g, '<br>\n');
}

/**
 * Handler: Save draft
 *
 * Writes the draft straight to the iCloud Drafts mailbox via IMAP APPEND
 * (\Draft flag), regardless of the server's local/cloud mode. This is the
 * only route that produces a draft every mail client renders correctly:
 * AppleScript compose on current macOS leaves the text/plain part empty and
 * buries the body in a share-wrapper blockquote (Spark then shows "No
 * Content"). We send multipart/alternative - a plain-text part and an HTML
 * part - because Spark's composer collapses newlines in a text-only draft
 * and won't fold a `> ` quote; the HTML part gives it real <br> breaks and
 * a <blockquote type="cite"> it can collapse. For a reply, the original
 * message is read (locally or over IMAP) for its Message-ID / References
 * (threading) and its body (quoted beneath the reply).
 */
async function handleSaveDraft(args) {
  if (!hasCredentials()) {
    return formatError(new Error(
      'save-draft writes directly to your iCloud Drafts folder over IMAP so the draft threads and renders correctly in every mail client. That needs credentials: set ICLOUD_EMAIL and ICLOUD_APP_PASSWORD (an app-specific password from appleid.apple.com) in the server env or .env.'
    ));
  }

  let { to, cc, bcc, subject } = args;
  const body = args.body || '';
  let inReplyToId = null;
  let references = [];
  let quotedText = '';
  let quotedHtml = '';

  if (args.inReplyTo) {
    const orig = isLocalMode()
      ? await localClient.getMessageForReply(args.inReplyTo)
      : await cloudClient.getMessageForReply(args.inReplyTo, args.folder || 'inbox');

    inReplyToId = orig.messageId || null;
    references = [...(orig.references || []), orig.messageId].filter(Boolean);
    to = to || orig.from;
    subject = /^\s*re:/i.test(orig.subject || '')
      ? orig.subject
      : `Re: ${orig.subject || ''}`.trim();

    const attribution = `On ${orig.dateStr || 'an earlier date'}, ${orig.from || 'someone'} wrote:`;
    const origClean = (orig.text || '')
      .replace(/\r\n/g, '\n')
      .replace(/￼/g, '')       // object-replacement chars (inline images/emoji)
      .replace(/[ \t]+\n/g, '\n')   // trailing whitespace
      .replace(/\n{3,}/g, '\n\n')   // runs of blank lines
      .trim();
    const origLines = origClean.split('\n');
    quotedText = `\n\n${attribution}\n\n${origLines.map(l => `> ${l}`).join('\n')}`;
    quotedHtml =
      `<br><div>${escapeHtml(attribution)}</div>` +
      `<blockquote type="cite" style="margin:0 0 0 0.8ex;border-left:2px solid #ccc;padding-left:1ex">` +
      origLines.map(escapeHtml).join('<br>\n') +
      `</blockquote>`;
  }

  const text = `${body}${quotedText}`;
  const html = `<div>${textToBrHtml(body)}</div>${quotedHtml}`;

  let result;
  try {
    result = await cloudClient.saveDraft({ to, cc, bcc, subject, text, html, inReplyTo: inReplyToId, references });
  } catch (e) {
    if (e.message === 'UNAUTHORIZED') {
      return formatError(new Error('iCloud rejected the credentials - check ICLOUD_EMAIL / ICLOUD_APP_PASSWORD.'));
    }
    throw e;
  }

  const preview = text.length > 500 ? `${text.slice(0, 500)}\n... (truncated)` : text;
  return formatSuccess(
    `Draft saved to "${result.mailbox}" via IMAP (${result.bytes} bytes).\n\n` +
    `${to ? `To: ${to}\n` : ''}${cc ? `CC: ${cc}\n` : ''}Subject: ${subject || '(no subject)'}\n` +
    `${inReplyToId ? `In-Reply-To: ${inReplyToId}\n` : ''}` +
    `\n--- body ---\n${preview}`
  );
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
    description: 'Saves an email as a draft without sending it. Builds the message and writes it straight to your iCloud Drafts mailbox via IMAP APPEND, so it renders and threads correctly in every mail client (Mail, Spark, webmail). Requires credentials (ICLOUD_EMAIL / ICLOUD_APP_PASSWORD) whatever mode the server is in. To draft a reply with correct threading, pass inReplyTo: the original message\'s Message-ID and References chain become In-Reply-To/References, the subject becomes "Re: ...", the recipient is taken from the original sender, and the original body is quoted beneath yours. Plain text only.',
    inputSchema: {
      to: z.string().optional().describe('Recipient email address(es), comma-separated. Optional for a plain draft; for a reply it defaults to the original sender.'),
      cc: z.string().optional().describe('CC recipient(s), comma-separated'),
      bcc: z.string().optional().describe('BCC recipient(s), comma-separated'),
      subject: z.string().optional().describe('Email subject. For a reply it is derived ("Re: ...") if omitted.'),
      body: z.string().optional().describe('Your message text (plain text). For a reply this goes above the quoted original.'),
      inReplyTo: z.string().optional().describe('Handle of the message to reply to (from list-emails/search-emails: a Mail.app id in local mode, an IMAP UID in cloud mode). Drives the threading headers, subject, recipient and quoted body.'),
      folder: z.string().optional().describe('Folder the inReplyTo message is in (cloud mode only, default: inbox)')
    },
    annotations: {"readOnlyHint":false,"destructiveHint":false,"idempotentHint":false,"openWorldHint":true},
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
