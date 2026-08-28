/**
 * Local Email Client
 * Accesses Mail.app via AppleScript
 */

const { runAppleScript, runJXA, escapeAppleScript, escapeJXA, asInt, asBool } = require('../utils/applescript');
const config = require('../config');

/**
 * List emails from a mailbox
 * @param {string} folder - Folder name (inbox, sent, drafts, etc.)
 * @param {number} count - Number of emails to retrieve
 * @returns {Promise<Array>} - List of emails
 */
async function listEmails(folder = 'inbox', count = 25) {
  const mailboxName = getMailboxName(folder);

  const script = `
    ObjC.import('Foundation');
    const mail = Application('Mail');
    const accounts = mail.accounts();
    let emails = [];

    for (let account of accounts) {
      try {
        const mailbox = account.mailboxes.byName('${escapeJXA(mailboxName)}');
        const messages = mailbox.messages();
        const limit = Math.min(${asInt(count, 25)}, messages.length);

        for (let i = 0; i < limit; i++) {
          const msg = messages[i];
          emails.push({
            id: msg.id(),
            subject: msg.subject(),
            from: msg.sender(),
            date: msg.dateReceived().toISOString(),
            read: msg.readStatus(),
            account: account.name()
          });
        }
      } catch (e) {
        // Mailbox might not exist in this account
      }
    }

    JSON.stringify(emails.slice(0, ${count}));
  `;

  const result = await runJXA(script);
  return result ? JSON.parse(result) : [];
}

/**
 * Read a specific email
 * @param {string} emailId - Email ID
 * @returns {Promise<Object>} - Email content
 */
async function readEmail(emailId) {
  const script = `
    const mail = Application('Mail');
    const msg = mail.messages.byId(${asInt(emailId)});

    JSON.stringify({
      id: msg.id(),
      subject: msg.subject(),
      from: msg.sender(),
      to: msg.toRecipients().map(r => r.address()),
      cc: msg.ccRecipients().map(r => r.address()),
      date: msg.dateReceived().toISOString(),
      body: msg.content(),
      read: msg.readStatus()
    });
  `;

  const result = await runJXA(script);
  return result ? JSON.parse(result) : null;
}

/**
 * AppleScript statements that search every mailbox of every account for a
 * message by numeric id and leave it in `resultVar`.
 *
 * Mail.app's bare `message id <n>` reference form does not compile on
 * current macOS/Mail.app - the dictionary requires an enclosing mailbox and
 * account (`message id <n> of mailbox <m> of account <a>`), which callers
 * here never have (a UID from list-emails/read-email doesn't carry that).
 * This searches for it instead. Must be run inside `tell application "Mail"`.
 * @param {string} resultVar - Variable name to assign the found message to
 * @param {string|number} id - Message id to search for
 * @returns {string} - AppleScript statements
 */
function findMessageByIdStmts(resultVar, id) {
  const n = asInt(id);
  return [
    `set ${resultVar} to missing value`,
    `repeat with acct in accounts`,
    `  repeat with mb in mailboxes of acct`,
    `    try`,
    `      set ${resultVar} to (first message of mb whose id is ${n})`,
    `      exit repeat`,
    `    end try`,
    `  end repeat`,
    `  if ${resultVar} is not missing value then exit repeat`,
    `end repeat`,
    `if ${resultVar} is missing value then error "Message not found: ${n}"`
  ].join('\n      ');
}

/**
 * Build the AppleScript that produces `newMessage` and then runs finalCommand
 * against it. Two shapes:
 *
 * - inReplyTo set: uses Mail's own `reply` command against the referenced
 *   message, so In-Reply-To/References, the "Re:" subject and the recipient
 *   all come from Mail.app itself the way a real reply would - `to` and
 *   `subject` are ignored, since reply() already derives them correctly and
 *   overriding them would defeat the point. The quoted original is rebuilt
 *   here rather than taken from Mail (see the comment in the branch).
 * - inReplyTo unset: builds a blank outgoing message, as before.
 *
 * @param {Object} options
 * @returns {string} - Full AppleScript
 */
function buildComposeScript({ to, cc, bcc, subject, body, inReplyTo, replyToAll, finalCommand }) {
  const ccRecipients = cc ? (Array.isArray(cc) ? cc : [cc]) : [];
  const bccRecipients = bcc ? (Array.isArray(bcc) ? bcc : [bcc]) : [];

  let script = '\n    tell application "Mail"\n';

  if (inReplyTo) {
    // Mail's `reply` command threads the draft correctly (In-Reply-To /
    // References headers, "Re:" subject, recipient) but on current macOS its
    // `content` property is inert from AppleScript: reads always return ""
    // and a write only takes effect when the message is saved, at which
    // point that single write wins outright - Mail's own asynchronously
    // inserted quoted original is discarded. So there is no way to *append*
    // our text to Mail's quote; the moment we touch `content` at all, the
    // quote is gone. Instead: skip the compose window entirely, rebuild the
    // quoted original ourselves from the source message, and set the whole
    // body (our text + our quote) in one write. Confirmed against real
    // drafts - headers thread correctly in Mail and on the recipient side,
    // and the body persists.
    const escBody = escapeAppleScript(body || '');
    script += `      ${findMessageByIdStmts('theMessage', inReplyTo)}\n`;
    script += `      set replyBody to "${escBody}"\n`;
    // Rebuild "On <date>, <sender> wrote:" + a >-quoted copy of the
    // original. Wrapped in try so a message whose content can't be read
    // (rare, but Mail does throw -1708 on some) still yields a threaded
    // draft carrying at least our text.
    script += `      set quotedOriginal to ""\n`;
    script += `      try\n`;
    script += `        set origSender to sender of theMessage\n`;
    script += `        set origWhen to (date received of theMessage) as string\n`;
    script += `        set origContent to content of theMessage\n`;
    script += `        set quotedOriginal to (linefeed & linefeed & "On " & origWhen & ", " & origSender & " wrote:" & linefeed & linefeed)\n`;
    script += `        repeat with ln in (paragraphs of origContent)\n`;
    script += `          set quotedOriginal to quotedOriginal & "> " & (ln as string) & linefeed\n`;
    script += `        end repeat\n`;
    script += `      end try\n`;
    script += `      set newMessage to reply theMessage without opening window${replyToAll ? ' with reply to all' : ''}\n`;
    script += `      tell newMessage\n`;
    script += `        set content to replyBody & quotedOriginal\n`;
    for (const recipient of ccRecipients) {
      script += `        make new cc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
    }
    for (const recipient of bccRecipients) {
      script += `        make new bcc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
    }
    script += `      end tell\n`;
    script += `      ${finalCommand}\n`;
    // Subject and recipient read back reliably off the outgoing message;
    // `content` does not (see above), so the caller reports the body it
    // passed in for the preview instead.
    script += `      set resultSubject to subject of newMessage\n`;
    script += `      set resultTo to ""\n`;
    script += `      repeat with r in (to recipients of newMessage)\n`;
    script += `        set resultTo to resultTo & (address of r) & ","\n`;
    script += `      end repeat\n`;
    script += '    end tell\n';
    script += `    return resultSubject & "|||" & resultTo & "|||" & "${escapeAppleScript((body || '').slice(0, 200))}"\n`;
    return script;
  }

  const toRecipients = to ? (Array.isArray(to) ? to : [to]) : [];
  script += `      set newMessage to make new outgoing message with properties {subject:"${escapeAppleScript(subject || '')}", content:"${escapeAppleScript(body || '')}", visible:false}\n`;
  script += `      tell newMessage\n`;
  for (const recipient of toRecipients) {
    script += `        make new to recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
  }
  for (const recipient of ccRecipients) {
    script += `        make new cc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
  }
  for (const recipient of bccRecipients) {
    script += `        make new bcc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
  }
  script += `      end tell\n`;
  script += `      ${finalCommand}\n`;
  script += '    end tell\n';
  script += `    return "${finalCommand.startsWith('send') ? 'sent' : 'saved'}"\n`;
  return script;
}

/**
 * Parse the "subject|||to|||bodyPreview" return value the inReplyTo path
 * produces into a plain object, or null for the non-reply path (nothing to
 * parse - caller already knows what it sent). subject and to are read back
 * from Mail; bodyPreview is the text we asked Mail to set (its `content`
 * property can't be read back live on current macOS).
 */
function parseReplyResult(raw) {
  if (typeof raw !== 'string' || !raw.includes('|||')) return null;
  const [subject, toRaw, content] = raw.split('|||');
  return {
    subject,
    to: toRaw.replace(/,$/, ''),
    contentPreview: (content || '').slice(0, 200)
  };
}

/**
 * Send an email
 * @param {Object} options - Email options
 * @returns {Promise<Object>} - Send result
 */
async function sendEmail({ to, cc, bcc, subject, body, inReplyTo, replyToAll }) {
  const script = buildComposeScript({ to, cc, bcc, subject, body, inReplyTo, replyToAll, finalCommand: 'send newMessage' });
  const raw = await runAppleScript(script);
  const reply = parseReplyResult(raw);
  return { success: true, message: 'Email sent successfully', reply };
}

/**
 * Save an email as a draft in Mail.app without sending it
 * @param {Object} options - Email options
 * @returns {Promise<Object>} - Save result
 */
async function saveDraft({ to, cc, bcc, subject, body, inReplyTo, replyToAll }) {
  const script = buildComposeScript({ to, cc, bcc, subject, body, inReplyTo, replyToAll, finalCommand: 'save newMessage' });
  const raw = await runAppleScript(script);
  const reply = parseReplyResult(raw);
  return { success: true, message: 'Draft saved successfully', reply };
}

/**
 * Split `raw` on `delim` into exactly `n` fields, keeping any further
 * occurrences of `delim` inside the final field (so a message body that
 * happens to contain the delimiter survives intact).
 */
function splitFields(raw, delim, n) {
  const out = [];
  let rest = raw;
  for (let i = 0; i < n - 1; i++) {
    const k = rest.indexOf(delim);
    if (k < 0) { out.push(rest); rest = ''; continue; }
    out.push(rest.slice(0, k));
    rest = rest.slice(k + delim.length);
  }
  out.push(rest);
  return out;
}

/**
 * Pull the fields needed to draft a threaded reply to a Mail.app message
 * (identified by its numeric id from list-emails): the original Message-ID
 * and References chain for the threading headers, plus sender / subject /
 * date / body for the quote block. Read straight off the message via
 * AppleScript - only *writing* a reply's content through AppleScript is
 * broken on current macOS, reading is fine.
 */
async function getMessageForReply(id) {
  const DELIM = '-=|=-';
  const script = `
    tell application "Mail"
      ${findMessageByIdStmts('theMessage', id)}
      set msgId to ""
      set refsHdr to ""
      repeat with hh in headers of theMessage
        set hn to (name of hh)
        ignoring case
          if hn is "message-id" then set msgId to (content of hh)
          if hn is "references" then set refsHdr to (content of hh)
        end ignoring
      end repeat
      set theSender to ""
      try
        set theSender to (sender of theMessage)
      end try
      set theSubject to ""
      try
        set theSubject to (subject of theMessage)
      end try
      set theWhen to ""
      try
        set theWhen to ((date sent of theMessage) as string)
      end try
      set theBody to ""
      try
        set theBody to (content of theMessage)
      end try
      return msgId & "${DELIM}" & refsHdr & "${DELIM}" & theSender & "${DELIM}" & theSubject & "${DELIM}" & theWhen & "${DELIM}" & theBody
    end tell
  `;

  const raw = await runAppleScript(script);
  const [messageId, refsHdr, from, subject, dateStr, text] = splitFields(raw, DELIM, 6);
  const references = (refsHdr || '').split(/[\s,]+/).map(s => s.trim()).filter(Boolean);
  return {
    messageId: (messageId || '').trim(),
    references,
    from: (from || '').trim(),
    subject: (subject || '').trim(),
    dateStr: (dateStr || '').trim(),
    text: text || ''
  };
}

/**
 * Search emails
 * @param {Object} options - Search options
 * @returns {Promise<Array>} - Matching emails
 */
async function searchEmails({ query, from, subject, folder = 'inbox', count = 25 }) {
  const mailboxName = getMailboxName(folder);

  let conditions = [];
  if (query) conditions.push(`(msg.subject().toLowerCase().includes("${escapeJXA(query.toLowerCase())}") || msg.content().toLowerCase().includes("${escapeJXA(query.toLowerCase())}"))`);
  if (from) conditions.push(`msg.sender().toLowerCase().includes("${escapeJXA(from.toLowerCase())}")`);
  if (subject) conditions.push(`msg.subject().toLowerCase().includes("${escapeJXA(subject.toLowerCase())}")`);

  const filterCondition = conditions.length > 0 ? conditions.join(' && ') : 'true';

  const script = `
    const mail = Application('Mail');
    const accounts = mail.accounts();
    let emails = [];

    for (let account of accounts) {
      try {
        const mailbox = account.mailboxes.byName('${escapeJXA(mailboxName)}');
        const messages = mailbox.messages();

        for (let i = 0; i < messages.length && emails.length < ${asInt(count, 25)}; i++) {
          const msg = messages[i];
          if (${filterCondition}) {
            emails.push({
              id: msg.id(),
              subject: msg.subject(),
              from: msg.sender(),
              date: msg.dateReceived().toISOString(),
              read: msg.readStatus(),
              account: account.name()
            });
          }
        }
      } catch (e) {}
    }

    JSON.stringify(emails);
  `;

  const result = await runJXA(script);
  return result ? JSON.parse(result) : [];
}

/**
 * Mark email as read/unread
 * @param {string} emailId - Email ID
 * @param {boolean} isRead - Read status
 * @returns {Promise<Object>} - Result
 */
async function markAsRead(emailId, isRead = true) {
  const script = `
    tell application "Mail"
      ${findMessageByIdStmts('theMessage', emailId)}
      set read status of theMessage to ${asBool(isRead, true)}
    end tell
    return "done"
  `;

  await runAppleScript(script);
  return { success: true, message: `Email marked as ${isRead ? 'read' : 'unread'}` };
}

/**
 * List mail folders/mailboxes
 * @returns {Promise<Array>} - List of folders
 */
async function listFolders() {
  const script = `
    const mail = Application('Mail');
    const accounts = mail.accounts();
    let folders = [];

    for (let account of accounts) {
      const mailboxes = account.mailboxes();
      for (let mb of mailboxes) {
        folders.push({
          name: mb.name(),
          account: account.name(),
          unreadCount: mb.unreadCount()
        });
      }
    }

    JSON.stringify(folders);
  `;

  const result = await runJXA(script);
  return result ? JSON.parse(result) : [];
}

/**
 * Delete an email
 * @param {string} emailId - Email ID
 * @returns {Promise<Object>} - Result
 */
async function deleteEmail(emailId) {
  const script = `
    tell application "Mail"
      ${findMessageByIdStmts('theMessage', emailId)}
      delete theMessage
    end tell
    return "deleted"
  `;

  await runAppleScript(script);
  return { success: true, message: 'Email deleted' };
}

/**
 * Map folder names to Mail.app mailbox names
 */
function getMailboxName(folder) {
  const mapping = {
    'inbox': 'INBOX',
    'sent': 'Sent Messages',
    'drafts': 'Drafts',
    'trash': 'Deleted Messages',
    'archive': 'Archive',
    'junk': 'Junk'
  };
  return mapping[folder.toLowerCase()] || folder;
}

module.exports = {
  listEmails,
  readEmail,
  getMessageForReply,
  sendEmail,
  saveDraft,
  searchEmails,
  markAsRead,
  listFolders,
  deleteEmail
};
