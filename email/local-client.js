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
 *   message, so In-Reply-To/References, the "Re:" subject, the quoted body,
 *   and the recipient all come from Mail.app itself the way a real reply
 *   would - `to` and `subject` are ignored, since reply() already derives
 *   them correctly and overriding them would defeat the point.
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
    // Mail.app silently drops `set content to ...` on a reply created
    // `without opening window` - the compose editor needs a real text view
    // backing it for edits to stick. Open the window, edit, save/send, then
    // close it - confirmed against a real draft that visible was required
    // (invisible produced a message with no content at all, not even the
    // text this script set).
    script += `      ${findMessageByIdStmts('theMessage', inReplyTo)}\n`;
    script += `      set newMessage to reply theMessage with opening window${replyToAll ? ' with reply to all' : ''}\n`;
    // Mail.app fills in the quoted body asynchronously even with the window
    // open; editing content before that finishes gets silently clobbered
    // when Mail's own population completes a moment later and overwrites
    // the whole field. Poll until content stops changing (i.e. Mail is
    // actually done) instead of guessing a fixed delay - confirmed against
    // a real draft that a blind 1s delay was too short for a longer quoted
    // chain and lost the prepended text entirely.
    script += `      set prevContent to "___unset___"\n`;
    script += `      set stableCount to 0\n`;
    script += `      repeat 20 times\n`;
    script += `        delay 0.5\n`;
    script += `        set curContent to (content of newMessage)\n`;
    script += `        if curContent is prevContent and curContent is not "" then\n`;
    script += `          set stableCount to stableCount + 1\n`;
    script += `          if stableCount > 2 then exit repeat\n`;
    script += `        else\n`;
    script += `          set stableCount to 0\n`;
    script += `        end if\n`;
    script += `        set prevContent to curContent\n`;
    script += `      end repeat\n`;
    script += `      tell newMessage\n`;
    script += `        set content to "${escapeAppleScript(body || '')}" & return & return & content\n`;
    for (const recipient of ccRecipients) {
      script += `        make new cc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
    }
    for (const recipient of bccRecipients) {
      script += `        make new bcc recipient with properties {address:"${escapeAppleScript(recipient)}"}\n`;
    }
    script += `      end tell\n`;
    script += `      ${finalCommand}\n`;
    // Read back what Mail.app actually set, rather than trusting args the
    // caller passed in (which are ignored on this path) - callers can then
    // report what really landed instead of echoing input back as if verified.
    script += `      set resultSubject to subject of newMessage\n`;
    script += `      set resultTo to ""\n`;
    script += `      repeat with r in (to recipients of newMessage)\n`;
    script += `        set resultTo to resultTo & (address of r) & ","\n`;
    script += `      end repeat\n`;
    script += `      set resultContent to content of newMessage\n`;
    script += `      try\n`;
    script += `        close newMessage\n`;
    script += `      end try\n`;
    script += '    end tell\n';
    script += '    return resultSubject & "|||" & resultTo & "|||" & resultContent\n';
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
 * Parse the "subject|||to|||content" return value the inReplyTo path
 * produces into a plain object, or null for the non-reply path (nothing to
 * parse - caller already knows what it sent).
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
  sendEmail,
  saveDraft,
  searchEmails,
  markAsRead,
  listFolders,
  deleteEmail
};
