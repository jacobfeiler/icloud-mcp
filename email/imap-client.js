/**
 * IMAP client wrapper for iCloud Mail
 */

const Imap = require('imap');
const { simpleParser } = require('mailparser');
const MailComposer = require('nodemailer/lib/mail-composer');
const config = require('../config');
const { getCredentials } = require('../auth');

/**
 * Create IMAP connection
 */
function createConnection() {
  const creds = getCredentials();

  return new Imap({
    user: creds.email,
    password: creds.password,
    host: config.IMAP.HOST,
    port: config.IMAP.PORT,
    tls: config.IMAP.TLS,
    authTimeout: config.IMAP.AUTH_TIMEOUT,
    connTimeout: config.IMAP.CONN_TIMEOUT
  });
}

/**
 * Execute IMAP operation with connection management
 */
function withImap(operation) {
  return new Promise((resolve, reject) => {
    const imap = createConnection();

    imap.once('ready', async () => {
      try {
        const result = await operation(imap);
        imap.end();
        resolve(result);
      } catch (err) {
        imap.end();
        reject(err);
      }
    });

    imap.once('error', (err) => {
      if (/AUTHENTICATIONFAILED|authentication failed|invalid credentials|LOGIN failed/i.test(err.message || '')) {
        reject(new Error('UNAUTHORIZED'));
      } else {
        reject(err);
      }
    });

    imap.connect();
  });
}

/**
 * Get folder name from user-friendly name
 */
function getFolderName(folder) {
  const lower = (folder || 'inbox').toLowerCase();
  return config.EMAIL_FOLDERS[lower] || folder;
}

/**
 * List emails from a folder
 */
async function listEmails(folder = 'inbox', count = 25) {
  return withImap(async (imap) => {
    return new Promise((resolve, reject) => {
      const folderName = getFolderName(folder);

      imap.openBox(folderName, true, (err, box) => {
        if (err) {
          reject(err);
          return;
        }

        const total = box.messages.total;
        if (total === 0) {
          resolve([]);
          return;
        }

        // Fetch most recent emails
        const start = Math.max(1, total - count + 1);
        const range = `${start}:${total}`;

        const emails = [];
        const fetch = imap.seq.fetch(range, {
          bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'],
          struct: true
        });

        fetch.on('message', (msg, seqno) => {
          const email = { seqno, uid: null };

          msg.on('body', (stream, info) => {
            let buffer = '';
            stream.on('data', (chunk) => buffer += chunk.toString('utf8'));
            stream.on('end', () => {
              const headers = Imap.parseHeader(buffer);
              email.from = headers.from?.[0] || '';
              email.to = headers.to?.[0] || '';
              email.subject = headers.subject?.[0] || '(No subject)';
              email.date = headers.date?.[0] || '';
            });
          });

          msg.once('attributes', (attrs) => {
            email.uid = attrs.uid;
            email.flags = attrs.flags || [];
          });

          msg.once('end', () => {
            emails.push(email);
          });
        });

        fetch.once('error', reject);
        fetch.once('end', () => {
          // Sort by date descending (newest first)
          emails.sort((a, b) => new Date(b.date) - new Date(a.date));
          resolve(emails);
        });
      });
    });
  });
}

/**
 * Read full email content
 */
async function readEmail(uid, folder = 'inbox') {
  return withImap(async (imap) => {
    return new Promise((resolve, reject) => {
      const folderName = getFolderName(folder);

      imap.openBox(folderName, true, (err) => {
        if (err) {
          reject(err);
          return;
        }

        const fetch = imap.fetch(uid, { bodies: '' });
        let rawEmail = '';

        fetch.on('message', (msg) => {
          msg.on('body', (stream) => {
            stream.on('data', (chunk) => rawEmail += chunk.toString('utf8'));
          });
        });

        fetch.once('error', reject);
        fetch.once('end', async () => {
          try {
            const parsed = await simpleParser(rawEmail);
            const refs = Array.isArray(parsed.references)
              ? parsed.references
              : (parsed.references ? [parsed.references] : []);
            resolve({
              uid,
              from: parsed.from?.text || '',
              to: parsed.to?.text || '',
              cc: parsed.cc?.text || '',
              subject: parsed.subject || '(No subject)',
              date: parsed.date,
              messageId: parsed.messageId || '',
              references: refs,
              text: parsed.text || '',
              html: parsed.html || '',
              attachments: (parsed.attachments || []).map(a => ({
                filename: a.filename,
                contentType: a.contentType,
                size: a.size
              }))
            });
          } catch (parseErr) {
            reject(parseErr);
          }
        });
      });
    });
  });
}

/**
 * Search emails
 */
async function searchEmails(criteria, folder = 'inbox', count = 25) {
  return withImap(async (imap) => {
    return new Promise((resolve, reject) => {
      const folderName = getFolderName(folder);

      imap.openBox(folderName, true, (err, box) => {
        if (err) {
          reject(err);
          return;
        }

        // Build IMAP search criteria
        const searchCriteria = [];

        if (criteria.from) {
          searchCriteria.push(['FROM', criteria.from]);
        }
        if (criteria.subject) {
          searchCriteria.push(['SUBJECT', criteria.subject]);
        }
        if (criteria.since) {
          searchCriteria.push(['SINCE', criteria.since]);
        }
        if (criteria.before) {
          searchCriteria.push(['BEFORE', criteria.before]);
        }
        if (criteria.unseen) {
          searchCriteria.push('UNSEEN');
        }
        if (criteria.text) {
          searchCriteria.push(['TEXT', criteria.text]);
        }

        // Default to ALL if no criteria
        if (searchCriteria.length === 0) {
          searchCriteria.push('ALL');
        }

        imap.search(searchCriteria, (err, uids) => {
          if (err) {
            reject(err);
            return;
          }

          if (!uids || uids.length === 0) {
            resolve([]);
            return;
          }

          // Get most recent results
          const recentUids = uids.slice(-count);

          const emails = [];
          const fetch = imap.fetch(recentUids, {
            bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'],
            struct: true
          });

          fetch.on('message', (msg, seqno) => {
            const email = { seqno };

            msg.on('body', (stream) => {
              let buffer = '';
              stream.on('data', (chunk) => buffer += chunk.toString('utf8'));
              stream.on('end', () => {
                const headers = Imap.parseHeader(buffer);
                email.from = headers.from?.[0] || '';
                email.to = headers.to?.[0] || '';
                email.subject = headers.subject?.[0] || '(No subject)';
                email.date = headers.date?.[0] || '';
              });
            });

            msg.once('attributes', (attrs) => {
              email.uid = attrs.uid;
              email.flags = attrs.flags || [];
            });

            msg.once('end', () => {
              emails.push(email);
            });
          });

          fetch.once('error', reject);
          fetch.once('end', () => {
            emails.sort((a, b) => new Date(b.date) - new Date(a.date));
            resolve(emails);
          });
        });
      });
    });
  });
}

/**
 * Mark email as read/unread
 */
async function markAsRead(uid, folder = 'inbox', isRead = true) {
  return withImap(async (imap) => {
    return new Promise((resolve, reject) => {
      const folderName = getFolderName(folder);

      imap.openBox(folderName, false, (err) => {
        if (err) {
          reject(err);
          return;
        }

        const flags = ['\\Seen'];
        const method = isRead ? 'addFlags' : 'delFlags';

        imap[method](uid, flags, (err) => {
          if (err) {
            reject(err);
          } else {
            resolve(true);
          }
        });
      });
    });
  });
}

/**
 * List folders
 */
async function listFolders() {
  return withImap(async (imap) => {
    return new Promise((resolve, reject) => {
      imap.getBoxes((err, boxes) => {
        if (err) {
          reject(err);
          return;
        }

        const folders = [];

        function processBoxes(boxObj, prefix = '') {
          for (const [name, box] of Object.entries(boxObj)) {
            const fullPath = prefix ? `${prefix}${box.delimiter}${name}` : name;
            folders.push({
              name: fullPath,
              delimiter: box.delimiter,
              flags: box.attribs || []
            });

            if (box.children) {
              processBoxes(box.children, fullPath);
            }
          }
        }

        processBoxes(boxes);
        resolve(folders);
      });
    });
  });
}

/**
 * Pull just the fields needed to draft a threaded reply to `uid`:
 * the original Message-ID and its References chain (for In-Reply-To /
 * References), plus sender / subject / date / body for the quote block.
 */
async function getMessageForReply(uid, folder = 'inbox') {
  const m = await readEmail(uid, folder);
  return {
    messageId: m.messageId || '',
    references: m.references || [],
    from: m.from || '',
    subject: m.subject || '',
    dateStr: m.date ? new Date(m.date).toUTCString() : '',
    text: m.text || ''
  };
}

/**
 * Walk the getBoxes() tree for a mailbox carrying an RFC 6154 special-use
 * attribute (e.g. "\\Drafts"). Returns the full path or null.
 */
function findSpecialUse(boxes, flag, prefix = '') {
  for (const [name, box] of Object.entries(boxes || {})) {
    const full = prefix ? `${prefix}${box.delimiter}${name}` : name;
    const attribs = box.attribs || box.attribute || [];
    if (attribs.includes(flag) || box.special_use_attrib === flag) return full;
    if (box.children) {
      const hit = findSpecialUse(box.children, flag, full);
      if (hit) return hit;
    }
  }
  return null;
}

function resolveDraftsMailbox(imap) {
  return new Promise((resolve) => {
    imap.getBoxes((err, boxes) => {
      if (err || !boxes) return resolve(config.EMAIL_FOLDERS.drafts || 'Drafts');
      resolve(findSpecialUse(boxes, '\\Drafts') || config.EMAIL_FOLDERS.drafts || 'Drafts');
    });
  });
}

/**
 * Build a MIME message with nodemailer's MailComposer and APPEND it to the
 * Drafts mailbox with the \Draft flag - a real IMAP draft that every mail
 * client (Mail, Spark, webmail) renders correctly, unlike an AppleScript
 * compose. `inReplyTo` / `references` produce the threading headers.
 */
async function saveDraft({ from, to, cc, bcc, subject, text, html, inReplyTo, references }) {
  const creds = getCredentials();

  const composer = new MailComposer({
    from: from || creds.email,
    to: to || undefined,
    cc: cc || undefined,
    bcc: bcc || undefined,
    subject: subject || '',
    text: text != null ? text : undefined,
    html: html || undefined,
    inReplyTo: inReplyTo || undefined,
    references: (references && references.length) ? references : undefined
  });

  const mime = await new Promise((resolve, reject) => {
    composer.compile().build((err, msg) => (err ? reject(err) : resolve(msg)));
  });

  return withImap(async (imap) => {
    const mailbox = await resolveDraftsMailbox(imap);
    return new Promise((resolve, reject) => {
      // No `date` option: node-imap@0.8.19 gates it behind `util.isDate`,
      // which Node 23+ removed, so passing it throws "isDate is not a
      // function". The Date header in the MIME already carries the time;
      // the IMAP internal-date is optional and the server fills it in.
      imap.append(mime, { mailbox, flags: ['\\Draft'] }, (err) => {
        if (err) reject(err);
        else resolve({ mailbox, bytes: mime.length });
      });
    });
  });
}

module.exports = {
  listEmails,
  readEmail,
  getMessageForReply,
  saveDraft,
  searchEmails,
  markAsRead,
  listFolders,
  getFolderName
};
