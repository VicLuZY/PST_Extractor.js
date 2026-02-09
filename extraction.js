/**
 * PST extraction logic - ported from Python.
 * Uses pst-parser for PST reading.
 */

const TIME_BLOCK_RE = /(.+?)\s+(?:\[)?(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm))(?:\])?:\s*/gi;

function countTokens(text) {
  return Math.max(1, Math.floor(text.length / 4));
}

function sanitizeFilename(name) {
  if (!name || !String(name).trim()) return 'unnamed';
  return String(name)
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .trim()
    .replace(/\.+$/, '')
    .slice(0, 200) || 'unnamed';
}

function isTeamsOrSkype(rec) {
  const source = (rec.source || '').toLowerCase();
  const body = (rec.body || '').toLowerCase();
  const subject = (rec.subject || '').toLowerCase();
  const toStr = (rec.to || '').toLowerCase();
  const msgClass = (rec.message_class || '').toLowerCase();
  if (/conversation|teams|skype/.test(msgClass)) return true;
  if (/conversation-history|conversation history/.test(source)) return true;
  const indicators = [
    'teams.microsoft.com', 'skype for business', 'conversation with',
    'duration:', ' minutes ', ' anonymous.invalid', 'thread.skype',
  ];
  const combined = body + ' ' + subject + ' ' + toStr;
  return indicators.some(ind => combined.includes(ind));
}

function parseChatLines(body) {
  if (!body || !String(body).trim()) return [];
  const text = String(body)
    .replace(/&nbsp;/g, ' ')
    .replace(/\t/g, ' ')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)));
  const messages = [];
  const regex = new RegExp(TIME_BLOCK_RE.source, 'gi');
  let m;
  const matches = [];
  while ((m = regex.exec(text)) !== null) matches.push(m);
  for (let i = 0; i < matches.length; i++) {
    const g = matches[i];
    const senderRaw = g[1].trim();
    const timeStr = g[2].trim();
    const start = g.index + g[0].length;
    const end = i + 1 < matches.length ? matches[i + 1].index : text.length;
    let msgText = text.slice(start, end).trim();
    if (i + 1 < matches.length && msgText.endsWith(matches[i + 1][1].trim())) {
      msgText = msgText.slice(0, -matches[i + 1][1].trim().length).trim();
    }
    if (!msgText || msgText.slice(0, 50).includes('Duration:')) continue;
    const isEmail = senderRaw.includes('@') && senderRaw.split('@').pop().includes('.');
    if (!isEmail && (senderRaw.length > 40 || !senderRaw || !/^[A-Z]/.test(senderRaw))) continue;
    messages.push({
      sender_email: isEmail ? senderRaw : null,
      sender: isEmail ? null : senderRaw,
      time: timeStr,
      text: msgText.slice(0, 5000),
    });
  }
  return messages;
}

function inferPlatform(rec) {
  const body = (rec.body || '').toLowerCase();
  if (body.includes('teams.microsoft.com') || body.includes('thread.skype')) return 'teams';
  if (body.includes('skype for business')) return 'skype';
  return 'teams_or_skype';
}

function buildConversationId(rec) {
  const raw = `${rec.subject || ''}|${rec.from || ''}|${rec.to || ''}|${rec.date || ''}`;
  let h = 0;
  for (let i = 0; i < raw.length; i++) h = ((h << 5) - h + raw.charCodeAt(i)) | 0;
  return String(h >>> 0);
}

function extractAndNormalizeTeams(rec) {
  const body = rec.body || '';
  const parsed = parseChatLines(body);
  if (!parsed.length) {
    if (body.trim()) {
      return [{
        source_file: rec.source || '',
        conversation_id: buildConversationId(rec),
        subject: rec.subject,
        outlook_date: rec.date,
        platform: inferPlatform(rec),
        text: body.slice(0, 10000),
        is_parsed: false,
      }];
    }
    return [];
  }
  return parsed.map(p => ({
    source_file: rec.source || '',
    conversation_id: buildConversationId(rec),
    subject: rec.subject,
    outlook_date: rec.date,
    platform: inferPlatform(rec),
    sender: p.sender,
    sender_email: p.sender_email,
    message_time: p.time,
    text: p.text || '',
    is_parsed: true,
  }));
}

function parseHeaders(headerStr) {
  const out = {};
  if (!headerStr) return out;
  const keyRe = /^([A-Za-z\-]+):\s*(.*)$/;
  for (const line of String(headerStr).split('\r\n')) {
    if (!line.trim()) continue;
    const mm = line.match(keyRe);
    if (mm) {
      const key = mm[1].toLowerCase();
      const val = mm[2].trim();
      if (out[key]) out[key] = Array.isArray(out[key]) ? [...out[key], val] : [out[key], val];
      else out[key] = val;
    }
  }
  return out;
}

function str(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'object' && v.buffer) return new TextDecoder().decode(v);
  return String(v);
}

function msgToDict(msg, pstName, folderPath) {
  let body = '';
  try { body = str(msg.body) || str(msg.bodyHTML) || ''; } catch (_) {}
  if (body && typeof msg.bodyHTML === 'string') body = body.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  const props = msg.getAllProperties ? msg.getAllProperties() : {};
  const get = (...keys) => { for (const k of keys) if (props[k] != null) return str(props[k]); return null; };
  let subject = '';
  try { subject = str(msg.subject) || get('37', '0x37', '0037', 'Subject') || ''; } catch (_) {}
  let from = get('0c1a', '0C1A', 'Sender name', '0042', 'Sent representing name') || '';
  let to = get('0e04', 'Display to', '0c1f', 'Sender e-mail address') || '';
  let date = '';
  try {
    const dv = msg.messageDeliveryTime || msg.creationTime || get('0e06', '3007');
    date = dv ? (dv instanceof Date ? dv.toISOString() : str(dv)) : '';
  } catch (_) {}
  let msgId = get('1035', 'Internet message identifier') || '';
  let headers = {};
  try {
    const th = get('007d', '0078', 'Transport message headers');
    if (th) headers = parseHeaders(th);
  } catch (_) {}
  const source = folderPath ? `${pstName}::${folderPath}` : pstName;
  return {
    source,
    message_class: get('001a', 'Message class') || '',
    from: (headers.from || from || '').toString(),
    to: (headers.to || to || '').toString(),
    cc: (headers.cc || '').toString(),
    subject: (subject || '').toString(),
    date: (date || '').toString(),
    message_id: (msgId || '').toString(),
    body: (body || '').toString(),
  };
}

function getAttachmentExt(bytes) {
  if (!bytes || bytes.length < 4) return '.bin';
  if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) return '.pdf';
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return '.jpg';
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return '.png';
  return '.bin';
}

/**
 * Extract one PST file. Returns { messages, attachments }.
 * attachments: [{ folderPath, name, data }]
 */
export async function extractPst(buffer, pstName, PST) {
  const messages = [];
  const attachments = [];
  const warnings = [];
  const warn = (ctx, err) => {
    const msg = err && err.message ? err.message : String(err || 'Unknown error');
    warnings.push(`${ctx}: ${msg}`);
    if (typeof console !== 'undefined' && console.warn) {
      console.warn(`[PST extractor] ${ctx}:`, err);
    }
  };
  const pst = new PST.PSTFile(buffer);
  const store = pst.getMessageStore();
  if (!store) throw new Error('No message store');
  const root = store.getRootFolder();
  if (!root) throw new Error('No root folder');

  function walk(folder, folderPath) {
    const name = (folder.displayName || 'Folder').replace(/[^\w\-.]/g, '_').slice(0, 80);
    const fullPath = folderPath ? `${folderPath}/${folder.displayName || 'Folder'}` : (folder.displayName || 'Folder');
    let subFolderEntries = [];
    try {
      subFolderEntries = folder.getSubFolderEntries() || [];
    } catch (err) {
      warn(`Skipping subfolders for ${fullPath}`, err);
    }
    for (const entry of subFolderEntries) {
      try {
        const sub = folder.getSubFolder(entry.nid);
        if (sub) walk(sub, fullPath);
      } catch (err) {
        warn(`Skipping subfolder nid=${entry?.nid} in ${fullPath}`, err);
      }
    }
    let contents = [];
    try {
      contents = folder.getContents() || [];
    } catch (err) {
      warn(`Skipping contents for ${fullPath}`, err);
    }
    for (const entry of contents) {
      let msg = null;
      try {
        msg = folder.getMessage(entry.nid);
      } catch (err) {
        warn(`Skipping message nid=${entry?.nid} in ${fullPath}`, err);
      }
      if (!msg) continue;
      let rec;
      try {
        rec = msgToDict(msg, pstName, fullPath);
      } catch (err) {
        warn(`Skipping message record nid=${entry?.nid} in ${fullPath}`, err);
        continue;
      }
      let props = {};
      try {
        props = msg.getAllProperties ? msg.getAllProperties() : {};
      } catch (err) {
        warn(`Unable to read properties for nid=${entry?.nid} in ${fullPath}`, err);
      }
      const getProp = (...keys) => { for (const k of keys) if (props[k] != null) return props[k]; return null; };
      if (!rec.from) rec.from = getProp('0c1a', '0C1A', '0x0c1a', 'Sender name', 'Sender entry name');
      if (!rec.to) {
        let r = [];
        try {
          r = msg.getAllRecipients?.() || [];
        } catch (err) {
          warn(`Unable to read recipients for nid=${entry?.nid} in ${fullPath}`, err);
        }
        rec.to = getProp('0e04', '0E04', 'Display to', '0c1f', '0C1F') || (r.length ? r.map(x => x['3003'] || x['0c1f'] || x['Email address'] || x['3001'] || '').join('; ') : '');
      }
      if (!rec.subject) rec.subject = getProp('37', '0x37', '0037', 'Subject');
      if (!rec.body) {
        let html = '';
        try {
          html = msg.bodyHTML ? String(msg.bodyHTML).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';
        } catch (err) {
          warn(`Unable to read bodyHTML for nid=${entry?.nid} in ${fullPath}`, err);
        }
        rec.body = getProp('1000', '0x1000', '1000', 'Body', 'Plain text message body') || html;
      }
      messages.push(rec);
      let entries = [];
      try {
        entries = msg.getAttachmentEntries ? msg.getAttachmentEntries() : [];
      } catch (err) {
        warn(`Unable to read attachment entries for nid=${entry?.nid} in ${fullPath}`, err);
      }
      for (let i = 0; i < entries.length; i++) {
        try {
          const att = msg.getAttachment(i);
          if (!att) continue;
          const attProps = typeof att === 'object' ? att : {};
          const dataKeys = ['3701', '0x3701', 'Attachment binary data', 'Attachment object', 'data'];
          let raw = null;
          for (const k of dataKeys) if (attProps[k]) { raw = attProps[k]; break; }
          if (!raw) continue;
          const bytes = raw instanceof Uint8Array ? raw : raw instanceof ArrayBuffer ? new Uint8Array(raw) : new Uint8Array(raw.buffer || raw);
          if (!bytes.length) continue;
          const ext = bytes.length >= 4 ? getAttachmentExt(bytes) : '.bin';
          const fnameKeys = ['3704', '3707', '0x3704', 'Attachment (short) filename', 'Attachment long filename', 'filename', '3001', 'Display name'];
          let base = `attachment_${i}`;
          for (const k of fnameKeys) if (attProps[k]) { base = attProps[k]; break; }
          const safe = sanitizeFilename(typeof base === 'string' ? base : 'attachment') || 'attachment';
          const extFromName = /\.([a-z0-9]{2,6})$/i.exec(safe);
          const finalName = extFromName ? safe : `${safe}${ext}`;
          attachments.push({ folderPath: name, name: finalName, data: bytes });
        } catch (err) {
          warn(`Skipping attachment index=${i} for nid=${entry?.nid} in ${fullPath}`, err);
        }
      }
    }
  }

  walk(root, '');
  if (!messages.length && warnings.length) {
    throw new Error(`Unable to parse PST entries (${warnings[0]})`);
  }
  return { messages, attachments };
}

/**
 * Build emails JSONL, teams JSONL, and add to ZIP.
 */
export function buildOutput(messages, attachments, zip, basePath) {
  const MAX_TOKENS = 2500;
  let fileIdx = 0;
  let currentBatch = [];
  let currentTokens = 0;
  for (const m of messages) {
    const j = JSON.stringify(m);
    const t = countTokens(j);
    if (t > MAX_TOKENS && currentBatch.length) {
      zip.file(`${basePath}/emails_jsonl/mail_${String(fileIdx++).padStart(4, '0')}.jsonl`, currentBatch.map(x => JSON.stringify(x)).join('\n') + '\n');
      currentBatch = [];
      currentTokens = 0;
    }
    if (t > MAX_TOKENS) {
      zip.file(`${basePath}/emails_jsonl/mail_${String(fileIdx++).padStart(4, '0')}.jsonl`, j + '\n');
      continue;
    }
    if (currentTokens + t > MAX_TOKENS && currentBatch.length) {
      zip.file(`${basePath}/emails_jsonl/mail_${String(fileIdx++).padStart(4, '0')}.jsonl`, currentBatch.map(x => JSON.stringify(x)).join('\n') + '\n');
      currentBatch = [];
      currentTokens = 0;
    }
    currentBatch.push(m);
    currentTokens += t;
  }
  if (currentBatch.length) {
    zip.file(`${basePath}/emails_jsonl/mail_${String(fileIdx++).padStart(4, '0')}.jsonl`, currentBatch.map(x => JSON.stringify(x)).join('\n') + '\n');
  }
  const teams = [];
  for (const rec of messages) {
    if (!isTeamsOrSkype(rec)) continue;
    teams.push(...extractAndNormalizeTeams(rec));
  }
  if (teams.length) {
    zip.file(`${basePath}/teams_messages/teams_messages.jsonl`, teams.map(x => JSON.stringify(x)).join('\n') + '\n');
  }
  const seen = new Set();
  for (const a of attachments) {
    const path = `${basePath}/attachments/${a.folderPath}/${a.name}`;
    let p = path;
    let idx = 0;
    while (seen.has(p)) p = `${basePath}/attachments/${a.folderPath}/${a.name.replace(/(\.[^.]+)$/, `_${++idx}$1`)}`;
    seen.add(p);
    zip.file(p, a.data);
  }
  return { emails: messages.length, attachments: attachments.length, teams_messages: teams.length };
}
