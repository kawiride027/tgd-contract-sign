/**
 * TGD Contract Sign — Apps Script backend
 *
 * Deploy as Web app:  Execute as = Me,  Who has access = Anyone
 * Script Properties (Project Settings → Script properties):
 *   SLACK_WEBHOOK_URL  incoming webhook for #signed-contracts-onsite
 *   DRIVER_PIN         shared PIN drivers enter once per phone
 *   (FOLDER_ID / LOG_SHEET_ID are created by setup())
 *
 * The contract wording lives in the Google Docs below — edit the Doc and the
 * app + signed PDFs pick it up automatically. No redeploy needed.
 */

const CONTRACTS = {
  residential: {
    name: 'Residential Dumpster Rental Agreement',
    docId: '1e9HFEBcAWAL3mAEBhENXMpXw1HE-Ey5ylqKDi7qtGxs',
  },
  contractor: {
    name: 'Contractor Dumpster Rental Agreement',
    docId: '11wK3Im5hdL-tjmDvQ7X0p9Sg7WcUhktHzHfzUTTBdLU',
  },
};

const TZ = 'America/Los_Angeles';
const FOLDER_NAME = 'TGD Signed Contracts (Onsite)';
const LOG_HEADERS = [
  'Record ID', 'Signed At', 'Contract', 'Printed Name', 'Company', 'Service Address',
  'Job #', 'Dumpster Size', 'Customer Email', 'Customer Phone', 'Driver',
  'GPS', 'PDF', 'Contract Version', 'Emailed Copy', 'Slack', 'Device',
];

// ---------- Web endpoints ----------

function doGet(e) {
  try {
    return json_(getContracts_());
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request' });
  }

  const props = PropertiesService.getScriptProperties();
  const pin = String(props.getProperty('DRIVER_PIN') || '');
  if (!pin || String(p.pin || '') !== pin) return json_({ ok: false, error: 'Wrong driver PIN', code: 'PIN' });
  if (p.action === 'checkPin') return json_({ ok: true });

  const c = CONTRACTS[p.contractKey];
  if (!c) return json_({ ok: false, error: 'Unknown contract' });
  const missing = ['id', 'printedName', 'signature', 'serviceAddress', 'driver'].filter(k => !String(p[k] || '').trim());
  if (missing.length) return json_({ ok: false, error: 'Missing: ' + missing.join(', ') });
  if (!(p.consent && p.consent.readAgree && p.consent.esign)) return json_({ ok: false, error: 'Consent boxes not checked' });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const log = logSheet_();
    // Phone retries after a timeout must not create a second contract.
    const dup = log.createTextFinder(p.id).matchEntireCell(true).findNext();
    if (dup) return json_({ ok: true, duplicate: true, pdfUrl: log.getRange(dup.getRow(), 13).getValue() });

    const signedAt = new Date();
    const pdf = buildPdf_(p, c, signedAt);
    const emailed = emailCustomer_(p, c, pdf);
    const slack = postSlack_(p, c, pdf, signedAt);

    log.appendRow([
      p.id, signedAt, c.name, p.printedName, p.company || '', p.serviceAddress,
      p.jobNumber || '', p.dumpsterSize || '', p.customerEmail || '', p.customerPhone || '', p.driver,
      gpsText_(p.geo), pdf.getUrl(), p.contractVersion || '', emailed, slack, p.userAgent || '',
    ]);
    return json_({ ok: true, pdfUrl: pdf.getUrl(), emailed: emailed });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Contracts ----------

function getContracts_() {
  const cache = CacheService.getScriptCache();
  const contracts = Object.keys(CONTRACTS).map(key => {
    const c = CONTRACTS[key];
    const version = Utilities.formatDate(DriveApp.getFileById(c.docId).getLastUpdated(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
    const cacheKey = 'html_' + c.docId + '_' + version;
    let html = cache.get(cacheKey);
    if (!html) {
      html = exportHtml_(c.docId);
      try { cache.put(cacheKey, html, 21600); } catch (e) { /* >100KB, skip cache */ }
    }
    return { key: key, name: c.name, version: version, html: html };
  });
  return { ok: true, contracts: contracts };
}

function exportHtml_(docId) {
  const res = UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId + '/export?format=html', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('Could not load contract (' + res.getResponseCode() + ')');
  return res.getContentText();
}

// ---------- Signed PDF ----------

function buildPdf_(p, c, signedAt) {
  const folder = DriveApp.getFolderById(folderId_());
  const title = [
    Utilities.formatDate(signedAt, TZ, 'yyyy-MM-dd'),
    c.name.replace(' Dumpster Rental Agreement', ''),
    p.printedName,
    p.jobNumber ? 'Job ' + p.jobNumber : '',
    p.serviceAddress,
  ].filter(Boolean).join(' - ').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 180);

  // Copy the live contract, append the signature page, export to PDF.
  const copy = DriveApp.getFileById(c.docId).makeCopy(title, folder);
  try {
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    body.appendPageBreak();
    body.appendParagraph('SIGNATURE AND ACCEPTANCE').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const details = [
      ['Lessee (printed name)', p.printedName],
      ['Company', p.company || '—'],
      ['Service address', p.serviceAddress],
      ['Job / order #', p.jobNumber || '—'],
      ['Dumpster size', p.dumpsterSize || '—'],
      ['Lessee email', p.customerEmail || '—'],
      ['Lessee phone', p.customerPhone || '—'],
      ['Date signed', Utilities.formatDate(signedAt, TZ, "MMMM d, yyyy 'at' h:mm a z")],
    ].map(r => r.map(String));
    body.appendTable(details);

    body.appendParagraph('Lessee signature:').editAsText().setBold(true);
    const sigBlob = Utilities.newBlob(
      Utilities.base64Decode(String(p.signature).split(',')[1]), 'image/png', 'signature.png');
    const img = body.appendParagraph('').appendInlineImage(sigBlob);
    const w = 300;
    img.setHeight(Math.round(img.getHeight() * w / img.getWidth())).setWidth(w);
    body.appendParagraph('X  ' + p.printedName);

    body.appendParagraph('Electronic signature record').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const audit = [
      ['Record ID', p.id],
      ['Signed on', 'Driver phone, in person, at service location'],
      ['Driver', p.driver],
      ['Device time', p.signedAtClient || '—'],
      ['Server time', Utilities.formatDate(signedAt, TZ, "yyyy-MM-dd HH:mm:ss z")],
      ['GPS', gpsText_(p.geo)],
      ['Contract version', c.name + ' — last edited ' + (p.contractVersion || '—')],
      ['Scrolled to end of agreement', p.viewedToEnd ? 'Yes' : 'No'],
      ['Consent', 'Checked: "I have read and agree to this Agreement, including Sections 2.5, 2.8, 7, 8 and 10." '
        + 'Checked: "I agree to sign electronically and receive my copy electronically."'],
      ['Device', p.userAgent || '—'],
    ].map(r => r.map(String));
    const t = body.appendTable(audit);
    t.editAsText().setFontSize(8);

    doc.saveAndClose();
    const pdf = folder.createFile(copy.getAs('application/pdf').setName(title + '.pdf'));
    try { pdf.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* folder sharing still applies */ }
    return pdf;
  } finally {
    copy.setTrashed(true); // temp working copy only; the PDF is the record
  }
}

// ---------- Customer copy ----------

function emailCustomer_(p, c, pdf) {
  const to = String(p.customerEmail || '').trim();
  if (!to || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) return 'no email';
  try {
    MailApp.sendEmail({
      to: to,
      subject: 'Your signed ' + c.name + ' — The Green Dumpster',
      name: 'The Green Dumpster',
      replyTo: 'office@thegreendumpster.com',
      htmlBody:
        '<p>Hi ' + esc_(p.printedName) + ',</p>' +
        '<p>Thanks for renting with The Green Dumpster. Your signed copy of the ' + esc_(c.name) +
        ' for <b>' + esc_(p.serviceAddress) + '</b> is attached.</p>' +
        '<p>Questions? Reply to this email or contact office@thegreendumpster.com.</p>' +
        '<p>— The Green Dumpster</p>',
      attachments: [pdf.getBlob()],
    });
    return 'sent';
  } catch (err) {
    console.error(err);
    return 'failed: ' + err.message;
  }
}

// ---------- Slack ----------

function postSlack_(p, c, pdf, signedAt) {
  const url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!url) return 'no webhook';
  const s = v => String(v || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const lines = [
    ':pencil: *' + s(c.name) + '* signed',
    '*Customer:* ' + s(p.printedName) + (p.company ? ' (' + s(p.company) + ')' : ''),
    '*Address:* ' + s(p.serviceAddress),
    '*Job #:* ' + (s(p.jobNumber) || '—') + '   *Size:* ' + (s(p.dumpsterSize) || '—'),
    '*Driver:* ' + s(p.driver) + '   *Signed:* ' + Utilities.formatDate(signedAt, TZ, 'MMM d, h:mm a'),
    '*Customer copy:* ' + (p.customerEmail ? s(p.customerEmail) : 'no email given'),
    '<' + pdf.getUrl() + '|View signed PDF>' +
      (p.geo ? '  ·  <https://maps.google.com/?q=' + p.geo.lat + ',' + p.geo.lng + '|Where it was signed>' : ''),
  ];
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: lines.join('\n'), unfurl_links: false }),
      muteHttpExceptions: true,
    });
    return res.getResponseCode() === 200 ? 'posted' : 'failed ' + res.getResponseCode();
  } catch (err) {
    return 'failed: ' + err.message;
  }
}

// ---------- Storage ----------

function folderId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('FOLDER_ID');
  if (id) return id;
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  id = (it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME)).getId();
  props.setProperty('FOLDER_ID', id);
  return id;
}

function logSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LOG_SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('TGD Signed Contracts Log');
    DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId_()));
    const sh = ss.getSheets()[0].setName('Signed');
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
    id = ss.getId();
    props.setProperty('LOG_SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id).getSheetByName('Signed');
}

// ---------- Helpers ----------

function gpsText_(g) {
  return g && g.lat ? g.lat.toFixed(6) + ', ' + g.lng.toFixed(6) + ' (±' + Math.round(g.acc || 0) + 'm)' : 'not available';
}

function esc_(v) {
  return String(v || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Run once from the editor ----------

/** Creates the Drive folder + log sheet, checks both contracts load, and approves all permissions. */
function setup() {
  const folder = DriveApp.getFolderById(folderId_());
  const sheet = logSheet_();
  const res = getContracts_();
  console.log('Folder: ' + folder.getUrl());
  console.log('Log sheet: ' + sheet.getParent().getUrl());
  res.contracts.forEach(c => console.log(c.name + ' — version ' + c.version + ', ' + c.html.length + ' chars'));
  const props = PropertiesService.getScriptProperties();
  ['SLACK_WEBHOOK_URL', 'DRIVER_PIN'].forEach(k => {
    if (!props.getProperty(k)) console.warn('Script property ' + k + ' is not set yet');
  });
  MailApp.getRemainingDailyQuota();
}

/** Posts a test message to Slack so you can confirm the webhook. */
function testSlack() {
  const url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ text: ':white_check_mark: Contract signing app is connected to this channel.' }),
    muteHttpExceptions: true,
  });
  console.log(res.getResponseCode() + ' ' + res.getContentText());
}

