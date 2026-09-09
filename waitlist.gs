/**
 * YAPP waitlist — Google Apps Script
 *
 * Manual deploy:
 * 1. Create a Google Sheet (e.g. "YAPP Waitlist")
 * 2. Extensions → Apps Script → paste this file's contents → Deploy → Web app
 *    (execute as me, access: Anyone)
 * 3. Copy the deployed web app URL into the WAITLIST_SCRIPT_URL constant
 *    at the top of index.html
 * 4. Optional: Sheet → Tools → Notification rules, as a backup alert channel
 *
 * MailApp daily quota is ~100 messages/day on consumer Gmail
 * (higher on Workspace). Fine for early waitlist volume.
 */

var SHEET_NAME = 'Waitlist';
var BURST_LIMIT = 20;
var BURST_WINDOW_SEC = 60;
var EMAIL_COOLDOWN_SEC = 60;

function doPost(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var honeypot = String(params.honeypot || '').trim();
  var email = String(params.email || '').trim();
  var phone = String(params.phone || '').trim();
  var source = String(params.source || '').trim();

  if (honeypot) {
    return json_(true);
  }

  if (!isValidEmail_(email) || !isValidPhone_(phone)) {
    return json_(false);
  }

  if (hitBurstLimit_()) {
    return json_(false);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_(false);
  }

  try {
    var sheet = getWaitlistSheet_();
    var normalized = email.toLowerCase();

    if (emailExists_(sheet, normalized)) {
      return json_(true);
    }

    if (emailOnCooldown_(normalized)) {
      return json_(false);
    }

    sheet.appendRow([new Date(), email, phone, source]);
    setEmailCooldown_(normalized);

    try {
      sendOwnerNotification_(email, phone, source);
      sendVisitorConfirmation_(email);
    } catch (mailErr) {
      // Row is already stored; do not fail the signup if MailApp quota/send errors.
    }

    return json_(true);
  } catch (err) {
    return json_(false);
  } finally {
    lock.releaseLock();
  }
}

function json_(ok) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getWaitlistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['timestamp', 'email', 'phone', 'source']);
  }
  return sheet;
}

function emailExists_(sheet, normalizedEmail) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var values = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === normalizedEmail) {
      return true;
    }
  }
  return false;
}

function hitBurstLimit_() {
  var cache = CacheService.getScriptCache();
  var now = Date.now();
  var data = { t: now, n: 0 };
  try {
    var raw = cache.get('burst');
    if (raw) data = JSON.parse(raw);
  } catch (err) {
    data = { t: now, n: 0 };
  }
  if (now - data.t > BURST_WINDOW_SEC * 1000) {
    data = { t: now, n: 0 };
  }
  data.n += 1;
  cache.put('burst', JSON.stringify(data), BURST_WINDOW_SEC);
  return data.n > BURST_LIMIT;
}

function emailOnCooldown_(normalizedEmail) {
  var cache = CacheService.getScriptCache();
  return !!cache.get(emailCooldownKey_(normalizedEmail));
}

function setEmailCooldown_(normalizedEmail) {
  CacheService.getScriptCache().put(
    emailCooldownKey_(normalizedEmail),
    '1',
    EMAIL_COOLDOWN_SEC
  );
}

function emailCooldownKey_(normalizedEmail) {
  return 'email_' + normalizedEmail;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidPhone_(phone) {
  var digits = phone.replace(/\D/g, '');
  return digits.length >= 7 && digits.length <= 15;
}

function sendOwnerNotification_(email, phone, source) {
  var owner = Session.getEffectiveUser().getEmail();
  if (!owner) return;
  MailApp.sendEmail({
    to: owner,
    subject: 'New YAPP waitlist signup',
    body: 'Email: ' + email + '\nPhone: ' + phone + '\nSource: ' + source
  });
}

function sendVisitorConfirmation_(email) {
  MailApp.sendEmail({
    to: email,
    subject: "You're on the YAPP waitlist",
    body: "You're on the list — we'll email you at launch."
  });
}
