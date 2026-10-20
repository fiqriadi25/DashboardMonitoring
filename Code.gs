const SPREADSHEET_ID = "1d5eFnUM65lzpaJUJXSdpZG0BmK70SCf3TF5K4tljCZw";
const SHEET_NAME = "Database";

const EMERGENCY_ALLOWED_EMAILS = [
  'admin.bitrans@b-log.co.id'
];

const ACCESS_ALLOWLIST_SEED = [
  { email: 'admin.bitrans@b-log.co.id', name: 'Admin Bitrans', note: 'Akses admin darurat' },
  { email: 'arsyal.aswindi1996@gmail.com', name: '', note: 'Migrasi allowedEmails lama' },
  { email: 'hherdi123@gmail.com', name: '', note: 'Migrasi allowedEmails lama' },
  { email: 'ginginanjar683@gmail.com', name: '', note: 'Migrasi allowedEmails lama' }
];

const allowedDomains = ['b-log.co.id', 'b-log.id'];
const ACCESS_ADMIN_EMAIL = 'admin.bitrans@b-log.co.id';
const ACCESS_REQUEST_SHEET = 'Access_Request';
const ACCESS_ALLOWLIST_SHEET = 'Access_Allowlist';
const ACCESS_REGISTRY_CACHE_KEY = 'ACCESS_REGISTRY_V2';
const ACCESS_REGISTRY_CACHE_SECONDS = 60;
const ACCESS_REQUEST_COOLDOWN_SECONDS = 600;
const ACCESS_REQUEST_DAILY_LIMIT = 5;
const ACCESS_VERIFICATION_TTL_SECONDS = 600;
const ACCESS_VERIFICATION_SEND_COOLDOWN_SECONDS = 60;
const ACCESS_VERIFICATION_MAX_ATTEMPTS = 5;
const ACCESS_VERIFICATION_SECRET_PROPERTY = 'ACCESS_VERIFICATION_SECRET';
const ACCESS_ALLOWLIST_HEADERS = ['Email', 'Nama', 'Status', 'Keterangan', 'Updated_At'];
const ACCESS_REQUEST_HEADERS = [
  'Timestamp',
  'Request_ID',
  'User_Key_Hash',
  'Nama',
  'Email',
  'No_WhatsApp',
  'Alasan_Akses',
  'Status',
  'Reviewed_By',
  'Reviewed_At',
  'Last_Update'
];

/* ================= WEB APP & ACCESS CONTROL ================== */

function doGet() {
  try {
    ensureAccessAllowlistSheet_();
  } catch (error) {
    console.warn('Access_Allowlist belum dapat disiapkan:', error);
  }
  const access = getAccessContext_();

  if (!access.allowed) {
    return HtmlService.createHtmlOutput(buildAccessRequestPage_(access))
      .setTitle('Permintaan Akses - Dashboard Monitoring Rekon')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dashboard Digital Enablement')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeWhatsApp_(value) {
  let digits = String(value || '').replace(/\D/g, '');
  if (digits.indexOf('0') === 0) digits = '62' + digits.substring(1);
  else if (digits.indexOf('8') === 0) digits = '62' + digits;
  return digits;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getActiveEmail_() {
  try {
    return normalizeEmail_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function getActiveUserKeyHash_() {
  try {
    const key = String(Session.getTemporaryActiveUserKey() || '').trim();
    return key ? cacheDigest_(key) : '';
  } catch (error) {
    return '';
  }
}

function getEmailDomain_(email) {
  const normalized = normalizeEmail_(email);
  const atIndex = normalized.lastIndexOf('@');
  return atIndex > -1 ? normalized.substring(atIndex + 1) : '';
}

function isInternalEmail_(email) {
  const domain = getEmailDomain_(email);
  return allowedDomains.some(function(item) {
    return String(item || '').trim().toLowerCase() === domain;
  });
}

function clearAccessRegistryCache_() {
  try {
    CacheService.getScriptCache().remove(ACCESS_REGISTRY_CACHE_KEY);
  } catch (error) {
    console.warn('Gagal menghapus cache akses:', error);
  }
}

function ensureAccessAllowlistSheet_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(ACCESS_ALLOWLIST_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(ACCESS_ALLOWLIST_SHEET);
  const header = sheet.getRange(1, 1, 1, ACCESS_ALLOWLIST_HEADERS.length);
  header.setValues([ACCESS_ALLOWLIST_HEADERS]);
  header
    .setBackground('#006b49')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 280);
  sheet.setColumnWidth(5, 160);

  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ACTIVE', 'INACTIVE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(statusValidation);

  if (ACCESS_ALLOWLIST_SEED.length) {
    const now = new Date();
    const rows = ACCESS_ALLOWLIST_SEED.map(function(item) {
      return [
        normalizeEmail_(item.email),
        String(item.name || '').trim(),
        'ACTIVE',
        String(item.note || '').trim(),
        now
      ];
    });
    sheet.getRange(2, 1, rows.length, ACCESS_ALLOWLIST_HEADERS.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  return sheet;
}

function getAccessRegistryEmails_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    try {
      const cached = cache.get(ACCESS_REGISTRY_CACHE_KEY);
      if (cached) return JSON.parse(cached);
    } catch (error) {
      console.warn('Cache akses tidak dapat dibaca:', error);
    }
  }

  const emailStates = {};
  const approvedRequests = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allowlistSheet = ensureAccessAllowlistSheet_(ss);

  if (allowlistSheet.getLastRow() >= 2) {
    const rows = allowlistSheet
      .getRange(2, 1, allowlistSheet.getLastRow() - 1, ACCESS_ALLOWLIST_HEADERS.length)
      .getValues();

    rows.forEach(function(row) {
      const email = normalizeEmail_(row[0]);
      const status = String(row[2] || '').trim().toUpperCase();
      if (!email) return;
      if (status === 'ACTIVE' || status === 'INACTIVE') emailStates[email] = status;
    });
  }

  const requestSheet = getAccessRequestSheet_(ss);
  if (requestSheet && requestSheet.getLastRow() >= 2) {
    const rows = requestSheet.getRange(2, 5, requestSheet.getLastRow() - 1, 4).getValues();
    rows.forEach(function(row) {
      const email = normalizeEmail_(row[0]);
      const status = String(row[3] || '').trim().toUpperCase();
      if (email && status === 'APPROVED') approvedRequests[email] = true;
    });
  }

  const allowed = {};
  Object.keys(approvedRequests).forEach(function(email) {
    if (emailStates[email] !== 'INACTIVE') allowed[email] = true;
  });
  Object.keys(emailStates).forEach(function(email) {
    if (emailStates[email] === 'ACTIVE') allowed[email] = true;
    if (emailStates[email] === 'INACTIVE') delete allowed[email];
  });
  EMERGENCY_ALLOWED_EMAILS.forEach(function(email) {
    const normalized = normalizeEmail_(email);
    if (normalized) allowed[normalized] = true;
  });

  const result = Object.keys(allowed).sort();
  try {
    cache.put(ACCESS_REGISTRY_CACHE_KEY, JSON.stringify(result), ACCESS_REGISTRY_CACHE_SECONDS);
  } catch (error) {
    console.warn('Cache akses tidak dapat disimpan:', error);
  }
  return result;
}

function isExplicitEmailAllowed_(email, forceRefresh) {
  const normalized = normalizeEmail_(email);
  return Boolean(normalized) && getAccessRegistryEmails_(Boolean(forceRefresh)).indexOf(normalized) > -1;
}

function getAccessRequestSheet_(spreadsheet) {
  try {
    return (spreadsheet || SpreadsheetApp.openById(SPREADSHEET_ID)).getSheetByName(ACCESS_REQUEST_SHEET);
  } catch (error) {
    return null;
  }
}

function getLatestAccessRequestByKey_(userKeyHash, forceRefresh) {
  if (!userKeyHash) return null;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ACCESS_REQUEST_LATEST:' + cacheDigest_(userKeyHash);
  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        return parsed.found ? parsed.request : null;
      } catch (error) {
        cache.remove(cacheKey);
      }
    }
  }
  const sheet = getAccessRequestSheet_();
  if (!sheet || sheet.getLastRow() < 2) {
    cache.put(cacheKey, JSON.stringify({ found: false }), ACCESS_REGISTRY_CACHE_SECONDS);
    return null;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, ACCESS_REQUEST_HEADERS.length).getValues();
  for (let index = values.length - 1; index >= 0; index--) {
    const row = values[index];
    if (String(row[2] || '') === userKeyHash) {
      const request = {
        rowNumber: index + 2,
        requestId: String(row[1] || ''),
        email: normalizeEmail_(row[4]),
        status: String(row[7] || 'PENDING').trim().toUpperCase()
      };
      cache.put(cacheKey, JSON.stringify({ found: true, request: request }), ACCESS_REGISTRY_CACHE_SECONDS);
      return request;
    }
  }
  cache.put(cacheKey, JSON.stringify({ found: false }), ACCESS_REGISTRY_CACHE_SECONDS);
  return null;
}

function clearLatestAccessRequestCache_(userKeyHash) {
  if (!userKeyHash) return;
  CacheService.getScriptCache().remove('ACCESS_REQUEST_LATEST:' + cacheDigest_(userKeyHash));
}

function countRecentAccessRequests_(userKeyHash, email, sinceDate) {
  const sheet = getAccessRequestSheet_();
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const normalizedEmail = normalizeEmail_(email);
  const since = sinceDate instanceof Date ? sinceDate.getTime() : 0;
  return values.reduce(function(count, row) {
    const timestamp = parseDate_(row[0]);
    if (!timestamp || timestamp.getTime() < since) return count;
    const sameUser = userKeyHash && String(row[2] || '') === userKeyHash;
    const sameEmail = normalizedEmail && normalizeEmail_(row[4]) === normalizedEmail;
    return count + (sameUser || sameEmail ? 1 : 0);
  }, 0);
}

function appendAccessRequestRow_(request, userKeyHash, status, review) {
  const sheet = ensureAccessRequestSheet_();
  const now = new Date();
  const requestId = 'ACC-' + Utilities.formatDate(
    now,
    Session.getScriptTimeZone() || 'Asia/Jakarta',
    'yyyyMMdd-HHmmss'
  ) + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const reviewData = review || {};

  sheet.appendRow([
    now,
    requestId,
    userKeyHash,
    request.name,
    request.email,
    request.whatsapp,
    request.reason,
    status,
    reviewData.reviewedBy || '',
    reviewData.reviewedAt || '',
    now
  ]);
  const rowNumber = sheet.getLastRow();
  sheet.getRange(rowNumber, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sheet.getRange(rowNumber, 7).setWrap(true);
  if (reviewData.reviewedAt) {
    sheet.getRange(rowNumber, 10, 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  clearLatestAccessRequestCache_(userKeyHash);
  return { sheet: sheet, rowNumber: rowNumber, requestId: requestId };
}

function getAccessVerificationSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(ACCESS_VERIFICATION_SECRET_PROPERTY);
  if (secret) return secret;

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    secret = props.getProperty(ACCESS_VERIFICATION_SECRET_PROPERTY);
    if (!secret) {
      secret = Utilities.getUuid() + Utilities.getUuid();
      props.setProperty(ACCESS_VERIFICATION_SECRET_PROPERTY, secret);
    }
    return secret;
  } finally {
    lock.releaseLock();
  }
}

function accessVerificationCacheKey_(userKeyHash, email) {
  return 'ACCESS_VERIFY:' + cacheDigest_(userKeyHash + '|' + normalizeEmail_(email));
}

function hashAccessVerificationCode_(userKeyHash, email, code) {
  return cacheDigest_([
    getAccessVerificationSecret_(),
    userKeyHash,
    normalizeEmail_(email),
    String(code || '').trim()
  ].join('|'));
}

function generateAccessVerificationCode_() {
  const hex = Utilities.getUuid().replace(/-/g, '').substring(0, 12);
  return String(parseInt(hex, 16) % 1000000).padStart(6, '0');
}

function sendApprovedEmailVerification_(request, userKeyHash) {
  if (!(isInternalEmail_(request.email) || isExplicitEmailAllowed_(request.email, true))) {
    return null;
  }

  const cache = CacheService.getScriptCache();
  const sendKey = 'ACCESS_VERIFY_SEND:' + cacheDigest_(userKeyHash + '|' + request.email);
  const emailSendKey = 'ACCESS_VERIFY_EMAIL:' + cacheDigest_(request.email);
  if (cache.get(sendKey) || cache.get(emailSendKey)) {
    throw new Error('Kode verifikasi baru saja dikirim. Tunggu satu menit sebelum meminta kode baru.');
  }
  if (MailApp.getRemainingDailyQuota() < 1) {
    throw new Error('Kuota email verifikasi hari ini sudah habis. Silakan hubungi admin.');
  }

  const code = generateAccessVerificationCode_();
  const verificationKey = accessVerificationCacheKey_(userKeyHash, request.email);
  cache.put(verificationKey, JSON.stringify({
    codeHash: hashAccessVerificationCode_(userKeyHash, request.email, code),
    attempts: 0,
    name: request.name,
    email: request.email,
    whatsapp: request.whatsapp
  }), ACCESS_VERIFICATION_TTL_SECONDS);
  cache.put(sendKey, '1', ACCESS_VERIFICATION_SEND_COOLDOWN_SECONDS);
  cache.put(emailSendKey, '1', ACCESS_VERIFICATION_SEND_COOLDOWN_SECONDS);

  MailApp.sendEmail({
    to: request.email,
    name: 'Dashboard Monitoring Rekon',
    subject: '[Verifikasi Akses Dashboard] Kode ' + code,
    body: [
      'Kode verifikasi akses Dashboard Monitoring Rekon Anda adalah: ' + code,
      '',
      'Kode berlaku selama 10 menit dan hanya boleh digunakan pada sesi browser yang meminta kode.',
      'Jika Anda tidak meminta kode ini, abaikan email ini.'
    ].join('\n'),
    htmlBody:
      '<div style="font-family:Arial,sans-serif;color:#1f2937;max-width:560px">' +
        '<h2 style="color:#006b49">Verifikasi Akses Dashboard</h2>' +
        '<p>Gunakan kode berikut untuk menautkan kembali sesi browser Anda:</p>' +
        '<div style="font-size:30px;font-weight:800;letter-spacing:8px;padding:16px 20px;background:#eef5f2;border-radius:14px;text-align:center">' + escapeHtml_(code) + '</div>' +
        '<p style="color:#64748b">Kode berlaku selama 10 menit. Abaikan email ini jika Anda tidak memintanya.</p>' +
      '</div>'
  });

  return {
    success: true,
    authorized: false,
    verificationRequired: true,
    email: request.email,
    message: 'Email ini sudah disetujui. Masukkan kode 6 digit yang dikirim ke email tersebut.'
  };
}

function getAccessContext_(forceRefresh) {
  const activeEmail = getActiveEmail_();
  const userKeyHash = getActiveUserKeyHash_();

  if (activeEmail && isInternalEmail_(activeEmail)) {
    return { allowed: true, email: activeEmail, accessType: 'internal', userKeyHash: userKeyHash };
  }

  if (activeEmail && isExplicitEmailAllowed_(activeEmail, forceRefresh)) {
    return { allowed: true, email: activeEmail, accessType: 'external-detected', userKeyHash: userKeyHash };
  }

  const request = getLatestAccessRequestByKey_(userKeyHash, Boolean(forceRefresh));
  if (
    !activeEmail &&
    request &&
    request.status === 'APPROVED' &&
    (isInternalEmail_(request.email) || isExplicitEmailAllowed_(request.email, forceRefresh))
  ) {
    return {
      allowed: true,
      email: request.email,
      accessType: 'external-registered',
      userKeyHash: userKeyHash,
      request: request
    };
  }

  return {
    allowed: false,
    email: activeEmail,
    accessType: activeEmail ? 'external-unregistered' : 'identity-hidden',
    userKeyHash: userKeyHash,
    request: request
  };
}

function assertAuthorized_() {
  const access = getAccessContext_();
  if (!access.allowed) {
    throw new Error('ACCESS_DENIED: Email belum terdaftar atau sesi pengguna tidak dapat diverifikasi. Muat ulang halaman untuk mengajukan akses.');
  }
  return access;
}

function isTrustedTriggerEvent_(event) {
  if (!event || !event.triggerUid) return false;
  const triggerUid = String(event.triggerUid);
  try {
    return ScriptApp.getProjectTriggers().some(function(trigger) {
      return String(trigger.getUniqueId()) === triggerUid;
    });
  } catch (error) {
    return false;
  }
}

function assertAuthorizedOrTrustedTrigger_(event) {
  if (isTrustedTriggerEvent_(event)) {
    return { allowed: true, accessType: 'trusted-trigger' };
  }
  return assertAdmin_();
}

function ensureAccessRequestSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(ACCESS_REQUEST_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(ACCESS_REQUEST_SHEET);
  const header = sheet.getRange(1, 1, 1, ACCESS_REQUEST_HEADERS.length);
  header.setValues([ACCESS_REQUEST_HEADERS]);
  header
    .setBackground('#006b49')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 145);
  sheet.setColumnWidth(7, 360);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 190);
  sheet.setColumnWidth(10, 155);
  sheet.setColumnWidth(11, 155);
  sheet.hideColumns(3);

  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PENDING', 'APPROVED', 'REJECTED', 'EMAIL_FAILED'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 8, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(statusValidation);
  return sheet;
}

function validateAccessRequest_(payload) {
  const source = payload && typeof payload === 'object' ? payload : {};
  const name = String(source.name || '').replace(/\s+/g, ' ').trim();
  const email = normalizeEmail_(source.email);
  const whatsapp = normalizeWhatsApp_(source.whatsapp);
  const reason = String(source.reason || '').replace(/\s+/g, ' ').trim();

  if (name.length < 3 || name.length > 100) {
    throw new Error('Nama lengkap harus terdiri dari 3-100 karakter.');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 150) {
    throw new Error('Format email tidak valid.');
  }
  if (!/^62\d{8,13}$/.test(whatsapp)) {
    throw new Error('Nomor WhatsApp tidak valid. Gunakan nomor Indonesia yang aktif.');
  }
  if (reason.length < 10 || reason.length > 500) {
    throw new Error('Alasan akses harus terdiri dari 10-500 karakter.');
  }

  return { name: name, email: email, whatsapp: whatsapp, reason: reason };
}

function submitAccessRequest(payload) {
  const request = validateAccessRequest_(payload);
  const activeEmail = getActiveEmail_();
  const userKeyHash = getActiveUserKeyHash_();

  if (!userKeyHash) {
    throw new Error('Sesi Google tidak dapat dikenali. Pastikan browser mengizinkan cookie, lalu muat ulang halaman.');
  }
  if (activeEmail && activeEmail !== request.email) {
    throw new Error('Email harus sama dengan akun Google yang sedang aktif: ' + activeEmail);
  }
  if (activeEmail && (isInternalEmail_(activeEmail) || isExplicitEmailAllowed_(activeEmail, true))) {
    return { success: true, authorized: true, message: 'Akses Anda sudah aktif. Dashboard akan dimuat ulang.' };
  }

  if (!activeEmail) {
    const verificationResult = sendApprovedEmailVerification_(request, userKeyHash);
    if (verificationResult) return verificationResult;
  }

  const cache = CacheService.getScriptCache();
  const cooldownKey = 'ACCESS_REQUEST:' + userKeyHash;
  if (cache.get(cooldownKey)) {
    throw new Error('Permintaan sudah dikirim. Silakan tunggu 10 menit sebelum mengirim ulang.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Sistem sedang memproses permintaan lain. Silakan coba kembali beberapa saat lagi.');
  }

  let created;

  try {
    const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
    if (countRecentAccessRequests_(userKeyHash, request.email, since) >= ACCESS_REQUEST_DAILY_LIMIT) {
      throw new Error('Batas permintaan akses harian tercapai. Silakan hubungi admin.');
    }
    cache.put(cooldownKey, '1', ACCESS_REQUEST_COOLDOWN_SECONDS);
    created = appendAccessRequestRow_(request, userKeyHash, 'PENDING');
  } finally {
    lock.releaseLock();
  }

  const sheet = created.sheet;
  const rowNumber = created.rowNumber;
  const requestId = created.requestId;

  try {
    if (MailApp.getRemainingDailyQuota() < 1) {
      throw new Error('Kuota pengiriman email admin hari ini sudah habis.');
    }

    const adminWhatsAppLink = 'https://wa.me/' + request.whatsapp + '?text=' + encodeURIComponent(
      'Halo ' + request.name + ', terkait permintaan akses Dashboard Monitoring Rekon dengan ID ' + requestId + '.'
    );
    const subject = '[Permintaan Akses Dashboard] ' + request.name + ' - ' + request.email;
    const body = [
      'Permintaan akses baru telah diterima.',
      '',
      'Request ID: ' + requestId,
      'Nama: ' + request.name,
      'Email: ' + request.email,
      'No. WhatsApp: +' + request.whatsapp,
      'Alasan akses: ' + request.reason,
      'Baris sheet: ' + ACCESS_REQUEST_SHEET + '!' + rowNumber,
      '',
      'Setelah identitas pengguna diverifikasi:',
      '1. Verifikasi identitas pengguna melalui data yang dikirim.',
      '2. Ubah Status pada baris tersebut menjadi APPROVED.',
      '3. Akses akan aktif tanpa mengubah kode atau membuat deployment baru.',
      '',
      'Hubungi pendaftar: ' + adminWhatsAppLink
    ].join('\n');
    const htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:620px;color:#1f2937">' +
        '<h2 style="color:#006b49">Permintaan Akses Dashboard</h2>' +
        '<table style="border-collapse:collapse;width:100%">' +
          '<tr><td style="padding:8px;border:1px solid #d1d5db"><b>Request ID</b></td><td style="padding:8px;border:1px solid #d1d5db">' + escapeHtml_(requestId) + '</td></tr>' +
          '<tr><td style="padding:8px;border:1px solid #d1d5db"><b>Nama</b></td><td style="padding:8px;border:1px solid #d1d5db">' + escapeHtml_(request.name) + '</td></tr>' +
          '<tr><td style="padding:8px;border:1px solid #d1d5db"><b>Email</b></td><td style="padding:8px;border:1px solid #d1d5db">' + escapeHtml_(request.email) + '</td></tr>' +
          '<tr><td style="padding:8px;border:1px solid #d1d5db"><b>WhatsApp</b></td><td style="padding:8px;border:1px solid #d1d5db">+' + escapeHtml_(request.whatsapp) + '</td></tr>' +
          '<tr><td style="padding:8px;border:1px solid #d1d5db"><b>Alasan akses</b></td><td style="padding:8px;border:1px solid #d1d5db">' + escapeHtml_(request.reason) + '</td></tr>' +
        '</table>' +
        '<p style="margin-top:18px">Verifikasi pengguna, lalu ubah Status pada <b>' + escapeHtml_(ACCESS_REQUEST_SHEET + '!H' + rowNumber) + '</b> menjadi <b>APPROVED</b>. Akses akan aktif tanpa deployment ulang.</p>' +
        '<a href="' + adminWhatsAppLink + '" style="display:inline-block;background:#25d366;color:#fff;text-decoration:none;padding:10px 18px;border-radius:999px;font-weight:bold">Hubungi via WhatsApp</a>' +
      '</div>';

    MailApp.sendEmail({
      to: ACCESS_ADMIN_EMAIL,
      replyTo: request.email,
      name: 'Dashboard Monitoring Rekon',
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
  } catch (mailError) {
    if (sheet && rowNumber) {
      sheet.getRange(rowNumber, 8).setValue('EMAIL_FAILED');
      sheet.getRange(rowNumber, 11).setValue(new Date());
    }
    throw new Error('Permintaan tercatat, tetapi email admin gagal dikirim: ' + mailError.message);
  }

  return {
    success: true,
    authorized: false,
    requestId: requestId,
    message: 'Permintaan akses berhasil dikirim ke ' + ACCESS_ADMIN_EMAIL + '. Silakan tunggu verifikasi admin.'
  };
}

function verifyApprovedEmailAccess(payload) {
  const source = payload && typeof payload === 'object' ? payload : {};
  const email = normalizeEmail_(source.email);
  const code = String(source.code || '').replace(/\D/g, '');
  const userKeyHash = getActiveUserKeyHash_();

  if (!userKeyHash) throw new Error('Sesi Google tidak dapat dikenali. Muat ulang halaman dan coba kembali.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email verifikasi tidak valid.');
  if (!/^\d{6}$/.test(code)) throw new Error('Kode verifikasi harus terdiri dari 6 digit.');

  const cache = CacheService.getScriptCache();
  const verificationKey = accessVerificationCacheKey_(userKeyHash, email);
  const raw = cache.get(verificationKey);
  if (!raw) throw new Error('Kode verifikasi sudah kedaluwarsa. Minta kode baru.');

  const verification = JSON.parse(raw);
  verification.attempts = Number(verification.attempts || 0) + 1;
  if (verification.attempts > ACCESS_VERIFICATION_MAX_ATTEMPTS) {
    cache.remove(verificationKey);
    throw new Error('Terlalu banyak percobaan. Minta kode verifikasi baru.');
  }

  const expectedHash = hashAccessVerificationCode_(userKeyHash, email, code);
  if (expectedHash !== verification.codeHash) {
    cache.put(verificationKey, JSON.stringify(verification), ACCESS_VERIFICATION_TTL_SECONDS);
    throw new Error('Kode verifikasi tidak sesuai.');
  }
  if (!(isInternalEmail_(email) || isExplicitEmailAllowed_(email, true))) {
    cache.remove(verificationKey);
    throw new Error('Akses email ini sudah tidak aktif. Hubungi admin.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    appendAccessRequestRow_({
      name: String(verification.name || '').trim() || 'Pengguna terverifikasi',
      email: email,
      whatsapp: String(verification.whatsapp || ''),
      reason: 'Re-verifikasi kepemilikan email untuk sesi pengguna baru.'
    }, userKeyHash, 'APPROVED', {
      reviewedBy: 'EMAIL_OTP',
      reviewedAt: new Date()
    });
  } finally {
    lock.releaseLock();
  }

  cache.remove(verificationKey);
  clearAccessRegistryCache_();
  return {
    success: true,
    authorized: true,
    message: 'Email berhasil diverifikasi. Dashboard akan dimuat ulang.'
  };
}

function getAccessRequestStatus() {
  const access = getAccessContext_(true);
  if (access.allowed) {
    return { authorized: true, status: 'APPROVED', message: 'Akses sudah disetujui.' };
  }

  return {
    authorized: false,
    status: access.request ? access.request.status : 'NOT_SUBMITTED',
    requestId: access.request ? access.request.requestId : '',
    message: access.request
      ? 'Status permintaan: ' + access.request.status
      : 'Belum ada permintaan akses untuk sesi ini.'
  };
}

function buildAccessRequestPage_(access) {
  const request = access.request || null;
  const prefillEmail = normalizeEmail_(access.email || (request && request.email) || '');
  const emailReadOnly = access.email ? ' readonly' : '';
  const identityMessage = access.email
    ? 'Email Google terdeteksi, tetapi belum terdaftar pada sistem.'
    : 'Email Google tidak dapat dideteksi. Silakan ajukan permintaan akses.';
  const currentStatus = request ? request.status : '';

  return `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root { --green:#006b49; --green2:#00875f; --ink:#17332a; --muted:#64748b; --line:#dbe7e2; --bg:#eef5f2; --danger:#b91c1c; }
    * { box-sizing:border-box; }
    body { margin:0; min-height:100vh; display:grid; place-items:center; padding:24px; font-family:Arial,sans-serif; color:var(--ink); background:linear-gradient(145deg,#e8f3ef,#f8fbfa); }
    .shell { width:min(100%,560px); }
    .brand { display:flex; align-items:center; gap:12px; margin-bottom:16px; }
    .brand-mark { width:46px; height:46px; border-radius:14px; display:grid; place-items:center; color:#fff; background:linear-gradient(135deg,var(--green),var(--green2)); font-size:22px; font-weight:800; box-shadow:0 10px 24px rgba(0,107,73,.22); }
    .brand h1 { margin:0; font-size:21px; }
    .brand p { margin:3px 0 0; color:var(--muted); font-size:13px; }
    .card { background:#fff; border:1px solid var(--line); border-radius:22px; padding:26px; box-shadow:0 22px 55px rgba(15,60,45,.12); }
    .notice { padding:12px 14px; border-radius:13px; background:#eff8f4; border:1px solid #cce7db; color:#285f4c; font-size:13px; line-height:1.5; margin-bottom:20px; }
    .status { font-weight:700; }
    .grid { display:grid; gap:15px; }
    label { display:block; font-size:13px; font-weight:700; margin-bottom:7px; }
    input, textarea { width:100%; border:1px solid #cadbd4; border-radius:12px; padding:12px 13px; font:inherit; color:var(--ink); background:#fff; outline:none; transition:border-color .2s,box-shadow .2s; }
    input:focus, textarea:focus { border-color:var(--green2); box-shadow:0 0 0 3px rgba(0,135,95,.12); }
    input[readonly] { background:#f3f7f5; color:#526a61; }
    textarea { min-height:104px; resize:vertical; }
    .hint { margin-top:6px; color:var(--muted); font-size:12px; line-height:1.45; }
    .actions { display:grid; grid-template-columns:1fr auto; gap:10px; margin-top:20px; }
    button { border:0; border-radius:12px; padding:12px 16px; font-weight:700; cursor:pointer; transition:transform .15s,opacity .15s,background .15s; }
    button:hover { transform:translateY(-1px); }
    button:disabled { opacity:.58; cursor:not-allowed; transform:none; }
    .primary { color:#fff; background:linear-gradient(135deg,var(--green),var(--green2)); }
    .secondary { color:var(--green); background:#e8f4ef; }
    .verification { display:none; margin-top:18px; padding:16px; border:1px solid #cce7db; border-radius:14px; background:#f6fbf9; }
    .verification.active { display:block; }
    .verification-actions { display:flex; gap:10px; margin-top:12px; }
    .verification input { letter-spacing:7px; text-align:center; font-size:20px; font-weight:800; }
    .message { display:none; margin-top:15px; padding:12px 14px; border-radius:12px; font-size:13px; line-height:1.5; }
    .message.success { display:block; background:#eaf8f1; color:#16643f; border:1px solid #bfe4d0; }
    .message.error { display:block; background:#fff0f0; color:var(--danger); border:1px solid #f4c4c4; }
    .foot { color:var(--muted); text-align:center; font-size:11px; margin-top:14px; }
    @media (max-width:560px) { body{padding:15px}.card{padding:20px}.actions{grid-template-columns:1fr}.secondary{order:2} }
  </style>
</head>
<body>
  <main class="shell">
    <div class="brand">
      <div class="brand-mark">B</div>
      <div><h1>Dashboard Monitoring Rekon</h1><p>Permintaan akses pengguna</p></div>
    </div>
    <section class="card">
      <div class="notice">
        ${escapeHtml_(identityMessage)}
        ${currentStatus ? '<br>Status terakhir: <span class="status">' + escapeHtml_(currentStatus) + '</span>' : ''}
      </div>
      <form id="accessForm" novalidate>
        <div class="grid">
          <div><label for="name">Nama lengkap</label><input id="name" name="name" maxlength="100" autocomplete="name" required></div>
          <div><label for="email">Email</label><input id="email" name="email" type="email" maxlength="150" value="${escapeHtml_(prefillEmail)}"${emailReadOnly} autocomplete="email" required><div class="hint">Email akan dinormalisasi menjadi huruf kecil.</div></div>
          <div><label for="whatsapp">Nomor WhatsApp</label><input id="whatsapp" name="whatsapp" type="tel" maxlength="20" placeholder="Contoh: 081234567890" autocomplete="tel" required><div class="hint">Gunakan nomor WhatsApp Indonesia yang aktif.</div></div>
          <div><label for="reason">Alasan akses</label><textarea id="reason" name="reason" maxlength="500" placeholder="Jelaskan kebutuhan akses dashboard Anda" required></textarea></div>
        </div>
        <div class="actions">
          <button id="submitBtn" class="primary" type="submit">Kirim Permintaan Akses</button>
          <button id="statusBtn" class="secondary" type="button">Cek Status</button>
        </div>
      </form>
      <div id="verificationPanel" class="verification">
        <label for="verificationCode">Kode verifikasi email</label>
        <input id="verificationCode" type="text" inputmode="numeric" maxlength="6" pattern="[0-9]{6}" placeholder="000000" autocomplete="one-time-code">
        <div class="hint">Kode berlaku 10 menit dan hanya dapat digunakan pada sesi browser ini.</div>
        <div class="verification-actions">
          <button id="verifyBtn" class="primary" type="button">Verifikasi &amp; Buka Dashboard</button>
        </div>
      </div>
      <div id="message" class="message" role="status" aria-live="polite"></div>
    </section>
    <div class="foot">Permintaan dikirim ke ${escapeHtml_(ACCESS_ADMIN_EMAIL)} untuk diverifikasi.</div>
  </main>
  <script>
    (function () {
      var form = document.getElementById('accessForm');
      var submitBtn = document.getElementById('submitBtn');
      var statusBtn = document.getElementById('statusBtn');
      var verificationPanel = document.getElementById('verificationPanel');
      var verificationCode = document.getElementById('verificationCode');
      var verifyBtn = document.getElementById('verifyBtn');
      var message = document.getElementById('message');
      var verificationEmail = '';

      function showMessage(text, type) {
        message.textContent = text || '';
        message.className = 'message ' + (type || '');
      }

      function errorText(error) {
        return error && error.message ? error.message : String(error || 'Terjadi kesalahan.');
      }

      form.addEventListener('submit', function (event) {
        event.preventDefault();
        if (!form.reportValidity()) return;
        submitBtn.disabled = true;
        submitBtn.textContent = 'Mengirim...';
        showMessage('', '');

        google.script.run
          .withSuccessHandler(function (result) {
            submitBtn.disabled = false;
            submitBtn.textContent = 'Kirim Permintaan Akses';
            if (result && result.authorized) {
              showMessage(result.message, 'success');
              setTimeout(function () { window.top.location.reload(); }, 700);
              return;
            }
            if (result && result.verificationRequired) {
              verificationEmail = result.email || document.getElementById('email').value;
              verificationPanel.classList.add('active');
              verificationCode.value = '';
              verificationCode.focus();
            }
            showMessage((result && result.message ? result.message : 'Permintaan berhasil dikirim.') + (result && result.requestId ? ' ID: ' + result.requestId : ''), 'success');
          })
          .withFailureHandler(function (error) {
            submitBtn.disabled = false;
            submitBtn.textContent = 'Kirim Permintaan Akses';
            showMessage(errorText(error), 'error');
          })
          .submitAccessRequest({
            name: document.getElementById('name').value,
            email: document.getElementById('email').value,
            whatsapp: document.getElementById('whatsapp').value,
            reason: document.getElementById('reason').value
          });
      });

      verifyBtn.addEventListener('click', function () {
        if (!verificationCode.reportValidity()) return;
        verifyBtn.disabled = true;
        verifyBtn.textContent = 'Memverifikasi...';

        google.script.run
          .withSuccessHandler(function (result) {
            verifyBtn.disabled = false;
            verifyBtn.textContent = 'Verifikasi & Buka Dashboard';
            showMessage(result && result.message ? result.message : 'Email berhasil diverifikasi.', 'success');
            setTimeout(function () { window.top.location.reload(); }, 700);
          })
          .withFailureHandler(function (error) {
            verifyBtn.disabled = false;
            verifyBtn.textContent = 'Verifikasi & Buka Dashboard';
            showMessage(errorText(error), 'error');
          })
          .verifyApprovedEmailAccess({
            email: verificationEmail || document.getElementById('email').value,
            code: verificationCode.value
          });
      });

      statusBtn.addEventListener('click', function () {
        statusBtn.disabled = true;
        statusBtn.textContent = 'Memeriksa...';
        google.script.run
          .withSuccessHandler(function (result) {
            statusBtn.disabled = false;
            statusBtn.textContent = 'Cek Status';
            if (result && result.authorized) {
              showMessage('Akses sudah disetujui. Membuka dashboard...', 'success');
              setTimeout(function () { window.top.location.reload(); }, 500);
              return;
            }
            showMessage(result && result.message ? result.message : 'Akses belum disetujui.', 'error');
          })
          .withFailureHandler(function (error) {
            statusBtn.disabled = false;
            statusBtn.textContent = 'Cek Status';
            showMessage(errorText(error), 'error');
          })
          .getAccessRequestStatus();
      });
    })();
  </script>
</body>
</html>`;
}

/* ================= HELPER ================= */

function cleanHeader_(h) {
  return String(h).trim().toLowerCase();
}

function cleanText_(v) {
  if (!v) return "";
  return String(v).replace(/\s+/g, ' ').trim();
}

/* ================= UNIVERSAL DATE PARSER ================= */

function parseDate_(val) {
  if (val === null || val === undefined || val === '') return null;

  if (Object.prototype.toString.call(val) === '[object Date]') {
    return isNaN(val.getTime()) ? null : val;
  }

  function createLocalDate(year, month, day, hour, minute, second) {
    const date = new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour || 0),
      Number(minute || 0),
      Number(second || 0)
    );
    if (
      date.getFullYear() !== Number(year) ||
      date.getMonth() !== Number(month) - 1 ||
      date.getDate() !== Number(day)
    ) return null;
    return date;
  }

  if (typeof val === 'string') {
    const text = val.trim();
    let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) return createLocalDate(match[3], match[2], match[1], match[4], match[5], match[6]);

    match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) return createLocalDate(match[1], match[2], match[3], match[4], match[5], match[6]);

    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
      const isoDate = new Date(text);
      return isNaN(isoDate.getTime()) ? null : isoDate;
    }
    return null;
  }

  if (typeof val === 'number' && isFinite(val)) {
    const wholeDays = Math.floor(val);
    const fraction = val - wholeDays;
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + wholeDays);
    base.setMilliseconds(Math.round(fraction * 86400000));
    return isNaN(base.getTime()) ? null : base;
  }

  return null;
}

/* ================= CACHE & MAIN FUNCTION ================= */

const DASHBOARD_CACHE_TTL_SECONDS = 300;
const DASHBOARD_CACHE_CHUNK_SIZE = 80000;
const MONITORING_CACHE_PREFIX = 'MONITORING_V5';
const MONITORING_INDEX_CACHE_PREFIX = 'MONITORING_MONTH_INDEX_V1';
const MONITORING_CACHE_TTL_SECONDS = 21600;
const MONITORING_INDEX_EPOCH_PROPERTY = 'MONITORING_INDEX_EPOCH';
const MONITORING_DATE_COLUMN_PROPERTY = 'MONITORING_DATE_COLUMN';
const MONITORING_PERSISTENT_INDEX_PROPERTY = 'MONITORING_INDEX_SNAPSHOT_V1';
const MONITORING_PERSISTENT_INDEX_MAX_CHARS = 8000;
const MONITORING_RANGE_MERGE_GAP = 12;
const MONITORING_MAX_READ_BLOCKS = 4;
const MONITORING_UPDATE_TOKEN_CACHE_KEY = 'MONITORING_UPDATE_TOKEN_V1';
const MONITORING_UPDATE_TOKEN_CACHE_SECONDS = 30;
const SPJ_CACHE_PREFIX = 'SPJ_V3';
const SPJ_SNAPSHOT_CACHE_TTL_SECONDS = 21600;
const MONITORING_TRIGGERS_READY_PROPERTY = 'MONITORING_TRIGGERS_READY';
const SPJ_QUERY_CACHE_TTL_SECONDS = 120;
const SPJ_CACHE_EPOCH_PROPERTY = 'SPJ_CACHE_EPOCH';
const SPJ_UPDATE_TOKEN_CACHE_KEY = 'SPJ_UPDATE_TOKEN_V1';
const SPJ_UPDATE_TOKEN_CACHE_SECONDS = 30;

function cacheDigest_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || 'initial'),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '').slice(0, 24);
}

function cacheKey_(prefix, token, suffix) {
  return [prefix, cacheDigest_(token), suffix || 'data'].join(':');
}

function getCachedObject_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const rawMeta = cache.get(key + ':meta');
    if (!rawMeta) return null;

    const meta = JSON.parse(rawMeta);
    if (!meta || !meta.parts) return null;

    const partKeys = [];
    for (let i = 0; i < meta.parts; i++) partKeys.push(key + ':part:' + i);

    const cachedParts = cache.getAll(partKeys);
    const encodedParts = [];
    for (let i = 0; i < partKeys.length; i++) {
      if (!cachedParts[partKeys[i]]) return null;
      encodedParts.push(cachedParts[partKeys[i]]);
    }

    const bytes = Utilities.base64Decode(encodedParts.join(''));
    const json = Utilities.ungzip(Utilities.newBlob(bytes)).getDataAsString('UTF-8');
    return JSON.parse(json);
  } catch (error) {
    console.warn('Cache read gagal:', key, error);
    return null;
  }
}

function putCachedObject_(key, value, ttlSeconds) {
  try {
    const cache = CacheService.getScriptCache();
    const json = JSON.stringify(value);
    const compressed = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
    const encoded = Utilities.base64Encode(compressed.getBytes());
    const parts = [];

    for (let offset = 0; offset < encoded.length; offset += DASHBOARD_CACHE_CHUNK_SIZE) {
      parts.push(encoded.slice(offset, offset + DASHBOARD_CACHE_CHUNK_SIZE));
    }

    const entries = {};
    parts.forEach(function(part, index) {
      entries[key + ':part:' + index] = part;
    });
    entries[key + ':meta'] = JSON.stringify({ parts: parts.length });

    cache.putAll(entries, ttlSeconds || DASHBOARD_CACHE_TTL_SECONDS);
    return true;
  } catch (error) {
    console.warn('Cache write gagal:', key, error);
    return false;
  }
}

function serializeForClient_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || 'Asia/Jakarta',
      "yyyy-MM-dd'T'HH:mm:ss"
    );
  }
  if (value === undefined || value === null) return '';
  return value;
}

function getMonitoringUpdateToken_() {
  const props = PropertiesService.getScriptProperties();
  const propertyToken = props.getProperty('LAST_UPDATE');
  if (propertyToken && props.getProperty(MONITORING_TRIGGERS_READY_PROPERTY) === 'true') {
    return propertyToken;
  }

  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get(MONITORING_UPDATE_TOKEN_CACHE_KEY);
  if (cachedToken) return cachedToken;

  try {
    const driveToken = String(DriveApp.getFileById(SPREADSHEET_ID).getLastUpdated().getTime());
    cache.put(
      MONITORING_UPDATE_TOKEN_CACHE_KEY,
      driveToken,
      MONITORING_UPDATE_TOKEN_CACHE_SECONDS
    );
    return driveToken;
  } catch (error) {
    console.warn('Drive update token tidak tersedia:', error);
  }

  return 'initial';
}

function getMonitoringIndexToken_() {
  return PropertiesService.getScriptProperties()
    .getProperty(MONITORING_INDEX_EPOCH_PROPERTY) || 'initial';
}

function touchMonitoringIndex_() {
  const token = String(Date.now());
  PropertiesService.getScriptProperties()
    .setProperty(MONITORING_INDEX_EPOCH_PROPERTY, token);
  return token;
}

function getPersistentMonitoringIndex_(indexToken) {
  try {
    const raw = PropertiesService.getScriptProperties()
      .getProperty(MONITORING_PERSISTENT_INDEX_PROPERTY);
    if (!raw) return null;
    const stored = JSON.parse(raw);
    return stored && stored.token === indexToken ? stored.index : null;
  } catch (error) {
    console.warn('Indeks monitoring persisten tidak dapat dibaca:', error);
    return null;
  }
}

function putPersistentMonitoringIndex_(indexToken, index) {
  try {
    const payload = JSON.stringify({ token: indexToken, index: index });
    const props = PropertiesService.getScriptProperties();
    if (payload.length <= MONITORING_PERSISTENT_INDEX_MAX_CHARS) {
      props.setProperty(MONITORING_PERSISTENT_INDEX_PROPERTY, payload);
    } else {
      props.deleteProperty(MONITORING_PERSISTENT_INDEX_PROPERTY);
    }
  } catch (error) {
    console.warn('Indeks monitoring persisten tidak dapat disimpan:', error);
  }
}

function openMonitoringSource_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan');
  return {
    spreadsheet: spreadsheet,
    sheet: sheet,
    masterLookup: getOperatingPointMappingLookup_(spreadsheet),
    timezone: spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Jakarta',
    scriptTimezone: Session.getScriptTimeZone() || 'Asia/Jakarta'
  };
}

function getMonitoringDateKey_(value, timezone, scriptTimezone) {
  const date = parseDate_(value);
  if (!date) return '';
  if (timezone === scriptTimezone) {
    return [
      String(date.getFullYear()).padStart(4, '0'),
      String(date.getMonth() + 1).padStart(2, '0'),
      String(date.getDate()).padStart(2, '0')
    ].join('-');
  }
  return Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
}

function buildMonitoringMonthIndex_(source) {
  const sheet = source.sheet;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) {
    return {
      meta: { months: [], defaultMonth: '', timezone: source.timezone },
      spansByMonth: {},
      schema: null
    };
  }

  const rawHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const header = rawHeader.map(cleanHeader_);
  const colOP = header.indexOf('kode op');
  const colNamaOP = header.indexOf('nama op');
  const colCustomer = header.indexOf('customer');
  const colProduk = header.indexOf('produk');
  const colTOH = header.indexOf('nama toh/woh');
  const colDate = header.indexOf('tanggal sla');
  const colStatus = header.indexOf('status rekon');
  if (colOP === -1 || colNamaOP === -1 || colCustomer === -1 || colDate === -1 || colStatus === -1) {
    throw new Error('Header wajib tidak ditemukan. Pastikan ada: Kode OP, Nama OP, Customer, Produk, Tanggal SLA, Status Rekon');
  }

  const coreColumns = [colOP, colNamaOP, colCustomer, colProduk, colTOH, colDate, colStatus]
    .filter(function(index) { return index > -1; });
  const firstCoreColumn = Math.min.apply(null, coreColumns);
  const lastCoreColumn = Math.max.apply(null, coreColumns);
  const dateValues = sheet.getRange(2, colDate + 1, lastRow - 1, 1).getValues();
  const spansByMonth = {};

  dateValues.forEach(function(row, index) {
    const dateKey = getMonitoringDateKey_(row[0], source.timezone, source.scriptTimezone);
    if (!dateKey) return;
    const monthKey = dateKey.substring(0, 7);
    const sheetRow = index + 2;
    if (!spansByMonth[monthKey]) spansByMonth[monthKey] = [];
    const spans = spansByMonth[monthKey];
    const previous = spans.length ? spans[spans.length - 1] : null;
    if (previous && previous[0] + previous[1] === sheetRow) {
      previous[1]++;
    } else {
      spans.push([sheetRow, 1]);
    }
  });

  const months = Object.keys(spansByMonth).sort().reverse();
  PropertiesService.getScriptProperties()
    .setProperty(MONITORING_DATE_COLUMN_PROPERTY, String(colDate + 1));

  return {
    meta: {
      months: months,
      defaultMonth: months[0] || '',
      timezone: source.timezone
    },
    spansByMonth: spansByMonth,
    schema: {
      firstColumn: firstCoreColumn + 1,
      columnCount: lastCoreColumn - firstCoreColumn + 1,
      op: colOP - firstCoreColumn,
      namaOP: colNamaOP - firstCoreColumn,
      customer: colCustomer - firstCoreColumn,
      produk: colProduk > -1 ? colProduk - firstCoreColumn : -1,
      toh: colTOH > -1 ? colTOH - firstCoreColumn : -1,
      date: colDate - firstCoreColumn,
      status: colStatus - firstCoreColumn
    }
  };
}

function normalizeMonitoringMonth_(selectedMonth, meta) {
  const requested = String(selectedMonth || '').trim();
  return meta.months.indexOf(requested) > -1 ? requested : meta.defaultMonth;
}

function coalesceMonitoringSpans_(spans) {
  if (!spans || !spans.length) return [];
  const merged = [];
  spans.forEach(function(span) {
    const start = Number(span[0]);
    const end = start + Number(span[1]) - 1;
    const previous = merged.length ? merged[merged.length - 1] : null;
    const previousEnd = previous ? previous[0] + previous[1] - 1 : 0;
    if (previous && start - previousEnd - 1 <= MONITORING_RANGE_MERGE_GAP) {
      previous[1] = Math.max(previousEnd, end) - previous[0] + 1;
    } else {
      merged.push([start, end - start + 1]);
    }
  });
  if (merged.length <= MONITORING_MAX_READ_BLOCKS) return merged;

  const grouped = [];
  const groupSize = Math.ceil(merged.length / MONITORING_MAX_READ_BLOCKS);
  for (let index = 0; index < merged.length; index += groupSize) {
    const group = merged.slice(index, index + groupSize);
    const start = group[0][0];
    const last = group[group.length - 1];
    const end = last[0] + last[1] - 1;
    grouped.push([start, end - start + 1]);
  }
  return grouped;
}

function readMonitoringMonthRows_(source, monthIndex, month) {
  if (!month || !monthIndex.schema) return [];
  const spans = coalesceMonitoringSpans_(monthIndex.spansByMonth[month] || []);
  if (!spans.length) return [];
  const schema = monthIndex.schema;
  const rows = [];

  spans.forEach(function(span) {
    const values = source.sheet.getRange(
      span[0],
      schema.firstColumn,
      span[1],
      schema.columnCount
    ).getValues();

    values.forEach(function(row) {
      const dateKey = getMonitoringDateKey_(row[schema.date], source.timezone, source.scriptTimezone);
      if (!dateKey || dateKey.substring(0, 7) !== month) return;
      const op = cleanText_(row[schema.op]);
      if (!op) return;
      let status = cleanText_(row[schema.status]);
      status = status.toLowerCase() === 'rekon' ? 'Rekon' : (status || 'Non Rekon');
      const mapping = findOperatingPointMapping_(
        source.masterLookup,
        op,
        cleanText_(row[schema.namaOP]),
        cleanText_(row[schema.customer])
      );
      const mappedTOH = mapping ? mapping.toh : '';
      rows.push([
        op,
        cleanText_(row[schema.namaOP]),
        cleanText_(row[schema.customer]),
        schema.produk > -1 ? cleanText_(row[schema.produk]) : '-',
        mappedTOH || (schema.toh > -1 && row[schema.toh] ? cleanText_(row[schema.toh]) : '-'),
        Number(dateKey.substring(8, 10)),
        status
      ]);
    });
  });
  return rows;
}

function expandMonitoringRows_(rows, month) {
  return (rows || []).map(function(row) {
    return {
      op: row[0],
      namaop: row[1],
      customer: row[2],
      produk: row[3],
      toh: row[4],
      month: month,
      day: String(row[5]),
      status: row[6]
    };
  });
}

function formatMonitoringLastUpdate_(updateToken, timezone) {
  let date = null;
  if (/^\d+$/.test(String(updateToken || ''))) date = new Date(Number(updateToken));
  else date = new Date(String(updateToken || ''));
  return date && !isNaN(date.getTime())
    ? Utilities.formatDate(date, timezone || 'Asia/Jakarta', 'dd MMM yyyy HH:mm')
    : '-';
}

function getMonitoringData(selectedMonth, forceRefresh) {
  assertAuthorized_();
  const startedAt = Date.now();
  let updateToken = '';
  try {
    updateToken = getMonitoringUpdateToken_();
    const baseKey = cacheKey_(MONITORING_CACHE_PREFIX, updateToken, 'snapshot');
    const indexToken = getMonitoringIndexToken_();
    const indexKey = cacheKey_(MONITORING_INDEX_CACHE_PREFIX, indexToken, 'index');
    let monthIndex = getCachedObject_(indexKey);
    if (!monthIndex) {
      monthIndex = getPersistentMonitoringIndex_(indexToken);
      if (monthIndex) putCachedObject_(indexKey, monthIndex, MONITORING_CACHE_TTL_SECONDS);
    }
    if (forceRefresh) monthIndex = null;

    let source = null;
    let indexBuilt = false;
    let monthReadFromSheet = false;
    if (!monthIndex) {
      source = openMonitoringSource_();
      monthIndex = buildMonitoringMonthIndex_(source);
      putCachedObject_(indexKey, monthIndex, MONITORING_CACHE_TTL_SECONDS);
      putPersistentMonitoringIndex_(indexToken, monthIndex);
      indexBuilt = true;
    }

    const month = normalizeMonitoringMonth_(selectedMonth, monthIndex.meta);
    let compactRows = !forceRefresh && month
      ? getCachedObject_(baseKey + ':month:' + month)
      : (month ? null : []);

    if (compactRows === null) {
      const lock = LockService.getScriptLock();
      const locked = lock.tryLock(5000);
      try {
        if (!forceRefresh) compactRows = getCachedObject_(baseKey + ':month:' + month);
        if (compactRows === null) {
          if (!source) source = openMonitoringSource_();
          compactRows = readMonitoringMonthRows_(source, monthIndex, month);
          putCachedObject_(baseKey + ':month:' + month, compactRows, MONITORING_CACHE_TTL_SECONDS);
          monthReadFromSheet = true;
        }
      } finally {
        if (locked) lock.releaseLock();
      }
    }

    return {
      monitoring: expandMonitoringRows_(compactRows || [], month),
      months: monthIndex.meta.months,
      selectedMonth: month,
      lastUpdate: formatMonitoringLastUpdate_(updateToken, monthIndex.meta.timezone),
      updateToken: updateToken,
      cached: !indexBuilt && !monthReadFromSheet,
      performance: {
        totalMs: Date.now() - startedAt,
        indexBuilt: indexBuilt,
        monthReadFromSheet: monthReadFromSheet,
        resultRows: (compactRows || []).length
      }
    };
  } catch (error) {
    console.error(error);
    return {
      error: error.toString(),
      monitoring: [],
      months: [],
      selectedMonth: '',
      lastUpdate: '-',
      updateToken: updateToken || getMonitoringUpdateToken_(),
      cached: false
    };
  }
}
/* ================= AUTO SEND EMAIL TOH/WOH ================= */

const EMAIL_LOG_HEADERS = [
  'Timestamp', 'Periode', 'TOH', 'Email', 'Jumlah_Kasus',
  'Total_Tanggal', 'Priority', 'Status', 'Send_Key'
];

function ensureEmailLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName('Email_Log');
  if (!sheet) sheet = spreadsheet.insertSheet('Email_Log');
  if (sheet.getLastRow() > 0 && cleanHeader_(sheet.getRange(1, 1).getValue()) !== 'timestamp') {
    sheet.insertRowsBefore(1, 1);
  }
  sheet.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length).setValues([EMAIL_LOG_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getSuccessfulEmailSendKeys_(logSheet) {
  const keys = {};
  if (!logSheet || logSheet.getLastRow() < 2) return keys;
  const rows = logSheet.getRange(2, 8, logSheet.getLastRow() - 1, 2).getValues();
  rows.forEach(function(row) {
    if (String(row[0] || '').toUpperCase() === 'SUCCESS' && row[1]) {
      keys[String(row[1])] = true;
    }
  });
  return keys;
}

function sendDailyNonRekonEmail(event) {
  assertAuthorizedOrTrustedTrigger_(event);

  const executionLock = LockService.getScriptLock();
  if (!executionLock.tryLock(10000)) {
    throw new Error('Pengiriman email lain masih berjalan.');
  }
  try {
    return sendDailyNonRekonEmail_();
  } finally {
    executionLock.releaseLock();
  }
}

function sendDailyNonRekonEmail_() {
  const today = new Date();
  const day = today.getDay();

  if (day === 0 || day === 6) {
    Logger.log("==================================");
    Logger.log("EMAIL REPORT NON REKON");
    Logger.log("Status : STOP (Weekend)");
    Logger.log("Tanggal Run : " + today);
    Logger.log("Keterangan : Email tidak dikirim karena hari Sabtu/Minggu");
    Logger.log("==================================");
    return;
  }

  const DASHBOARD_URL = "https://script.google.com/macros/s/AKfycbwkI628xKLAchR69opeYdGp4gdTYmwS45XaTJjN53-bHr2sHj9B8XmnKirgSnlllsDF8A/exec";

  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db       = ss.getSheetByName("Database");
  const master   = ss.getSheetByName("Master_OP_OH");

  if (!db)     throw new Error("Sheet Database tidak ditemukan");
  if (!master) throw new Error("Sheet Master_OP_OH tidak ditemukan");

  const logSheet = ensureEmailLogSheet_(ss);
  const successfulSendKeys = getSuccessfulEmailSendKeys_(logSheet);

  const dbData   = db.getDataRange().getValues();
  const dbHeader = dbData.shift();

  const masterData   = master.getDataRange().getValues();
  const masterHeader = masterData.shift();

  const normalizedDbHeader = dbHeader.map(cleanHeader_);
  const normalizedMasterHeader = masterHeader.map(cleanHeader_);
  const colOP       = normalizedDbHeader.indexOf('kode op');
  const colNamaOP   = normalizedDbHeader.indexOf('nama op');
  const colCustomer = normalizedDbHeader.indexOf('customer');
  const colProduk   = normalizedDbHeader.indexOf('produk');
  const colTanggal  = normalizedDbHeader.indexOf('tanggal sla');
  const colStatus   = normalizedDbHeader.indexOf('status rekon');
  const colTOH      = normalizedDbHeader.indexOf('nama toh/woh');

  const colMasterTOH = normalizedMasterHeader.indexOf('nama toh/woh');
  const colEmail     = normalizedMasterHeader.indexOf('email toh');
  if ([colOP, colNamaOP, colCustomer, colProduk, colTanggal, colStatus, colTOH].indexOf(-1) > -1) {
    throw new Error('Header Database untuk email tidak lengkap.');
  }
  if (colMasterTOH === -1 || colEmail === -1) {
    throw new Error("Kolom 'Nama TOH/WOH' atau 'Email TOH' belum ada di Master_OP_OH");
  }

  const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Jakarta';
  const runDateKey = Utilities.formatDate(today, timezone, 'yyyy-MM-dd');
  let remainingQuota = MailApp.getRemainingDailyQuota();

  /* 1. EMAIL MAPPING */
  let emailMap = {};
  masterData.forEach(r => {
    if (r[colMasterTOH] && r[colEmail]) {
      const name = r[colMasterTOH].toString().trim().toLowerCase();
      const email = normalizeEmail_(r[colEmail]);
      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) emailMap[name] = email;
    }
  });

  /* 2. CARI SLA TERAKHIR */
  let lastSLA = null;
  dbData.forEach(r => {
    if (!r[colTanggal]) return;
    const d = parseDate_(r[colTanggal]);
    if (!d) return;
    if (!lastSLA || d > lastSLA) lastSLA = d;
  });

  if (!lastSLA) {
    Logger.log("Tidak ada tanggal SLA ditemukan");
    return;
  }

  /* 3. MAX 3 BULAN KE BELAKANG */
  let monthsToCheck = [];
  for (let i = 0; i < 3; i++) {
    let d = new Date(lastSLA.getFullYear(), lastSLA.getMonth() - i, 1);
    monthsToCheck.push(Utilities.formatDate(d, timezone, "yyyy-MM"));
  }

  const bulanIndo = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];

  /* 4. KUMPULKAN DATA */
  let laporan = {};
  dbData.forEach(r => {
    if (!r[colTanggal]) return;
    const tgl = parseDate_(r[colTanggal]);
    if (!tgl) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "yyyy-MM");
    if (!monthsToCheck.includes(monthKey)) return;
    if (String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    const toh = String(r[colTOH]).trim();
    if (!toh) return;

    const key = JSON.stringify([
      cleanText_(r[colOP]),
      cleanText_(r[colCustomer]),
      cleanText_(r[colProduk])
    ]);
    if (!laporan[monthKey])       laporan[monthKey] = {};
    if (!laporan[monthKey][toh])  laporan[monthKey][toh] = {};
    if (!laporan[monthKey][toh][key]) {
      laporan[monthKey][toh][key] = {
        op:      r[colOP],
        nama:    r[colNamaOP],
        cust:    r[colCustomer],
        produk:  r[colProduk],
        tanggal: []
      };
    }

    const dayNum = tgl.getDate();
    if (!laporan[monthKey][toh][key].tanggal.includes(dayNum))
      laporan[monthKey][toh][key].tanggal.push(dayNum);
  });

  /* 5. KIRIM EMAIL */
  let totalEmailSent = 0;

  Object.keys(laporan).forEach(monthKey => {
    const [year, month] = monthKey.split("-");
    const periodeBulan  = bulanIndo[parseInt(month) - 1] + " " + year;

    Object.keys(laporan[monthKey]).forEach(toh => {
      const email = emailMap[toh.toLowerCase()];
      if (!email) return;

      const sendKey = [runDateKey, monthKey, normalizeEmail_(email), cleanText_(toh).toLowerCase()].join('|');
      if (successfulSendKeys[sendKey]) {
        Logger.log('Email dilewati karena sudah terkirim: ' + sendKey);
        return;
      }

      let rows = "";
      let totalTanggal = 0;

      Object.values(laporan[monthKey][toh]).forEach(d => {
        totalTanggal += d.tanggal.length;
        const tanggalList = d.tanggal.sort((a, b) => a - b).map(t =>
          `<span style="display:inline-block;background:${d.tanggal.length > 5 ? '#fee2e2' : '#fff1f0'};color:#b91c1c;padding:3px 8px;border-radius:20px;font-size:11px;font-weight:600;margin:2px;border:1px solid #ffcdd2;">${t}</span>`
        ).join(' ');

        rows += `
        <tr style="border-bottom:1px solid #edf2f7;">
          <td style="padding:12px;">
            <span style="background:linear-gradient(135deg,#667eea15,#764ba215);padding:4px 10px;border-radius:6px;font-size:12px;font-weight:600;color:#1e293b;border-left:2px solid #6366f1;">${escapeHtml_(d.op)}</span>
          </td>
          <td style="padding:12px;">
            <div style="font-weight:600;color:#0f172a;font-size:13px;">${escapeHtml_(d.nama)}</div>
          </td>
          <td style="padding:12px;">
            <span style="background:#e0f2fe;color:#0369a1;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:500;">${escapeHtml_(d.cust)}</span>
          </td>
          <td style="padding:12px;color:#334155;font-size:12px;">${escapeHtml_(d.produk)}</td>
          <td style="padding:12px;"><div style="max-width:250px;">${tanggalList}</div></td>
        </tr>`;
      });

      if (!rows) return;

      const jumlahKasus = Object.values(laporan[monthKey][toh]).length;
      const priority    = totalTanggal > 10 ? "High" : totalTanggal > 5 ? "Medium" : "Low";
      const priorityColor = priority === "High" ? "#ef4444" : priority === "Medium" ? "#f59e0b" : "#10b981";
      const priorityBg    = priority === "High" ? "#fee2e2" : priority === "Medium" ? "#fef3c7" : "#d1fae5";

      const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
@media only screen and (max-width:480px){
  .greeting-card{flex-direction:column!important;align-items:flex-start!important;gap:10px!important;}
  .priority-badge{align-self:flex-start!important;}
  .stats-card{padding:8px!important;}
  .stats-value{font-size:18px!important;}
  .action-button{padding:10px 16px!important;font-size:12px!important;}
}
</style>
</head>
<body style="margin:0;padding:10px;background:#f1f5f9;font-family:'Inter',Helvetica,Arial,sans-serif;">
<div style="max-width:900px;margin:0 auto;width:100%;">
<div style="background:white;border-radius:24px;overflow:hidden;box-shadow:0 20px 30px -10px rgba(0,0,0,.15);">
<div style="background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);padding:22px 20px 18px;text-align:center;">
  <div style="display:inline-block;background:rgba(255,255,255,.12);padding:6px 16px;border-radius:30px;margin-bottom:10px;">
    <span style="color:#fff;font-size:11px;font-weight:600;letter-spacing:.6px;">⚡ TRACKING DATA REKON</span>
  </div>
  <h1 style="color:white;font-size:28px;font-weight:700;margin:0 0 5px;line-height:1.2;">Monitoring Non-Rekon Bitrans</h1>
  <p style="color:#cbd5e1;font-size:18px;margin:5px 0 0;">Periode <span style="color:white;font-weight:600;">${periodeBulan}</span></p>
</div>
<div style="padding:20px;">
  <div class="greeting-card" style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;background:#f8fafc;padding:12px 16px;border-radius:16px;border:1px solid #e2e8f0;">
    <div>
      <h2 style="color:#0f172a;margin:0 0 3px;font-size:15px;font-weight:600;">Yth. <span style="color:#6366f1;">${escapeHtml_(toh)}</span></h2>
      <p style="color:#475569;margin:0;font-size:12px;">Daftar transaksi Non-Rekon yang perlu diselesaikan</p>
    </div>
    <div class="priority-badge" style="background:${priorityBg};padding:6px 12px;border-radius:12px;text-align:center;border-left:3px solid ${priorityColor};white-space:nowrap;margin-left:16px;">
      <span style="color:#64748b;font-size:9px;text-transform:uppercase;">Priority</span>
      <div style="color:${priorityColor};font-size:14px;font-weight:700;">${priority}</div>
    </div>
  </div>
  <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
    <tr>
      <td style="width:33.33%;padding-right:6px;">
        <div class="stats-card" style="border:1px solid #e2e8f0;border-radius:16px;padding:12px;">
          <div style="display:flex;align-items:center;gap:10px;">
            <span style="font-size:20px;">📋</span>
            <div><span style="color:#64748b;font-size:11px;">Total Kasus</span>
            <div class="stats-value" style="color:#0f172a;font-size:20px;font-weight:700;">${jumlahKasus}</div></div>
          </div>
        </div>
      </td>
      <td style="width:33.33%;padding:0 3px;">
        <div class="stats-card" style="border:1px solid #e2e8f0;border-radius:16px;padding:12px;">
          <div style="display:flex;align-items:center;gap:10px;">
            <span style="font-size:20px;">📅</span>
            <div><span style="color:#64748b;font-size:11px;">Total Tanggal</span>
            <div class="stats-value" style="color:#0f172a;font-size:20px;font-weight:700;">${totalTanggal}</div></div>
          </div>
        </div>
      </td>
      <td style="width:33.33%;padding-left:6px;">
        <div class="stats-card" style="border:1px solid #e2e8f0;border-radius:16px;padding:12px;">
          <div style="display:flex;align-items:center;gap:10px;">
            <span style="font-size:20px;">⚠️</span>
            <div><span style="color:#64748b;font-size:11px;">Perlu Tindakan</span>
            <div class="stats-value" style="color:#0f172a;font-size:20px;font-weight:700;">${jumlahKasus}</div></div>
          </div>
        </div>
      </td>
    </tr>
  </table>
  <div style="margin-bottom:20px;">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
      <h3 style="color:#0f172a;font-size:15px;font-weight:600;margin:0;">📊 Detail Transaksi</h3>
      <span style="background:#e2e8f0;color:#475569;padding:4px 10px;border-radius:30px;font-size:11px;font-weight:600;">${jumlahKasus} items</span>
    </div>
    <div style="border:1px solid #e2e8f0;border-radius:16px;overflow:hidden;">
      <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
        <table style="width:100%;border-collapse:collapse;font-size:12px;min-width:600px;">
          <thead>
            <tr style="background:#1e293b;">
              <th style="padding:10px 12px;text-align:left;color:white;font-weight:600;font-size:11px;white-space:nowrap;">Kode OP</th>
              <th style="padding:10px 12px;text-align:left;color:white;font-weight:600;font-size:11px;white-space:nowrap;">Nama OP</th>
              <th style="padding:10px 12px;text-align:left;color:white;font-weight:600;font-size:11px;white-space:nowrap;">Customer</th>
              <th style="padding:10px 12px;text-align:left;color:white;font-weight:600;font-size:11px;white-space:nowrap;">Produk</th>
              <th style="padding:10px 12px;text-align:left;color:white;font-weight:600;font-size:11px;white-space:nowrap;">Tanggal</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>
  </div>
  <div style="text-align:center;margin:20px 0 15px;">
    <a href="${DASHBOARD_URL}" class="action-button" style="background:linear-gradient(135deg,#6366f1,#8b5cf6);color:white;padding:10px 24px;border-radius:30px;text-decoration:none;font-size:13px;font-weight:600;display:inline-block;box-shadow:0 4px 10px -2px rgba(99,102,241,.3);">
      🔍 Buka Dashboard Monitoring
    </a>
  </div>
  <div style="border-top:1px solid #e2e8f0;padding-top:15px;text-align:center;">
    <p style="color:#64748b;font-size:11px;margin:0;">Email Otomatis • <strong>Sistem Monitoring TMS</strong></p>
    <p style="color:#94a3b8;font-size:10px;margin:5px 0 0;">© 2026 Team Digital • B-Log</p>
  </div>
</div>
</div>
</div>
</body></html>`;

      try {
        if (remainingQuota < 1) {
          logSheet.appendRow([new Date(), periodeBulan, toh, email, jumlahKasus, totalTanggal, priority, 'SKIPPED_QUOTA', sendKey]);
          return;
        }
        MailApp.sendEmail({
          to:       email,
          subject:  "⚠ [BITRANS] Non-Rekon " + periodeBulan + " • " + priority,
          body:     'Terdapat ' + jumlahKasus + ' transaksi Non-Rekon untuk periode ' + periodeBulan + '. Buka dashboard: ' + DASHBOARD_URL,
          htmlBody: html
        });
        remainingQuota--;
        totalEmailSent++;
        successfulSendKeys[sendKey] = true;
        logSheet.appendRow([new Date(), periodeBulan, toh, email, jumlahKasus, totalTanggal, priority, 'SUCCESS', sendKey]);
      } catch (err) {
        logSheet.appendRow([new Date(), periodeBulan, toh, email, jumlahKasus, totalTanggal, priority, 'FAILED: ' + err.message, sendKey]);
      }
    });
  });

  Logger.log("==================================");
  Logger.log("EMAIL REPORT NON REKON");
  Logger.log("Tanggal Run : " + new Date());
  Logger.log("Total Email Terkirim : " + totalEmailSent);
  Logger.log("==================================");
}

/* ================= DOWNLOAD REPORT EXCEL ================= */
// NOTE: Only one definition of downloadNonRekonReport (duplicate removed)

function downloadNonRekonReport(selectedMonth) {
  assertAuthorized_();
  selectedMonth = String(selectedMonth || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(selectedMonth)) {
    throw new Error('Format bulan tidak valid.');
  }

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db     = ss.getSheetByName("Database");
  if (!db) throw new Error('Sheet Database tidak ditemukan.');
  const data   = db.getDataRange().getValues();
  const header = data.shift();

  function findColumn(name) {
    return header.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase());
  }

  const colOP       = findColumn("Kode OP");
  const colNamaOP   = findColumn("Nama OP");
  const colCustomer = findColumn("Customer");
  const colProduk   = findColumn("Produk");
  const colTanggal  = findColumn("Tanggal SLA");
  const colStatus   = findColumn("Status Rekon");
  const colTOH      = findColumn("Nama TOH/WOH");

  if ([colOP, colNamaOP, colCustomer, colProduk, colTanggal, colStatus, colTOH].includes(-1)) {
    throw new Error("Header kolom tidak sesuai dengan database.");
  }

  const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Jakarta';

  let output = [["Kode OP", "Nama OP", "Customer", "Produk", "TOH", "Tanggal SLA", "Status"]];

  function safeExportText(value) {
    const text = value === null || value === undefined ? '' : String(value);
    return /^[=+\-@]/.test(text) ? "'" + text : text;
  }

  data.forEach(r => {
    if (!r[colTanggal]) return;
    const tgl = parseDate_(r[colTanggal]);
    if (!tgl) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "yyyy-MM");
    if (monthKey !== selectedMonth) return;
    if (String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    output.push([
      safeExportText(r[colOP]),
      safeExportText(r[colNamaOP]),
      safeExportText(r[colCustomer]),
      safeExportText(r[colProduk]),
      safeExportText(r[colTOH]),
      Utilities.formatDate(tgl, timezone, "d-MMMM-yyyy"),
      safeExportText(r[colStatus])
    ]);
  });

  if (output.length <= 1) throw new Error("Tidak ada data Non-Rekon pada bulan tersebut.");

  let fileId = '';
  try {
    const tempSS = SpreadsheetApp.create('TEMP_EXPORT_' + Utilities.getUuid());
    fileId = tempSS.getId();
    const sheet = tempSS.getActiveSheet();
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    sheet.getRange('A1:G1').setFontWeight('bold');
    sheet.autoResizeColumns(1, 7);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();

    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
      '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Gagal export XLSX. Respons Drive: HTTP ' + response.getResponseCode());
    }

    const blob = response.getBlob().setName('Rekap_Non_Rekon_' + selectedMonth + '.xlsx');
    return {
      fileName: blob.getName(),
      mimeType: blob.getContentType(),
      data: Utilities.base64Encode(blob.getBytes())
    };
  } finally {
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (cleanupError) {
        console.warn('File export sementara gagal dihapus:', fileId, cleanupError);
      }
    }
  }
}

/* ================= AUTO REFRESH MONITORING DASHBOARD ================= */

function updateDashboardTimestamp_() {
  const now = new Date().toISOString();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_UPDATE', now);
  props.setProperty('LAST_UPDATE_TS', String(Date.now()));
  CacheService.getScriptCache().remove(MONITORING_UPDATE_TOKEN_CACHE_KEY);
  Logger.log('Dashboard Updated : ' + now);
  return now;
}

function shouldTouchMonitoring_(event) {
  if (!event || !event.range) return false;
  const sheet = event.range.getSheet();
  return sheet && sheet.getName() === SHEET_NAME;
}

function shouldInvalidateMonitoringIndex_(event) {
  if (!event || !event.range || !shouldTouchMonitoring_(event)) return false;
  if (event.range.getRow() === 1) return true;
  const dateColumn = Number(
    PropertiesService.getScriptProperties().getProperty(MONITORING_DATE_COLUMN_PROPERTY)
  );
  if (!dateColumn) return true;
  return event.range.getColumn() <= dateColumn &&
    event.range.getLastColumn() >= dateColumn;
}

function shouldTouchAccess_(event) {
  if (!event || !event.range) return false;
  const sheetName = event.range.getSheet().getName();
  return sheetName === ACCESS_REQUEST_SHEET || sheetName === ACCESS_ALLOWLIST_SHEET;
}

function upsertAllowlistEmail_(email, name, note) {
  const normalized = normalizeEmail_(email);
  if (!normalized) return;

  const sheet = ensureAccessAllowlistSheet_();
  const lastRow = sheet.getLastRow();
  let targetRow = lastRow + 1;
  if (lastRow >= 2) {
    const emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let index = 0; index < emails.length; index++) {
      if (normalizeEmail_(emails[index][0]) === normalized) {
        targetRow = index + 2;
        break;
      }
    }
  }

  sheet.getRange(targetRow, 1, 1, ACCESS_ALLOWLIST_HEADERS.length).setValues([[
    normalized,
    cleanText_(name),
    'ACTIVE',
    cleanText_(note) || 'Disetujui dari Access_Request',
    new Date()
  ]]);
  sheet.getRange(targetRow, 5).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

function handleAccessReviewEdit_(event) {
  if (!event || !event.range) return;
  const sheet = event.range.getSheet();
  const firstColumn = event.range.getColumn();
  const lastColumn = event.range.getLastColumn();
  if (sheet.getName() !== ACCESS_REQUEST_SHEET || firstColumn > 8 || lastColumn < 8) return;

  const firstRow = Math.max(event.range.getRow(), 2);
  const lastRow = event.range.getLastRow();
  if (lastRow < firstRow) return;

  const reviewer = event.user && event.user.getEmail
    ? normalizeEmail_(event.user.getEmail())
    : getActiveEmail_();
  const now = new Date();
  const rows = sheet.getRange(firstRow, 4, lastRow - firstRow + 1, 5).getValues();
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    rows.forEach(function(row, index) {
      const status = String(row[4] || '').trim().toUpperCase();
      if (['APPROVED', 'REJECTED', 'EMAIL_FAILED'].indexOf(status) === -1) return;
      const rowNumber = firstRow + index;
      clearLatestAccessRequestCache_(sheet.getRange(rowNumber, 3).getValue());
      sheet.getRange(rowNumber, 9, 1, 3).setValues([[
        reviewer || 'ADMIN_SHEET',
        now,
        now
      ]]);
      sheet.getRange(rowNumber, 10, 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
      if (status === 'APPROVED') {
        upsertAllowlistEmail_(row[1], row[0], 'Disetujui oleh ' + (reviewer || 'admin'));
      }
    });
  } finally {
    lock.releaseLock();
  }
  clearAccessRegistryCache_();
}

function handleSpreadsheetEdit(event) {
  try {
    if (!event || !event.range) return;
    const source = event.source || event.range.getSheet().getParent();
    if (!source || String(source.getId()) !== String(SPREADSHEET_ID)) return;
    if (shouldTouchAccess_(event)) {
      handleAccessReviewEdit_(event);
      clearAccessRegistryCache_();
      return;
    }
    if (!shouldTouchMonitoring_(event)) return;

    if (shouldInvalidateMonitoringIndex_(event)) touchMonitoringIndex_();

    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const last = Number(props.getProperty('LAST_UPDATE_TS')) || 0;
    if (now - last >= 2000) updateDashboardTimestamp_();
  } catch (error) {
    console.error('handleSpreadsheetEdit gagal:', error);
  }
}

function onEdit(e) {
  handleSpreadsheetEdit(e);
}

function handleSpreadsheetChange(e) {
  try {
    if (!e || !e.source || typeof e.source.getId !== 'function') return;
    if (String(e.source.getId()) !== String(SPREADSHEET_ID)) return;
    clearAccessRegistryCache_();
    touchMonitoringIndex_();

    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const last = Number(props.getProperty('LAST_UPDATE_TS')) || 0;
    if (now - last < 2000) return;

    updateDashboardTimestamp_();
  } catch (error) {
    Logger.log(error);
  }
}

function onChange(e) {
  handleSpreadsheetChange(e);
}

function setupProjectTriggers() {
  assertAdmin_();
  const handlers = {
    handleSpreadsheetEdit: true,
    handleSpreadsheetChange: true
  };
  const existing = ScriptApp.getProjectTriggers();
  const sourceId = String(SPREADSHEET_ID);

  function hasTrigger(handlerName, eventType) {
    return existing.some(function(trigger) {
      return trigger.getHandlerFunction() === handlerName &&
        trigger.getEventType() === eventType &&
        String(trigger.getTriggerSourceId() || '') === sourceId;
    });
  }

  if (!hasTrigger('handleSpreadsheetEdit', ScriptApp.EventType.ON_EDIT) &&
      !hasTrigger('onEdit', ScriptApp.EventType.ON_EDIT)) {
    ScriptApp.newTrigger('handleSpreadsheetEdit').forSpreadsheet(SPREADSHEET_ID).onEdit().create();
  }
  if (!hasTrigger('handleSpreadsheetChange', ScriptApp.EventType.ON_CHANGE) &&
      !hasTrigger('onChange', ScriptApp.EventType.ON_CHANGE)) {
    ScriptApp.newTrigger('handleSpreadsheetChange').forSpreadsheet(SPREADSHEET_ID).onChange().create();
  }

  touchMonitoringIndex_();
  updateDashboardTimestamp_();
  PropertiesService.getScriptProperties().setProperty(MONITORING_TRIGGERS_READY_PROPERTY, 'true');
  clearAccessRegistryCache_();
  return { success: true, installedHandlers: Object.keys(handlers) };
}

function getLastUpdate() {
  assertAuthorized_();
  return getMonitoringUpdateToken_();
}

/* =========================================================
   DASHBOARD 2 — Tracking SPJ
   ========================================================= */

const SPREADSHEET_ID_SPJ = "12FcleafDZHI-OkyxWEfFQM4BXcx2RBNIfG2ShnoptGo";
const SHEET_DATA          = "Data_Rekon";
const SHEET_MASTER_SPJ    = "Master_OP_OH";
const SHEET_CONFIG_SPJ    = "Config";

// Data SPJ menggunakan cache terpisah agar pembacaan Data_Rekon dan Master_OP_OH
// tidak diulang pada setiap perpindahan menu.
function isSPJDateColumn_(rawHeader, header) {
  if (rawHeader instanceof Date && !isNaN(rawHeader.getTime())) return true;
  const text = String(header || '').trim();
  return /^(?:\d{1,2}|\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?|\d{4}-\d{2}-\d{2})$/.test(text) ||
    /^tanggal\b/i.test(text);
}

function buildSPJSnapshot_(forceRefresh, spreadsheet, cacheEpoch) {
  try {
    const ss = spreadsheet || SpreadsheetApp.openById(SPREADSHEET_ID_SPJ);
    const tanggalUpdate = getSPJLastUpdate_(ss);
    const spjCacheKey = cacheKey_(
      SPJ_CACHE_PREFIX,
      tanggalUpdate + '|' + String(cacheEpoch || 'initial'),
      'payload'
    );

    if (!forceRefresh) {
      const cachedResult = getCachedObject_(spjCacheKey);
      if (cachedResult) {
        cachedResult.cached = true;
        return cachedResult;
      }
    }

    const dataSheet = ss.getSheetByName(SHEET_DATA);
    if (!dataSheet) throw new Error('Sheet "' + SHEET_DATA + '" tidak ditemukan.');

    const dataRows = dataSheet.getDataRange().getValues();
    if (dataRows.length < 2) {
      const emptyResult = {
        success: true,
        data: [],
        filters: { kodeOP: [], customer: [], produk: [], toh: [], adjustment: [] },
        tanggalColumns: [],
        headers: dataRows[0] || [],
        tanggalUpdate: tanggalUpdate,
        cached: false
      };
      putCachedObject_(spjCacheKey, emptyResult, SPJ_SNAPSHOT_CACHE_TTL_SECONDS);
      return emptyResult;
    }

    const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Jakarta';
    const rawHeaders = dataRows[0];
    const headers = rawHeaders.map(function(header) {
      if (header instanceof Date && !isNaN(header.getTime())) {
        return Utilities.formatDate(header, timezone, 'dd/MM/yyyy');
      }
      return String(header).trim();
    });
    const requiredHeaders = [
      'Kode OP', 'Nama OP', 'Customer', 'SPJ', 'Produk',
      'Tanggal SLA', 'UJP', 'Charging', 'Status Rekon',
      'Nama TOH/WOH', 'Status Adjustment'
    ];
    const missingHeaders = requiredHeaders.filter(function(header) {
      return headers.indexOf(header) === -1;
    });
    if (missingHeaders.length) {
      throw new Error('Header Data_Rekon tidak lengkap: ' + missingHeaders.join(', '));
    }
    if (new Set(headers).size !== headers.length) {
      throw new Error('Header Data_Rekon mengandung nama kolom duplikat.');
    }
    const fixedColumns = [
      'Kode OP', 'Nama OP', 'Customer', 'SPJ', 'Produk',
      'Tanggal SLA', 'UJP', 'Charging', 'Status Rekon',
      'Nama TOH/WOH', 'Status Adjustment'
    ];
    const tanggalColumns = headers.filter(function(header, index) {
      return fixedColumns.indexOf(header) === -1 && isSPJDateColumn_(rawHeaders[index], header);
    });
    const projectedHeaders = new Set(fixedColumns.concat(tanggalColumns));
    const projectedColumns = [];
    headers.forEach(function(header, index) {
      if (projectedHeaders.has(header)) projectedColumns.push([header, index]);
    });
    const data = [];
    const textFields = [
      'Kode OP', 'Nama OP', 'Customer', 'SPJ', 'Produk',
      'Status Rekon', 'Nama TOH/WOH', 'Status Adjustment'
    ];

    for (let rowIndex = 1; rowIndex < dataRows.length; rowIndex++) {
      const row = dataRows[rowIndex];
      const rowData = {};

      projectedColumns.forEach(function(column) {
        rowData[column[0]] = serializeForClient_(row[column[1]]);
      });

      textFields.forEach(function(field) {
        if (Object.prototype.hasOwnProperty.call(rowData, field)) {
          rowData[field] = cleanText_(rowData[field]);
        }
      });

      let statusAdj = cleanText_(rowData['Status Adjustment']);
      if (statusAdj.toLowerCase() === 'reject') statusAdj = 'No Adjustment';
      rowData['Status Adjustment'] = statusAdj;
      data.push(rowData);
    }

    const masterSheet = ss.getSheetByName(SHEET_MASTER_SPJ);
    if (masterSheet) {
      const lookup = getOperatingPointMappingLookup_(ss, SHEET_MASTER_SPJ);
      data.forEach(function(item) {
        const mapping = findOperatingPointMapping_(
          lookup,
          item['Kode OP'],
          item['Nama OP'],
          item['Customer']
        );
        if (mapping && mapping.toh) item['Nama TOH/WOH'] = cleanText_(mapping.toh);
      });
    }

    const uniqueSorted = function(field) {
      return Array.from(new Set(data.map(function(item) {
        return cleanText_(item[field]);
      }).filter(Boolean))).sort(function(a, b) {
        return a.localeCompare(b, 'id', { numeric: true });
      });
    };

    const result = {
      success: true,
      data: data,
      filters: {
        kodeOP: uniqueSorted('Kode OP'),
        customer: uniqueSorted('Customer'),
        produk: uniqueSorted('Produk'),
        toh: uniqueSorted('Nama TOH/WOH'),
        adjustment: uniqueSorted('Status Adjustment')
      },
      tanggalColumns: tanggalColumns,
      headers: headers,
      tanggalUpdate: tanggalUpdate,
      cached: false
    };

    putCachedObject_(spjCacheKey, result, SPJ_SNAPSHOT_CACHE_TTL_SECONDS);
    return result;
  } catch (error) {
    console.error(error);
    return { success: false, error: error.toString() };
  }
}

function getSPJData(request) {
  assertAuthorized_();
  const options = request && typeof request === 'object' ? request : {};
  const forceRefresh = request === true || Boolean(options.forceRefresh);
  const pageSize = Math.min(Math.max(Number(options.pageSize) || 20, 1), 100);
  const page = Math.max(Number(options.page) || 1, 1);
  const filter = options.filters && typeof options.filters === 'object' ? options.filters : {};
  const op = cleanText_(filter.kodeOP);
  const customer = cleanText_(filter.customer);
  const produk = cleanText_(filter.produk);
  const toh = cleanText_(filter.toh);
  const adjustment = cleanText_(filter.adjustment);
  const search = cleanText_(options.search).toLowerCase();
  const includeFilters = options.includeFilters !== false;
  const scriptCache = CacheService.getScriptCache();
  let ss = null;
  let tanggalUpdate = forceRefresh ? '' : scriptCache.get(SPJ_UPDATE_TOKEN_CACHE_KEY);
  if (!tanggalUpdate) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID_SPJ);
    tanggalUpdate = getSPJLastUpdate_(ss);
    scriptCache.put(SPJ_UPDATE_TOKEN_CACHE_KEY, tanggalUpdate, SPJ_UPDATE_TOKEN_CACHE_SECONDS);
  }
  const props = PropertiesService.getScriptProperties();

  if (forceRefresh) {
    props.setProperty(SPJ_CACHE_EPOCH_PROPERTY, String(Date.now()));
  }
  const cacheEpoch = props.getProperty(SPJ_CACHE_EPOCH_PROPERTY) || 'initial';
  const queryDescriptor = JSON.stringify({
    page: page,
    pageSize: pageSize,
    search: search,
    includeFilters: includeFilters,
    filters: [op, customer, produk, toh, adjustment]
  });
  const queryCacheKey = cacheKey_(
    SPJ_CACHE_PREFIX,
    tanggalUpdate + '|' + cacheEpoch,
    'query:' + cacheDigest_(queryDescriptor)
  );

  if (!forceRefresh) {
    const cachedResponse = getCachedObject_(queryCacheKey);
    if (cachedResponse) {
      cachedResponse.cached = true;
      return cachedResponse;
    }
  }

  const snapshot = buildSPJSnapshot_(
    forceRefresh,
    ss || SpreadsheetApp.openById(SPREADSHEET_ID_SPJ),
    cacheEpoch
  );
  if (!snapshot || !snapshot.success) return snapshot;

  const filtered = snapshot.data.filter(function(item) {
    if (op && item['Kode OP'] !== op) return false;
    if (customer && item['Customer'] !== customer) return false;
    if (produk && item['Produk'] !== produk) return false;
    if (toh && item['Nama TOH/WOH'] !== toh) return false;
    if (adjustment && item['Status Adjustment'] !== adjustment) return false;
    if (search) {
      const haystack = [
        item['Kode OP'], item['Nama OP'], item['Customer'], item['Produk'],
        item['Nama TOH/WOH'], item['Status Adjustment'], item['SPJ']
      ].join(' ').toLowerCase();
      if (haystack.indexOf(search) === -1) return false;
    }
    return true;
  });

  const summary = { total: filtered.length, process: 0, done: 0, noAdjustment: 0 };
  filtered.forEach(function(item) {
    const status = String(item['Status Adjustment'] || '').toLowerCase();
    if (status === 'process') summary.process++;
    else if (status === 'done') summary.done++;
    else if (status === 'no adjustment') summary.noAdjustment++;
  });

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const start = (currentPage - 1) * pageSize;
  const response = {
    success: true,
    data: filtered.slice(start, start + pageSize),
    filters: includeFilters ? snapshot.filters : null,
    tanggalColumns: snapshot.tanggalColumns,
    headers: snapshot.headers,
    tanggalUpdate: snapshot.tanggalUpdate,
    cached: snapshot.cached,
    summary: summary,
    pagination: {
      page: currentPage,
      pageSize: pageSize,
      totalRows: filtered.length,
      totalPages: totalPages
    }
  };
  putCachedObject_(queryCacheKey, response, SPJ_QUERY_CACHE_TTL_SECONDS);
  return response;
}

function getSPJLastUpdate_(ss) {
  try {
    const configSheet = ss.getSheetByName(SHEET_CONFIG_SPJ);
    if (!configSheet) return 'Belum ada update';

    const value = configSheet.getRange('A1').getValue();
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(
        value,
        ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Jakarta',
        "yyyy-MM-dd'T'HH:mm:ss"
      );
    }
    return cleanText_(value) || 'Belum ada update';
  } catch (error) {
    return 'Error baca update';
  }
}

/* ================= ADMINISTRATION & DATA CONFIG ================= */

const ADMIN_CONFIG_SHEET = 'Dashboard_Config';
const ADMIN_CONFIG_HEADERS = ['Config_Key', 'Config_Value', 'Description', 'Status', 'Updated_At', 'Updated_By'];
const ADMIN_MAPPING_HEADERS = [
  'Area', 'Kode Cabang', 'Nama Cabang', 'Kode OP', 'Nama OP', 'Customer',
  'Nama TOH/WOH', 'Email TOH'
];
const ADMIN_MAPPING_LEGACY_HEADERS = ['Mapping ID', 'Status', 'Updated At', 'Updated By'];
const ADMIN_MAPPING_REQUIRED_FIELDS = [
  { property: 'area', label: 'Area' },
  { property: 'kodeCabang', label: 'Kode Cabang' },
  { property: 'namaCabang', label: 'Nama Cabang' },
  { property: 'kodeOP', label: 'Kode OP' },
  { property: 'namaOP', label: 'Nama OP' },
  { property: 'customer', label: 'Customer' },
  { property: 'toh', label: 'Nama TOH/WOH' },
  { property: 'email', label: 'Email TOH' }
];
const ADMIN_SETTING_DEFINITIONS = [
  { key: 'DASHBOARD_TITLE', label: 'Judul dashboard', type: 'text', defaultValue: 'Monitoring Dashboard', maxLength: 80 },
  { key: 'REKON_ROWS_PER_PAGE', label: 'Baris Rekon per halaman', type: 'number', defaultValue: '20', min: 5, max: 100 },
  { key: 'SPJ_ROWS_PER_PAGE', label: 'Baris SPJ per halaman', type: 'number', defaultValue: '20', min: 5, max: 100 },
  { key: 'REKON_POLL_SECONDS', label: 'Interval refresh otomatis (detik)', type: 'number', defaultValue: '60', min: 30, max: 900 },
  { key: 'DEFAULT_ADMIN_COVERAGE', label: 'Filter awal coverage', type: 'select', defaultValue: 'UNCOVERED', options: ['UNCOVERED', 'ALL', 'COVERED'] },
  { key: 'LOOKER_STUDIO_URL', label: 'URL Looker Studio', type: 'url', defaultValue: 'https://lookerstudio.google.com/embed/reporting/c52cbc36-25b6-4581-8f34-1f6fe427ee74/page/p_m62ko2r60d', maxLength: 1000 }
];
const ADMIN_MAPPING_SOURCES = {
  REKON: {
    label: 'Rekon Bitrans',
    spreadsheetId: SPREADSHEET_ID,
    dataSheet: SHEET_NAME,
    masterSheet: 'Master_OP_OH'
  },
  SPJ: {
    label: 'Data Non Rekon',
    spreadsheetId: SPREADSHEET_ID_SPJ,
    dataSheet: SHEET_DATA,
    masterSheet: SHEET_MASTER_SPJ
  }
};

function getDashboardContext() {
  const access = assertAuthorized_();
  return {
    email: access.email || '',
    isAdmin: isAdminEmail_(access.email),
    settings: getDashboardSettings_()
  };
}

function buildOperatingPointKey_(kodeOP, namaOP, customer) {
  return [kodeOP, namaOP, customer].map(function(value) {
    return cleanText_(value).toLowerCase();
  }).join('|');
}

function normalizeMappingMatchText_(value) {
  return cleanText_(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildOperatingPointAliasKey_(kodeOP, customer) {
  return [kodeOP, customer].map(normalizeMappingMatchText_).join('|');
}

function findHeaderIndex_(columns, candidates) {
  for (let index = 0; index < candidates.length; index++) {
    if (columns[candidates[index]] !== undefined) return columns[candidates[index]];
  }
  return undefined;
}

function findOperatingPointMapping_(lookup, kodeOP, namaOP, customer) {
  if (!lookup) return null;
  const exactKey = buildOperatingPointKey_(kodeOP, namaOP, customer);
  if (lookup[exactKey]) return lookup[exactKey];
  const aliasKey = buildOperatingPointAliasKey_(kodeOP, customer);
  return lookup.__relaxed && lookup.__relaxed[aliasKey]
    ? lookup.__relaxed[aliasKey]
    : null;
}

function ensureHeaders_(sheet, requiredHeaders) {
  const lastColumn = sheet.getLastColumn();
  const existing = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
        return String(header || '').trim();
      })
    : [];
  const normalized = existing.map(cleanHeader_);
  const missing = requiredHeaders.filter(function(header) {
    return normalized.indexOf(cleanHeader_(header)) === -1;
  });
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  const headers = existing.concat(missing);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#006b49')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  return headers;
}

function removeLegacyAdministrationMappingColumns_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const legacyHeaders = {};
  ADMIN_MAPPING_LEGACY_HEADERS.forEach(function(header) {
    legacyHeaders[cleanHeader_(header)] = true;
  });
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnsToDelete = [];
  headers.forEach(function(header, index) {
    if (legacyHeaders[cleanHeader_(header)]) columnsToDelete.push(index + 1);
  });
  columnsToDelete.sort(function(a, b) { return b - a; }).forEach(function(column) {
    sheet.deleteColumn(column);
  });
}

function ensureOperatingPointMaster_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  removeLegacyAdministrationMappingColumns_(sheet);
  ensureHeaders_(sheet, ADMIN_MAPPING_HEADERS);
  return sheet;
}

function getOperatingPointMappingLookup_(spreadsheet, sheetName) {
  const targetSheetName = sheetName || 'Master_OP_OH';
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  const lookup = { __relaxed: {}, __conflicts: {}, __allExact: {} };
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return lookup;

  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(cleanHeader_);
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  const required = ['kode op', 'nama op', 'customer'];
  if (required.some(function(header) { return columns[header] === undefined; })) return lookup;
  const tohColumn = findHeaderIndex_(columns, ['nama toh/woh', 'toh', 'nama toh']);
  const emailColumn = findHeaderIndex_(columns, ['email toh', 'email toh/woh', 'email']);
  if (tohColumn === undefined) return lookup;

  values.forEach(function(row, index) {
    const key = buildOperatingPointKey_(
      row[columns['kode op']],
      row[columns['nama op']],
      row[columns.customer]
    );
    if (key === '||') return;
    const record = {
      rowNumber: index + 2,
      area: columns.area === undefined ? '' : cleanText_(row[columns.area]),
      kodeCabang: columns['kode cabang'] === undefined ? '' : cleanText_(row[columns['kode cabang']]),
      namaCabang: columns['nama cabang'] === undefined ? '' : cleanText_(row[columns['nama cabang']]),
      kodeOP: cleanText_(row[columns['kode op']]),
      namaOP: cleanText_(row[columns['nama op']]),
      customer: cleanText_(row[columns.customer]),
      toh: cleanText_(row[tohColumn]),
      email: emailColumn === undefined ? '' : normalizeEmail_(row[emailColumn])
    };
    lookup.__allExact[key] = record;
    const aliasKey = buildOperatingPointAliasKey_(record.kodeOP, record.customer);
    lookup[key] = record;

    if (lookup.__conflicts[aliasKey]) return;
    const previous = lookup.__relaxed[aliasKey];
    if (!previous) {
      lookup.__relaxed[aliasKey] = record;
    } else if (previous.toh !== record.toh || previous.email !== record.email || previous.area !== record.area) {
      lookup.__conflicts[aliasKey] = true;
      delete lookup.__relaxed[aliasKey];
    }
  });
  return lookup;
}

function getMissingOperatingPointMappingFields_(mapping) {
  return ADMIN_MAPPING_REQUIRED_FIELDS.filter(function(field) {
    return !cleanText_(mapping && mapping[field.property]);
  }).map(function(field) {
    return field.label;
  });
}

function appendMissingOperatingPointsToMaster_(masterSheet, sourceRows, sourceColumns) {
  const masterLookup = getOperatingPointMappingLookup_(masterSheet.getParent(), masterSheet.getName());
  const seenSourceKeys = {};
  const missing = [];

  sourceRows.forEach(function(row) {
    const kodeOP = cleanText_(row[sourceColumns['kode op']]);
    const namaOP = cleanText_(row[sourceColumns['nama op']]);
    const customer = cleanText_(row[sourceColumns.customer]);
    if (!kodeOP && !namaOP) return;

    const sourceKey = buildOperatingPointKey_(kodeOP, namaOP, customer);
    if (seenSourceKeys[sourceKey]) return;
    seenSourceKeys[sourceKey] = true;
    if (masterLookup.__allExact[sourceKey]) return;

    missing.push({ kodeOP: kodeOP, namaOP: namaOP, customer: customer });
    const placeholder = {
      rowNumber: 0,
      area: '',
      kodeCabang: '',
      namaCabang: '',
      kodeOP: kodeOP,
      namaOP: namaOP,
      customer: customer,
      toh: '',
      email: ''
    };
    masterLookup.__allExact[sourceKey] = placeholder;
  });

  if (!missing.length) return 0;
  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn())
    .getValues()[0].map(cleanHeader_);
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  const rows = missing.map(function(item) {
    const row = new Array(headers.length).fill('');
    row[columns['kode op']] = item.kodeOP;
    row[columns['nama op']] = item.namaOP;
    row[columns.customer] = item.customer;
    return row;
  });
  const firstTargetRow = masterSheet.getLastRow() + 1;
  const requiredLastRow = firstTargetRow + rows.length - 1;
  if (requiredLastRow > masterSheet.getMaxRows()) {
    masterSheet.insertRowsAfter(masterSheet.getMaxRows(), requiredLastRow - masterSheet.getMaxRows());
  }
  masterSheet.getRange(firstTargetRow, 1, rows.length, headers.length).setValues(rows);
  return rows.length;
}

function getAdministrationCoverage_(sourceCode) {
  const definition = ADMIN_MAPPING_SOURCES[sourceCode];
  if (!definition) throw new Error('Sumber mapping tidak valid.');
  const spreadsheet = SpreadsheetApp.openById(definition.spreadsheetId);
  const masterSheet = ensureOperatingPointMaster_(spreadsheet, definition.masterSheet);
  const sheet = spreadsheet.getSheetByName(definition.dataSheet);
  if (!sheet) throw new Error('Sheet "' + definition.dataSheet + '" tidak ditemukan.');
  if (sheet.getLastRow() < 2) return { coverage: [], syncedMissingCount: 0 };

  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(cleanHeader_);
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  ['kode op', 'nama op', 'customer'].forEach(function(header) {
    if (columns[header] === undefined) {
      throw new Error('Kolom "' + header + '" tidak ditemukan di ' + definition.dataSheet + '.');
    }
  });

  const syncedMissingCount = appendMissingOperatingPointsToMaster_(masterSheet, values, columns);
  const lookup = getOperatingPointMappingLookup_(spreadsheet, definition.masterSheet);
  const coverage = {};
  values.forEach(function(row) {
    const kodeOP = cleanText_(row[columns['kode op']]);
    const namaOP = cleanText_(row[columns['nama op']]);
    const customer = cleanText_(row[columns.customer]);
    if (!kodeOP && !namaOP) return;
    const key = buildOperatingPointKey_(kodeOP, namaOP, customer);
    const mapping = findOperatingPointMapping_(lookup, kodeOP, namaOP, customer) || {};
    const aliasKey = buildOperatingPointAliasKey_(kodeOP, customer);
    if (!coverage[key]) {
      coverage[key] = {
        source: sourceCode,
        sourceLabel: definition.label,
        masterRowNumber: mapping.rowNumber || 0,
        area: mapping.area || '',
        kodeCabang: mapping.kodeCabang || '',
        namaCabang: mapping.namaCabang || '',
        kodeOP: kodeOP,
        namaOP: namaOP,
        customer: customer,
        toh: mapping.toh || '',
        email: mapping.email || '',
        origin: mapping.toh ? 'Master_OP_OH' : 'Master belum lengkap',
        conflict: Boolean(lookup.__conflicts && lookup.__conflicts[aliasKey]),
        rowCount: 0
      };
    }
    coverage[key].rowCount++;
  });

  const coverageRows = Object.keys(coverage).map(function(key) {
    const item = coverage[key];
    item.missingFields = getMissingOperatingPointMappingFields_(item);
    item.covered = item.missingFields.length === 0;
    item.areaMapped = Boolean(item.area);
    item.complete = item.covered;
    return item;
  });
  return { coverage: coverageRows, syncedMissingCount: syncedMissingCount };
}

function mergeAdministrationCoverage_(coverageRows) {
  const grouped = {};
  (coverageRows || []).forEach(function(item) {
    const key = buildOperatingPointKey_(item.kodeOP, item.namaOP, item.customer);
    if (!grouped[key]) {
      grouped[key] = {
        kodeOP: item.kodeOP,
        namaOP: item.namaOP,
        customer: item.customer,
        sources: [],
        sourceLabels: [],
        masterRowNumbers: {},
        sourceMappings: {},
        rowCount: 0,
        conflict: false
      };
    }
    const group = grouped[key];
    if (group.sources.indexOf(item.source) === -1) {
      group.sources.push(item.source);
      group.sourceLabels.push(item.sourceLabel);
    }
    group.masterRowNumbers[item.source] = item.masterRowNumber || 0;
    group.sourceMappings[item.source] = {
      area: item.area || '',
      kodeCabang: item.kodeCabang || '',
      namaCabang: item.namaCabang || '',
      kodeOP: item.kodeOP || '',
      namaOP: item.namaOP || '',
      customer: item.customer || '',
      toh: item.toh || '',
      email: item.email || '',
      complete: Boolean(item.complete),
      missingFields: item.missingFields || []
    };
    group.rowCount += Number(item.rowCount || 0);
    group.conflict = group.conflict || Boolean(item.conflict);
  });

  return Object.keys(grouped).map(function(key) {
    const group = grouped[key];
    const mappings = group.sources.map(function(sourceCode) {
      return group.sourceMappings[sourceCode];
    });
    const preferred = mappings.slice().sort(function(a, b) {
      const scoreA = ADMIN_MAPPING_REQUIRED_FIELDS.reduce(function(score, field) {
        return score + (cleanText_(a[field.property]) ? 1 : 0);
      }, 0);
      const scoreB = ADMIN_MAPPING_REQUIRED_FIELDS.reduce(function(score, field) {
        return score + (cleanText_(b[field.property]) ? 1 : 0);
      }, 0);
      return scoreB - scoreA;
    })[0] || {};

    ['area', 'kodeCabang', 'namaCabang', 'toh', 'email'].forEach(function(field) {
      const distinct = {};
      mappings.forEach(function(mapping) {
        const value = cleanText_(mapping[field]).toLowerCase();
        if (value) distinct[value] = true;
      });
      if (Object.keys(distinct).length > 1) group.conflict = true;
    });

    group.source = group.sources[0] || '';
    group.sourceLabel = group.sourceLabels.join(' + ');
    group.masterRowNumber = group.sources.length === 1
      ? Number(group.masterRowNumbers[group.source] || 0)
      : 0;
    group.area = preferred.area || '';
    group.kodeCabang = preferred.kodeCabang || '';
    group.namaCabang = preferred.namaCabang || '';
    group.toh = preferred.toh || '';
    group.email = preferred.email || '';
    group.areaMapped = mappings.length > 0 && mappings.every(function(mapping) { return Boolean(mapping.area); });
    group.complete = mappings.length > 0 && mappings.every(function(mapping) { return Boolean(mapping.complete); });
    group.covered = group.complete;
    group.missingFields = Array.from(new Set([].concat.apply([], mappings.map(function(mapping) {
      return mapping.missingFields || [];
    }))));
    group.origin = group.complete ? 'Master_OP_OH' : 'Master belum lengkap';
    return group;
  });
}

function ensureAdminConfigSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(ADMIN_CONFIG_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(ADMIN_CONFIG_SHEET);
  ensureHeaders_(sheet, ADMIN_CONFIG_HEADERS);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 360);
  sheet.setColumnWidth(4, 110);
  return sheet;
}

function getAdministrationConfigs_() {
  const sheet = ensureAdminConfigSheet_();
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(cleanHeader_);
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  return values.map(function(row) {
    return {
      key: cleanText_(row[columns['config_key']]),
      value: cleanText_(row[columns['config_value']]),
      description: cleanText_(row[columns.description]),
      status: cleanText_(row[columns.status]) || 'ACTIVE',
      updatedAt: serializeForClient_(row[columns['updated_at']]),
      updatedBy: cleanText_(row[columns['updated_by']])
    };
  }).filter(function(item) { return item.key; });
}

function getDashboardSettings_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(ADMIN_CONFIG_SHEET);
  const stored = {};
  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getDataRange().getValues();
    const headers = values.shift().map(cleanHeader_);
    const columns = {};
    headers.forEach(function(header, index) { columns[header] = index; });
    values.forEach(function(row) {
      const key = cleanText_(row[columns['config_key']]).toUpperCase();
      const status = cleanText_(row[columns.status]).toUpperCase() || 'ACTIVE';
      if (key && status === 'ACTIVE') stored[key] = cleanText_(row[columns['config_value']]);
    });
  }

  const settings = {};
  ADMIN_SETTING_DEFINITIONS.forEach(function(definition) {
    settings[definition.key] = Object.prototype.hasOwnProperty.call(stored, definition.key)
      ? stored[definition.key]
      : definition.defaultValue;
  });
  return settings;
}

function buildAdministrationAreaSummary_(coverage) {
  const areas = {};
  coverage.forEach(function(item) {
    const area = item.area || 'BELUM TERPETAKAN';
    if (!areas[area]) {
      areas[area] = { area: area, total: 0, complete: 0, uncovered: 0, conflicts: 0, rowCount: 0 };
    }
    const summary = areas[area];
    summary.total++;
    summary.rowCount += Number(item.rowCount || 0);
    if (item.complete) summary.complete++;
    if (!item.complete) summary.uncovered++;
    if (item.conflict) summary.conflicts++;
  });
  return Object.keys(areas).map(function(key) {
    const item = areas[key];
    item.percentage = item.total ? Math.round(item.complete / item.total * 100) : 0;
    return item;
  }).sort(function(a, b) {
    if (a.area === 'BELUM TERPETAKAN') return -1;
    if (b.area === 'BELUM TERPETAKAN') return 1;
    return a.area.localeCompare(b.area, 'id');
  });
}

function buildAdministrationOwners_(coverage) {
  const owners = {};
  coverage.forEach(function(item) {
    if (!item.toh) return;
    const key = normalizeMappingMatchText_(item.toh) + '|' + normalizeEmail_(item.email);
    if (!owners[key]) {
      owners[key] = { name: item.toh, email: item.email || '', areas: {}, mappingCount: 0 };
    }
    owners[key].mappingCount++;
    if (item.area) owners[key].areas[item.area] = true;
  });
  return Object.keys(owners).map(function(key) {
    const owner = owners[key];
    owner.areas = Object.keys(owner.areas).sort();
    return owner;
  }).sort(function(a, b) {
    return a.name.localeCompare(b.name, 'id');
  });
}

function getAdministrationData() {
  assertAdmin_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sourceCoverage = [];
    let syncedMissingCount = 0;
    Object.keys(ADMIN_MAPPING_SOURCES).forEach(function(sourceCode) {
      const sourceResult = getAdministrationCoverage_(sourceCode);
      Array.prototype.push.apply(sourceCoverage, sourceResult.coverage);
      syncedMissingCount += sourceResult.syncedMissingCount;
    });
    const coverage = mergeAdministrationCoverage_(sourceCoverage);
    coverage.sort(function(a, b) {
      if (a.complete !== b.complete) return a.complete ? 1 : -1;
      return (a.sourceLabel + a.kodeOP + a.namaOP).localeCompare(
        b.sourceLabel + b.kodeOP + b.namaOP,
        'id',
        { numeric: true }
      );
    });
    return {
      success: true,
      coverage: coverage,
      areaSummary: buildAdministrationAreaSummary_(coverage),
      owners: buildAdministrationOwners_(coverage),
      configs: getAdministrationConfigs_(),
      settings: getDashboardSettings_(),
      settingDefinitions: ADMIN_SETTING_DEFINITIONS,
      syncedMissingCount: syncedMissingCount,
      summary: {
        total: coverage.length,
        covered: coverage.filter(function(item) { return item.covered; }).length,
        complete: coverage.filter(function(item) { return item.complete; }).length,
        uncovered: coverage.filter(function(item) { return !item.complete; }).length,
        conflicts: coverage.filter(function(item) { return item.conflict; }).length
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function validateAdministrationMapping_(payload) {
  const source = payload && typeof payload === 'object' ? payload : {};
  const requestedSources = Array.isArray(source.sources) && source.sources.length
    ? source.sources
    : [source.source];
  const sourceCodes = [];
  requestedSources.forEach(function(value) {
    const sourceCode = cleanText_(value).toUpperCase();
    if (!ADMIN_MAPPING_SOURCES[sourceCode]) throw new Error('Sumber mapping tidak valid.');
    if (sourceCodes.indexOf(sourceCode) === -1) sourceCodes.push(sourceCode);
  });
  if (!sourceCodes.length) throw new Error('Sumber mapping tidak valid.');
  const suppliedRowNumbers = source.masterRowNumbers && typeof source.masterRowNumbers === 'object'
    ? source.masterRowNumbers
    : {};
  if (source.source && source.masterRowNumber && !suppliedRowNumbers[source.source]) {
    suppliedRowNumbers[source.source] = source.masterRowNumber;
  }
  const masterRowNumbers = {};
  sourceCodes.forEach(function(sourceCode) {
    masterRowNumbers[sourceCode] = Math.max(
      Math.floor(Number(suppliedRowNumbers[sourceCode]) || 0),
      0
    );
  });
  const area = cleanText_(source.area).toUpperCase();
  const kodeCabang = cleanText_(source.kodeCabang).toUpperCase();
  const namaCabang = cleanText_(source.namaCabang);
  const kodeOP = cleanText_(source.kodeOP);
  const namaOP = cleanText_(source.namaOP);
  const customer = cleanText_(source.customer);
  const toh = cleanText_(source.toh);
  const email = normalizeEmail_(source.email);
  if (!area) throw new Error('Area wajib diisi.');
  if (!kodeCabang) throw new Error('Kode Cabang wajib diisi.');
  if (!namaCabang) throw new Error('Nama Cabang wajib diisi.');
  if (!kodeOP) throw new Error('Kode OP wajib diisi.');
  if (!namaOP) throw new Error('Nama OP wajib diisi.');
  if (!customer) throw new Error('Customer wajib diisi.');
  if (!toh) throw new Error('Nama TOH/WOH wajib diisi.');
  if (!email) throw new Error('Email TOH wajib diisi.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Format email TOH tidak valid.');
  }
  return {
    sources: sourceCodes,
    masterRowNumbers: masterRowNumbers,
    area: area,
    kodeCabang: kodeCabang,
    namaCabang: namaCabang,
    kodeOP: kodeOP,
    namaOP: namaOP,
    customer: customer,
    toh: toh,
    email: email
  };
}

function expandAdministrationMappingSources_(mapping) {
  return mapping.sources.map(function(sourceCode) {
    return {
      source: sourceCode,
      masterRowNumber: Number(mapping.masterRowNumbers[sourceCode] || 0),
      area: mapping.area,
      kodeCabang: mapping.kodeCabang,
      namaCabang: mapping.namaCabang,
      kodeOP: mapping.kodeOP,
      namaOP: mapping.namaOP,
      customer: mapping.customer,
      toh: mapping.toh,
      email: mapping.email
    };
  });
}

function upsertAdministrationMapping_(mapping, resources) {
  const definition = ADMIN_MAPPING_SOURCES[mapping.source];
  let resource = resources[mapping.source];
  if (!resource) {
    const spreadsheet = SpreadsheetApp.openById(definition.spreadsheetId);
    const sheet = ensureOperatingPointMaster_(spreadsheet, definition.masterSheet);
    const values = sheet.getDataRange().getValues();
    const headers = values.shift().map(cleanHeader_);
    const columns = {};
    headers.forEach(function(header, index) { columns[header] = index; });
    resource = { spreadsheet: spreadsheet, sheet: sheet, values: values, headers: headers, columns: columns };
    resources[mapping.source] = resource;
  }

  const sheet = resource.sheet;
  const values = resource.values;
  const headers = resource.headers;
  const columns = resource.columns;
  const targetKey = buildOperatingPointKey_(mapping.kodeOP, mapping.namaOP, mapping.customer);
  let targetRow = sheet.getLastRow() + 1;
  if (mapping.masterRowNumber >= 2 && mapping.masterRowNumber <= sheet.getLastRow()) {
    const candidate = values[mapping.masterRowNumber - 2];
    const candidateKey = buildOperatingPointKey_(
      candidate[columns['kode op']],
      candidate[columns['nama op']],
      candidate[columns.customer]
    );
    if (candidateKey === targetKey) {
      targetRow = mapping.masterRowNumber;
    }
  }

  if (targetRow > sheet.getLastRow()) {
    values.some(function(row, index) {
      const key = buildOperatingPointKey_(
        row[columns['kode op']],
        row[columns['nama op']],
        row[columns.customer]
      );
      if (key === targetKey) {
        targetRow = index + 2;
        return true;
      }
      return false;
    });
  }

  const existingRow = targetRow <= sheet.getLastRow() ? values[targetRow - 2] : null;
  const row = existingRow ? existingRow.slice() : new Array(headers.length).fill('');
  while (row.length < headers.length) row.push('');
  if (columns.area !== undefined) row[columns.area] = mapping.area;
  if (columns['kode cabang'] !== undefined) row[columns['kode cabang']] = mapping.kodeCabang;
  if (columns['nama cabang'] !== undefined) row[columns['nama cabang']] = mapping.namaCabang;
  row[columns['kode op']] = mapping.kodeOP;
  row[columns['nama op']] = mapping.namaOP;
  row[columns.customer] = mapping.customer;
  row[columns['nama toh/woh']] = mapping.toh;
  row[columns['email toh']] = mapping.email;
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
  if (existingRow) values[targetRow - 2] = row;
  else values.push(row);
  return targetRow;
}

function invalidateAdministrationSourceCaches_(sources) {
  if (sources.REKON) {
    touchMonitoringIndex_();
    updateDashboardTimestamp_();
  }
  if (sources.SPJ) {
    PropertiesService.getScriptProperties().setProperty(SPJ_CACHE_EPOCH_PROPERTY, String(Date.now()));
    CacheService.getScriptCache().remove(SPJ_UPDATE_TOKEN_CACHE_KEY);
  }
}

function saveAdministrationMapping(payload) {
  assertAdmin_();
  const mapping = validateAdministrationMapping_(payload);
  const sourceMappings = expandAdministrationMappingSources_(mapping);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const resources = {};
    const touched = {};
    const masterRowNumbers = {};
    sourceMappings.forEach(function(sourceMapping) {
      touched[sourceMapping.source] = true;
      masterRowNumbers[sourceMapping.source] = upsertAdministrationMapping_(sourceMapping, resources);
    });
    invalidateAdministrationSourceCaches_(touched);
    return {
      success: true,
      masterRowNumbers: masterRowNumbers,
      message: 'Mapping Master_OP_OH berhasil disimpan untuk ' + sourceMappings.length + ' sumber.'
    };
  } finally {
    lock.releaseLock();
  }
}

function saveAdministrationMappingsBulk(payloads) {
  assertAdmin_();
  if (!Array.isArray(payloads) || !payloads.length) throw new Error('Pilih minimal satu Operating Point.');
  if (payloads.length > 100) throw new Error('Maksimal 100 mapping dalam satu proses.');
  const mappings = payloads.map(validateAdministrationMapping_);
  const expandedMappings = [];
  mappings.forEach(function(mapping) {
    Array.prototype.push.apply(expandedMappings, expandAdministrationMappingSources_(mapping));
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const resources = {};
    const touched = {};
    const rowNumbers = expandedMappings.map(function(mapping) {
      touched[mapping.source] = true;
      return upsertAdministrationMapping_(mapping, resources);
    });
    invalidateAdministrationSourceCaches_(touched);
    return {
      success: true,
      masterRowNumbers: rowNumbers,
      message: mappings.length + ' mapping unik berhasil disimpan ke ' + expandedMappings.length + ' master sumber.'
    };
  } finally {
    lock.releaseLock();
  }
}

function validateAdministrationConfig_(payload) {
  const source = payload && typeof payload === 'object' ? payload : {};
  const key = cleanText_(source.key).toUpperCase().replace(/\s+/g, '_');
  const value = String(source.value == null ? '' : source.value).trim();
  const description = cleanText_(source.description);
  const status = cleanText_(source.status).toUpperCase() || 'ACTIVE';
  if (!/^[A-Z][A-Z0-9_]{2,79}$/.test(key)) {
    throw new Error('Config key harus 3-80 karakter dan hanya berisi huruf, angka, atau underscore.');
  }
  if (value.length > 1000) throw new Error('Config value maksimal 1000 karakter.');
  if (description.length > 500) throw new Error('Deskripsi maksimal 500 karakter.');
  if (['ACTIVE', 'INACTIVE'].indexOf(status) === -1) throw new Error('Status config tidak valid.');
  return { key: key, value: value, description: description, status: status };
}

function saveAdministrationConfig(payload) {
  const admin = assertAdmin_();
  const config = validateAdministrationConfig_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureAdminConfigSheet_();
    const values = sheet.getDataRange().getValues();
    const headers = values.shift().map(cleanHeader_);
    const columns = {};
    headers.forEach(function(header, index) { columns[header] = index; });
    let targetRow = sheet.getLastRow() + 1;
    values.forEach(function(row, index) {
      if (cleanText_(row[columns['config_key']]).toUpperCase() === config.key) targetRow = index + 2;
    });
    const row = new Array(headers.length).fill('');
    row[columns['config_key']] = config.key;
    row[columns['config_value']] = config.value;
    row[columns.description] = config.description;
    row[columns.status] = config.status;
    row[columns['updated_at']] = new Date();
    row[columns['updated_by']] = admin.email;
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
    return { success: true, message: 'Konfigurasi berhasil disimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function validateDashboardSettingValue_(definition, value) {
  const text = String(value == null ? '' : value).trim();
  if (definition.type === 'number') {
    const number = Number(text);
    if (!Number.isFinite(number) || number < definition.min || number > definition.max) {
      throw new Error(definition.label + ' harus di antara ' + definition.min + ' dan ' + definition.max + '.');
    }
    return String(Math.round(number));
  }
  if (definition.type === 'select' && definition.options.indexOf(text) === -1) {
    throw new Error('Pilihan untuk ' + definition.label + ' tidak valid.');
  }
  if (definition.type === 'url' && text && !/^https:\/\//i.test(text)) {
    throw new Error(definition.label + ' harus menggunakan URL HTTPS.');
  }
  if (definition.maxLength && text.length > definition.maxLength) {
    throw new Error(definition.label + ' maksimal ' + definition.maxLength + ' karakter.');
  }
  return text;
}

function saveAdministrationSettings(payload) {
  const admin = assertAdmin_();
  const source = payload && typeof payload === 'object' ? payload : {};
  const validated = {};
  ADMIN_SETTING_DEFINITIONS.forEach(function(definition) {
    const rawValue = Object.prototype.hasOwnProperty.call(source, definition.key)
      ? source[definition.key]
      : definition.defaultValue;
    validated[definition.key] = validateDashboardSettingValue_(definition, rawValue);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureAdminConfigSheet_();
    const values = sheet.getDataRange().getValues();
    const headers = values.shift().map(cleanHeader_);
    const columns = {};
    headers.forEach(function(header, index) { columns[header] = index; });
    const rowByKey = {};
    values.forEach(function(row, index) {
      const key = cleanText_(row[columns['config_key']]).toUpperCase();
      if (key) rowByKey[key] = index + 2;
    });

    ADMIN_SETTING_DEFINITIONS.forEach(function(definition) {
      const targetRow = rowByKey[definition.key] || sheet.getLastRow() + 1;
      const row = new Array(headers.length).fill('');
      row[columns['config_key']] = definition.key;
      row[columns['config_value']] = validated[definition.key];
      row[columns.description] = definition.label;
      row[columns.status] = 'ACTIVE';
      row[columns['updated_at']] = new Date();
      row[columns['updated_by']] = admin.email;
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
      rowByKey[definition.key] = targetRow;
    });
    return { success: true, settings: validated, message: 'Pengaturan dashboard berhasil disimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function resetAdministrationCaches() {
  assertAdmin_();
  const cache = CacheService.getScriptCache();
  cache.remove(MONITORING_UPDATE_TOKEN_CACHE_KEY);
  cache.remove(SPJ_UPDATE_TOKEN_CACHE_KEY);
  clearAccessRegistryCache_();
  touchMonitoringIndex_();
  PropertiesService.getScriptProperties().setProperty(SPJ_CACHE_EPOCH_PROPERTY, String(Date.now()));
  return { success: true, message: 'Cache dashboard berhasil direset. Data terbaru akan dimuat pada request berikutnya.' };
}

function isAdminEmail_(email) {
  return normalizeEmail_(email) === normalizeEmail_(ACCESS_ADMIN_EMAIL);
}

function assertAdmin_() {
  const email = getActiveEmail_();
  if (!email || !isAdminEmail_(email)) {
    throw new Error('ACCESS_DENIED: Fungsi ini hanya dapat dijalankan oleh admin.');
  }
  return { allowed: true, email: email, accessType: 'admin' };
}