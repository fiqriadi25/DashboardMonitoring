const SPREADSHEET_ID = "1d5eFnUM65lzpaJUJXSdpZG0BmK70SCf3TF5K4tljCZw";
const SHEET_NAME = "Database";

/* ================= WEB APP ================= */

function doGet() {
  // ========== KONFIGURASI AKSES ==========
  const allowedEmails = [
    'admin.bitrans@b-log.co.id',
    'arsyal.aswindi1996@gmail.com',
    'hherdi123@gmail.com',
    'ginginanjar683@gmail.com',
  ];
  const allowedDomains = ['b-log.co.id', 'b-log.id'];
  
  // Nomor WhatsApp admin (gunakan kode negara tanpa tanda +, misal 628123456789)
  const ADMIN_WHATSAPP_NUMBER = '6285156451872'; // <-- GANTI DENGAN NOMOR ANDA
  
  // ========== CEK EMAIL USER ==========
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch(e) {
    return HtmlService.createHtmlOutput(`
      <h1>Akses Ditolak</h1>
      <p>Anda harus login dengan akun Google untuk mengakses halaman ini.</p>
    `);
  }
  
  const isDomainAllowed = allowedDomains.some(domain => userEmail.endsWith('@' + domain));
  const isEmailAllowed = allowedEmails.includes(userEmail);
  
  if (!isDomainAllowed && !isEmailAllowed) {
    // Membuat link WhatsApp
    const whatsappLink = `https://wa.me/${ADMIN_WHATSAPP_NUMBER}?text=Halo%20Team%2C%20Saya%20${encodeURIComponent(userEmail)}%20Tidak%20Bisa%20Mengakses%20Dashboard%20Digital%20Enablement.%20Mohon%20Bantuannya.`;
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Akses Ditolak</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background: #f1f5f9; }
          .card { background: white; border-radius: 20px; padding: 30px; max-width: 400px; margin: auto; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); }
          h1 { color: #dc2626; }
          p { color: #334155; }
          .wa-link {
            display: inline-block;
            background-color: #25d366;
            color: white;
            padding: 10px 20px;
            border-radius: 30px;
            text-decoration: none;
            font-weight: bold;
            margin-top: 20px;
          }
          .wa-link:hover {
            background-color: #128c7e;
          }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>⛔ Akses Ditolak</h1>
          <p>Anda tidak memiliki izin untuk mengakses dashboard ini.</p>
          <p>Email Anda: <strong>${userEmail}</strong></p>
          <p>Silakan hubungi administrator melalui WhatsApp dengan klik tombol di bawah:</p>
          <a href="${whatsappLink}" target="_blank" class="wa-link">💬 Hubungi Administrator via WhatsApp</a>
        </div>
      </body>
      </html>
    `);
  }
  
  // ========== JIKA LOLOS, TAMPILKAN DASHBOARD ==========
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("Dashboard Digital Enablement")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================= HELPER ================= */

function cleanHeader(h) {
  return String(h).trim().toLowerCase();
}

function cleanText(v) {
  if (!v) return "";
  return String(v).replace(/\s+/g, ' ').trim();
}

/* ================= UNIVERSAL DATE PARSER ================= */

function parseDate(val) {
  if (!val) return null;

  if (Object.prototype.toString.call(val) === "[object Date]")
    return val;

  // dd/mm/yyyy
  if (typeof val === "string" && val.includes("/")) {
    const p = val.split("/");
    if (p.length === 3)
      return new Date(p[2], p[1] - 1, p[0]);
  }

  // yyyy-mm-dd
  if (typeof val === "string" && val.includes("-")) {
    const d = new Date(val);
    if (!isNaN(d)) return d;
  }

  // excel serial
  if (typeof val === "number") {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }

  return null;
}

/* ================= MAIN FUNCTION ================= */

function getMonitoringData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" tidak ditemukan`);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1)
      return { monitoring: [], lastUpdate: "-" };

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const rawHeader = values.shift();
    const header = rawHeader.map(cleanHeader);

    const colOP         = header.indexOf("kode op");
    const colNamaOP     = header.indexOf("nama op");
    const colCustomer   = header.indexOf("customer");
    const colProduk     = header.indexOf("produk");
    const colTOH        = header.indexOf("nama toh/woh");
    const colDate       = header.indexOf("tanggal sla");
    const colStatus     = header.indexOf("status rekon");
    const colTanggalUpdate = header.indexOf("tanggal update");

    if (colOP === -1 || colNamaOP === -1 || colCustomer === -1 || colDate === -1 || colStatus === -1) {
      throw new Error("Header wajib tidak ditemukan. Pastikan ada: Kode OP, Nama OP, Customer, Produk, Tanggal SLA, Status Rekon");
    }

    let data = [];
    let lastUpdate = null;

    values.forEach(r => {
      const op       = cleanText(r[colOP]);
      const namaop   = cleanText(r[colNamaOP]);
      const customer = cleanText(r[colCustomer]);
      const produk   = colProduk > -1 ? cleanText(r[colProduk]) : "-";

      if (!op || !r[colDate]) return;

      const d = parseDate(r[colDate]);
      if (!d) return;

      let status = cleanText(r[colStatus]);
      if (!status) status = "Non Rekon";

      const monthKey = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");

      data.push({
        op:      op,
        namaop:  namaop,
        customer: customer,
        produk:  produk,
        toh:     (colTOH > -1 && r[colTOH]) ? cleanText(r[colTOH]) : "-",
        month:   monthKey,
        day:     String(d.getDate()),
        status:  status
      });

      if (colTanggalUpdate > -1 && r[colTanggalUpdate]) {
        const u = parseDate(r[colTanggalUpdate]);
        if (u && (!lastUpdate || u > lastUpdate))
          lastUpdate = u;
      }
    });

    return {
      monitoring: data,
      lastUpdate: lastUpdate
        ? Utilities.formatDate(lastUpdate, Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
        : "-"
    };

  } catch (err) {
    Logger.log(err);
    return {
      error: err.toString(),
      monitoring: [],
      lastUpdate: "-"
    };
  }
}

/* ================= AUTO SEND EMAIL TOH/WOH ================= */

function sendDailyNonRekonEmail() {

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
  const logSheet = ss.getSheetByName("Email_Log") || ss.insertSheet("Email_Log");

  if (!db)     throw new Error("Sheet Database tidak ditemukan");
  if (!master) throw new Error("Sheet Master_OP_OH tidak ditemukan");

  const dbData   = db.getDataRange().getValues();
  const dbHeader = dbData.shift();

  const masterData   = master.getDataRange().getValues();
  const masterHeader = masterData.shift();

  const colOP       = dbHeader.indexOf("Kode OP");
  const colNamaOP   = dbHeader.indexOf("Nama OP");
  const colCustomer = dbHeader.indexOf("Customer");
  const colProduk   = dbHeader.indexOf("Produk");
  const colTanggal  = dbHeader.indexOf("Tanggal SLA");
  const colStatus   = dbHeader.indexOf("Status Rekon");
  const colTOH      = dbHeader.indexOf("Nama TOH/WOH");

  const colMasterTOH = masterHeader.indexOf("Nama TOH/WOH");
  const colEmail     = masterHeader.indexOf("Email TOH");

  if (colEmail === -1)
    throw new Error("Kolom 'Email TOH' belum ada di Master_OP_OH");

  const timezone = Session.getScriptTimeZone();

  /* 1. EMAIL MAPPING */
  let emailMap = {};
  masterData.forEach(r => {
    if (r[colMasterTOH] && r[colEmail]) {
      const name = r[colMasterTOH].toString().trim().toLowerCase();
      emailMap[name] = r[colEmail].toString().trim();
    }
  });

  /* 2. CARI SLA TERAKHIR */
  let lastSLA = null;
  dbData.forEach(r => {
    if (!r[colTanggal]) return;
    let d = new Date(r[colTanggal]);
    if (isNaN(d)) return;
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
    let tgl = new Date(r[colTanggal]);
    if (isNaN(tgl)) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "yyyy-MM");
    if (!monthsToCheck.includes(monthKey)) return;
    if (String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    const toh = String(r[colTOH]).trim();
    if (!toh) return;

    const key = r[colOP] + "||" + r[colCustomer] + "||" + r[colProduk];
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
            <span style="background:linear-gradient(135deg,#667eea15,#764ba215);padding:4px 10px;border-radius:6px;font-size:12px;font-weight:600;color:#1e293b;border-left:2px solid #6366f1;">${d.op}</span>
          </td>
          <td style="padding:12px;">
            <div style="font-weight:600;color:#0f172a;font-size:13px;">${d.nama}</div>
          </td>
          <td style="padding:12px;">
            <span style="background:#e0f2fe;color:#0369a1;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:500;">${d.cust}</span>
          </td>
          <td style="padding:12px;color:#334155;font-size:12px;">${d.produk}</td>
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
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
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
      <h2 style="color:#0f172a;margin:0 0 3px;font-size:15px;font-weight:600;">Yth. <span style="color:#6366f1;">${toh}</span></h2>
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
        MailApp.sendEmail({
          to:       email,
          subject:  "⚠ [BITRANS] Non-Rekon " + periodeBulan + " • " + priority,
          htmlBody: html
        });
        totalEmailSent++;
        logSheet.appendRow([new Date(), periodeBulan, toh, email, jumlahKasus, totalTanggal, priority, "SUCCESS"]);
      } catch (err) {
        logSheet.appendRow([new Date(), periodeBulan, toh, email, jumlahKasus, totalTanggal, priority, "FAILED : " + err.message]);
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
  if (!selectedMonth) throw new Error("Bulan tidak dipilih.");

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const db     = ss.getSheetByName("Database");
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

  const timezone = Session.getScriptTimeZone();

  function normalizeMonth(m) {
    const parts = String(m).split("-");
    if (parts.length !== 2) return m;
    return parts[0] + "-" + parts[1].padStart(2, "0");
  }

  selectedMonth = normalizeMonth(selectedMonth);

  let output = [["Kode OP", "Nama OP", "Customer", "Produk", "TOH", "Tanggal SLA", "Status"]];

  data.forEach(r => {
    if (!r[colTanggal]) return;
    let tgl = new Date(r[colTanggal]);
    if (isNaN(tgl)) return;

    const monthKey = Utilities.formatDate(tgl, timezone, "yyyy-MM");
    if (monthKey !== selectedMonth) return;
    if (String(r[colStatus]).trim().toLowerCase() === "rekon") return;

    output.push([
      r[colOP], r[colNamaOP], r[colCustomer], r[colProduk], r[colTOH],
      Utilities.formatDate(tgl, timezone, "d-MMMM-yyyy"),
      r[colStatus]
    ]);
  });

  if (output.length <= 1) throw new Error("Tidak ada data Non-Rekon pada bulan tersebut.");

  const tempSS  = SpreadsheetApp.create("TEMP_EXPORT");
  const sheet   = tempSS.getActiveSheet();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange("A1:G1").setFontWeight("bold");
  sheet.autoResizeColumns(1, 7);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  const fileId = tempSS.getId();
  const url    = "https://www.googleapis.com/drive/v3/files/" + fileId
                 + "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const token  = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(fileId).setTrashed(true);
    throw new Error("Gagal export XLSX. Pastikan Drive API aktif.");
  }

  const blob = response.getBlob().setName("Rekap_Non_Rekon_" + selectedMonth + ".xlsx");
  DriveApp.getFileById(fileId).setTrashed(true);

  return {
    fileName: blob.getName(),
    mimeType: blob.getContentType(),
    data:     Utilities.base64Encode(blob.getBytes())
  };
}

/* ================= AUTO REFRESH MONITORING DASHBOARD ================= */

function updateDashboardTimestamp() {
  const now = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty("LAST_UPDATE", now);
  Logger.log("Dashboard Updated : " + now);
}

function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const editedColumn    = e.range.getColumn();
    const monitoredColumns = [4, 5, 6, 7];
    if (!monitoredColumns.includes(editedColumn)) return;

    const props = PropertiesService.getScriptProperties();
    const now   = Date.now();
    const last  = Number(props.getProperty("LAST_UPDATE_TS")) || 0;

    if (now - last < 5000) { Logger.log("Update dilewati (delay filter)"); return; }
    props.setProperty("LAST_UPDATE_TS", now);
    updateDashboardTimestamp();
  } catch (err) {
    Logger.log(err);
  }
}

function onChange(e) {
  try {
    if (!e) return;
    const props = PropertiesService.getScriptProperties();
    const now   = Date.now();
    const last  = Number(props.getProperty("LAST_UPDATE_TS")) || 0;

    if (now - last < 5000) { Logger.log("Change trigger dilewati (delay)"); return; }
    props.setProperty("LAST_UPDATE_TS", now);
    updateDashboardTimestamp();
  } catch (err) {
    Logger.log(err);
  }
}

function getLastUpdate() {
  return PropertiesService.getScriptProperties().getProperty("LAST_UPDATE") || null;
}

/* =========================================================
   DASHBOARD 2 — Tracking SPJ
   ========================================================= */

const SPREADSHEET_ID_SPJ = "12FcleafDZHI-OkyxWEfFQM4BXcx2RBNIfG2ShnoptGo";
const SHEET_DATA          = "Data_Rekon";
const SHEET_MASTER_SPJ    = "Master_OP_OH";
const SHEET_CONFIG_SPJ    = "Config";

// NOTE: Renamed from getData() → getSPJData() to avoid name collision with
// the client-side getData() filter function in index.html
function getSPJData() {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID_SPJ);
    const dataSheet = ss.getSheetByName(SHEET_DATA);
    if (!dataSheet) throw new Error(`Sheet "${SHEET_DATA}" tidak ditemukan.`);

    const dataRows = dataSheet.getDataRange().getValues();
    if (dataRows.length < 2) {
      return {
        success:       true,
        data:          [],
        filters:       { kodeOP: [], customer: [], produk: [], toh: [], adjustment: [] },
        tanggalColumns: [],
        headers:       dataRows[0] || [],
        tanggalUpdate: getSPJLastUpdate(ss)
      };
    }

    const headers = dataRows[0].map(h => String(h).trim());
    const data    = [];

    for (let i = 1; i < dataRows.length; i++) {
      const row     = dataRows[i];
      const rowData = {};
      headers.forEach((header, index) => { rowData[header] = row[index]; });

      let statusAdj = String(rowData["Status Adjustment"] || "").trim();
      if (statusAdj.toLowerCase() === "reject") statusAdj = "No Adjustment";
      rowData["Status Adjustment"] = statusAdj;
      data.push(rowData);
    }

    const FIXED_COLUMNS = [
      "Kode OP", "Nama OP", "Customer", "SPJ", "Produk",
      "Tanggal SLA", "UJP", "Charging", "Status Rekon",
      "Nama TOH/WOH", "Status Adjustment"
    ];
    const tanggalColumns = headers.filter(h => !FIXED_COLUMNS.includes(h));

    // Lookup Nama TOH/WOH dari Master_OP_OH
    const masterSheet = ss.getSheetByName(SHEET_MASTER_SPJ);
    if (masterSheet) {
      const masterRows    = masterSheet.getDataRange().getValues();
      const masterHeaders = masterRows.shift().map(h => String(h).trim());
      const lookup        = {};

      masterRows.forEach(row => {
        const rd  = {};
        masterHeaders.forEach((h, i) => { rd[h] = row[i]; });
        const key  = `${String(rd['Kode OP'] || '').trim()}|${String(rd['Nama OP'] || '').trim()}|${String(rd['Customer'] || '').trim()}`.toLowerCase();
        const nama = rd['Nama TOH/WOH'] || rd['TOH'] || rd['Nama TOH'] || '';
        lookup[key] = nama;
      });

      data.forEach(item => {
        const key = `${String(item['Kode OP'] || '').trim()}|${String(item['Nama OP'] || '').trim()}|${String(item['Customer'] || '').trim()}`.toLowerCase();
        if (lookup[key]) item['Nama TOH/WOH'] = lookup[key];
      });
    }

    const uniqueOP         = [...new Set(data.map(d => d['Kode OP']).filter(Boolean))];
    const uniqueCustomer   = [...new Set(data.map(d => d['Customer']).filter(Boolean))];
    const uniqueProduk     = [...new Set(data.map(d => d['Produk']).filter(Boolean))];
    const uniqueTOH        = [...new Set(data.map(d => d['Nama TOH/WOH']).filter(Boolean))];
    const uniqueAdjustment = [...new Set(data.map(d => d['Status Adjustment']).filter(Boolean))];

    return {
      success:        true,
      data:           data,
      filters:        { kodeOP: uniqueOP, customer: uniqueCustomer, produk: uniqueProduk, toh: uniqueTOH, adjustment: uniqueAdjustment },
      tanggalColumns: tanggalColumns,
      headers:        headers,
      tanggalUpdate:  getSPJLastUpdate(ss)
    };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.toString() };
  }
}

function getSPJLastUpdate(ss) {
  try {
    const configSheet = ss.getSheetByName(SHEET_CONFIG_SPJ);
    if (!configSheet) return 'Belum ada update';
    const val = configSheet.getRange('A1').getValue();
    return val || 'Belum ada update';
  } catch (e) {
    return 'Error baca update';
  }
}
