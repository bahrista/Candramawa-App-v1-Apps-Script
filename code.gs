/**
 * CANDRAMAWA APPS v1.0 — Backend Apps Script (Code.gs)
 * Owner: Syaeful Bahri
 * Integrated with CacheService, Custom Date Parsing, Brief Link, Smart Split Logic,
 * dan Tab Home (daftar brand aktif + link spreadsheet)
 */

var SOURCES_SHEET_NAME = 'Brand Sources';
var DEFAULT_CONTENT_PLAN_SHEET = 'Content Plan';
var CONTENT_HEADER_ROW = 4;
var CONTENT_FIRST_DATA_ROW = 5;
var SOURCES_FIRST_ROW = 6;
var DEFAULT_TITLE = 'Candramawa Apps';
var CACHE_TTL_SECONDS = 600; // Cache bertahan 10 menit (600 detik)

var USER_SHEET_NAME = 'User Management';
var USER_HEADER_ROW = 3;
var USER_FIRST_ROW = 4;
var SESSION_TTL_SECONDS = 21600; // Sesi login bertahan 6 jam (batas maksimum CacheService)

var FIELD_HEADERS = {
  publishDate:  'Publish Date',
  status:       'Status',
  title:        'Title',
  pilar:        'Content Pilar',
  format:       'Format',
  channel:      'Channel',
  reference:    'Reference',           // Link reference konten, kalau ada
  brief:        'Brief',               // Disingkat dari 'Brief (Link)'
  owner:        'Content Owner',
  multimediaCreator: 'Multimedia Creator', // Menggantikan 'Multimedia by'
  caption:      'Caption',             // Field baru
  multimedia:   'Multimedia Asset'     // Field baru (Folder Drive/Link Gambar)
};

// =====================================================================
// WEB APP ENTRY POINT & INCLUDE
// =====================================================================

function doGet(e) {
  var title = getDashboardTitle_();
  var template = HtmlService.createTemplateFromFile('index');
  template.dashboardTitle = title;

  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDashboardTitle_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCES_SHEET_NAME);
  if (!sheet) return DEFAULT_TITLE;
  var lastRow = Math.min(sheet.getLastRow(), 4);
  if (lastRow < 1) return DEFAULT_TITLE;
  var values = sheet.getRange(1, 1, lastRow, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim().toLowerCase();
    if (key.indexOf('nama dashboard') > -1 && values[i][1]) {
      return String(values[i][1]).trim();
    }
  }
  return DEFAULT_TITLE;
}

// =====================================================================
// HELPERS
// =====================================================================

function extractSpreadsheetId_(input) {
  input = String(input || '').trim();
  if (!input) return '';
  var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  return input;
}

/**
 * Membentuk URL spreadsheet yang bisa diklik untuk tabel di tab Home.
 * Kalau isi kolom B sudah berupa URL penuh, dipakai apa adanya.
 * Kalau isinya cuma Spreadsheet ID, dibungkus jadi URL lengkap.
 */
function buildSpreadsheetUrl_(idOrUrl) {
  var str = String(idOrUrl || '').trim();
  if (!str) return '';
  if (str.indexOf('http://') === 0 || str.indexOf('https://') === 0) return str;
  return 'https://docs.google.com/spreadsheets/d/' + str + '/edit';
}

function buildColIndex_(headerRowValues) {
  var colIndex = {};
  headerRowValues.forEach(function (headerText, i) {
    var norm = String(headerText || '').trim().toLowerCase();
    if (!norm) return;
    Object.keys(FIELD_HEADERS).forEach(function (fieldKey) {
      if (FIELD_HEADERS[fieldKey].toLowerCase() === norm) {
        colIndex[fieldKey] = i;
      }
    });
  });
  return colIndex;
}

function normalize_(val) {
  if (val === null || val === undefined || val === '') return '-';
  return String(val).trim();
}

/**
 * Custom Date Parser untuk menangani berbagai format tanggal (Indonesia/English/Excel)
 */
function parseCustomDate_(pubRaw, tz) {
  if (!pubRaw) return null;

  if (Object.prototype.toString.call(pubRaw) === '[object Date]') {
    if (isNaN(pubRaw.getTime())) return null;
    return Utilities.formatDate(pubRaw, tz, 'yyyy-MM-dd');
  }

  var str = String(pubRaw).trim();
  if (!str || str === '-') return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }

  var months = {
    jan: 1, januari: 1, january: 1,
    feb: 2, februari: 2, february: 2,
    mar: 3, maret: 3, march: 3,
    apr: 4, april: 4,
    mei: 5, may: 5,
    jun: 6, juni: 6, june: 6,
    jul: 7, juli: 7, july: 7,
    agu: 8, agustus: 8, aug: 8, august: 8,
    sep: 9, september: 9,
    okt: 10, oktober: 10, oct: 10, october: 10,
    nov: 11, november: 11,
    des: 12, desember: 12, dec: 12, december: 12
  };

  var parts = str.split(/[\s\/\-\.]+/);
  if (parts.length >= 3) {
    var day = parseInt(parts[0], 10);
    var monthStr = parts[1].toLowerCase();
    var year = parseInt(parts[2], 10);

    var month = parseInt(monthStr, 10);
    if (isNaN(month)) {
      month = months[monthStr] || 0;
    }

    if (day > 0 && day <= 31 && month > 0 && month <= 12 && year > 2000) {
      var mm = month < 10 ? '0' + month : '' + month;
      var dd = day < 10 ? '0' + day : '' + day;
      return year + '-' + mm + '-' + dd;
    }
  }

  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }

  return null;
}

// =====================================================================
// AUTENTIKASI (LOGIN)
// =====================================================================

/**
 * Cek username + password ke sheet "User Management". Kalau cocok, bikin token
 * sesi acak dan simpan info user-nya di CacheService selama SESSION_TTL_SECONDS.
 * Dipanggil dari layar login (script.html).
 */
function login(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();

  if (!username || !password) {
    return { ok: false, error: 'Username dan password wajib diisi.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    return { ok: false, error: 'Sheet "User Management" belum dibuat. Hubungi admin.' };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < USER_FIRST_ROW) {
    return { ok: false, error: 'Belum ada user yang terdaftar.' };
  }

  // Kolom: A Nama Lengkap, B Username, C Email, D WhatsApp, E Password
  var values = sheet.getRange(USER_FIRST_ROW, 1, lastRow - USER_FIRST_ROW + 1, 5).getValues();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowUsername = String(row[1] || '').trim();
    var rowPassword = String(row[4] || '').trim();

    if (rowUsername && rowUsername.toLowerCase() === username.toLowerCase() && rowPassword === password) {
      var token = Utilities.getUuid();
      var session = {
        fullName: String(row[0] || '').trim() || rowUsername,
        username: rowUsername,
        email: String(row[2] || '').trim()
      };
      CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), SESSION_TTL_SECONDS);
      return { ok: true, token: token, fullName: session.fullName, username: session.username };
    }
  }

  return { ok: false, error: 'Username atau password salah.' };
}

/**
 * Hapus sesi dari cache. Dipanggil pas user klik tombol "Keluar".
 */
function logout(token) {
  if (token) {
    CacheService.getScriptCache().remove('SESSION_' + String(token));
  }
  return { ok: true };
}

/**
 * Dipanggil frontend saat pertama kali dibuka, buat ngecek apakah token yang
 * tersimpan di localStorage browser masih berlaku (biar gak perlu login ulang
 * tiap kali refresh halaman selama sesi belum habis).
 */
function checkSession(token) {
  var session = validateSession_(token);
  if (!session) return { ok: false };
  return { ok: true, fullName: session.fullName, username: session.username };
}

/**
 * Helper internal: validasi token sesi, dipakai di awal fungsi-fungsi yang
 * butuh login (mis. getContentPlanData). Return data session kalau valid, null kalau tidak.
 */
function validateSession_(token) {
  token = String(token || '').trim();
  if (!token) return null;
  var cached = CacheService.getScriptCache().get('SESSION_' + token);
  if (!cached) return null;
  try {
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

// =====================================================================
// AGGREGASI DATA LINTAS BRAND
// =====================================================================

function getContentPlanData(token, forceRefresh) {
  var session = validateSession_(token);
  if (!session) {
    throw new Error('UNAUTHORIZED');
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = 'CONTENT_PLAN_DATA_CACHE';

  if (!forceRefresh) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {}
    }
  }

  var result = fetchRawContentPlanData_();

  try {
    cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS);
  } catch (e) {}

  return result;
}

function fetchRawContentPlanData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCES_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('Sheet "' + SOURCES_SHEET_NAME + '" tidak ditemukan.');
  }

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < SOURCES_FIRST_ROW) {
    return { items: [], brands: [], sources: [], warnings: [], generatedAt: new Date().toISOString() };
  }

  var raw = sourceSheet.getRange(SOURCES_FIRST_ROW, 1, lastRow - SOURCES_FIRST_ROW + 1, 4).getValues();
  var tz = Session.getScriptTimeZone();

  var allItems = [];
  var warnings = [];
  var brandNames = [];
  var sourcesList = []; // Dipakai untuk tabel "Brand Terhubung" di tab Home

  raw.forEach(function (row) {
    var brandName = String(row[0] || '').trim();
    var idOrUrl = String(row[1] || '').trim();
    // Kolom C "Nama Sheet Content Plan" -> menentukan sheet mana yang dibaca
    // dari Spreadsheet ID/URL di kolom B. Kalau dikosongkan, fallback ke "Content Plan".
    // Ini memungkinkan 2 brand pakai Spreadsheet ID yang sama tapi sheet berbeda.
    var sheetName = String(row[2] || '').trim() || DEFAULT_CONTENT_PLAN_SHEET;
    var aktif = String(row[3] || '').trim().toLowerCase();

    if (!brandName || !idOrUrl) return;
    if (aktif === 'tidak' || aktif === 'no' || aktif === 'false' || aktif === '0') return;

    brandNames.push(brandName);

    // Brand aktif -> selalu masuk daftar tab Home, terlepas dari sukses/gagalnya fetch data di bawah.
    sourcesList.push({
      brand: brandName,
      sheetName: sheetName,
      url: buildSpreadsheetUrl_(idOrUrl)
    });

    try {
      var ssId = extractSpreadsheetId_(idOrUrl);
      var targetSs = SpreadsheetApp.openById(ssId);
      var targetSheet = targetSs.getSheetByName(sheetName);

      if (!targetSheet) {
        warnings.push(brandName + ': sheet "' + sheetName + '" tidak ditemukan.');
        return;
      }

      var lastCol = targetSheet.getLastColumn();
      var headerRowValues = targetSheet.getRange(CONTENT_HEADER_ROW, 1, 1, lastCol).getValues()[0];
      var colIndex = buildColIndex_(headerRowValues);

      var required = ['publishDate', 'status', 'title'];
      var missing = required.filter(function (f) { return colIndex[f] === undefined; });
      if (missing.length > 0) {
        warnings.push(brandName + ': kolom wajib tidak ditemukan (' +
          missing.map(function (f) { return FIELD_HEADERS[f]; }).join(', ') + ').');
        return;
      }

      var tLastRow = targetSheet.getLastRow();
      if (tLastRow < CONTENT_FIRST_DATA_ROW) return;

      var numRows = tLastRow - CONTENT_FIRST_DATA_ROW + 1;
      var values = targetSheet.getRange(CONTENT_FIRST_DATA_ROW, 1, numRows, lastCol).getValues();

      function field(r, key) {
        var idx = colIndex[key];
        return (idx === undefined) ? '' : r[idx];
      }

      values.forEach(function (r, idx) {
        var pubRaw = field(r, 'publishDate');
        var pubDate = parseCustomDate_(pubRaw, tz);
        if (!pubDate) return;

        allItems.push({
          brand: brandName,
          row: CONTENT_FIRST_DATA_ROW + idx, // nomor baris asli di sheet brand ini, dipakai buat Edit/Delete
          publishDate: pubDate,
          status: normalize_(field(r, 'status')),
          title: normalize_(field(r, 'title')) || '(Tanpa Judul)',
          pilar: normalize_(field(r, 'pilar')),
          format: normalize_(field(r, 'format')),
          channel: normalize_(field(r, 'channel')),
          reference: normalize_(field(r, 'reference')),
          brief: normalize_(field(r, 'brief')),
          owner: normalize_(field(r, 'owner')),
          multimediaCreator: normalize_(field(r, 'multimediaCreator')),
          caption: normalize_(field(r, 'caption')),
          multimedia: normalize_(field(r, 'multimedia'))
        });
      });

    } catch (err) {
      warnings.push(brandName + ': gagal diakses (' + err.message + ').');
    }
  });

  return {
    items: allItems,
    brands: brandNames,
    sources: sourcesList,
    warnings: warnings,
    generatedAt: new Date().toISOString()
  };
}

// =====================================================================
// MENU & UTILITIES
// =====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Candramawa Apps')
    .addItem('Buat/Reset Sheet Brand Sources', 'setupSourcesSheet')
    .addItem('👤 Buat/Reset Sheet User Management', 'setupUserManagementSheet')
    .addItem('➕ Buat Sheet Content Plan Baru', 'createContentPlanSheet')
    .addItem('Clear Cache Dashboard', 'clearDashboardCache')
    .addToUi();
}

function clearDashboardCache() {
  CacheService.getScriptCache().remove('CONTENT_PLAN_DATA_CACHE');
  SpreadsheetApp.getUi().alert('Cache dashboard berhasil dibersihkan.');
}

function setupSourcesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var existing = ss.getSheetByName(SOURCES_SHEET_NAME);

  if (existing) {
    var resp = ui.alert(
      'Sheet "Brand Sources" sudah ada',
      'Reset ke tampilan & contoh default? Data brand yang sudah diisi akan hilang.',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }

  var sheet = ss.insertSheet(SOURCES_SHEET_NAME, 0);

  sheet.getRange('A1:D1').merge()
    .setValue('⚙️ SETUP — CANDRAMAWA APPS')
    .setFontWeight('bold').setFontSize(14)
    .setFontColor('#FFFFFF').setBackground('#1F3864')
    .setHorizontalAlignment('center');

  sheet.getRange('A2:D2').merge()
    .setValue('Daftar spreadsheet Content Plan tiap brand yang datanya mau digabung.')
    .setFontColor('#595959').setFontSize(10).setWrap(true);

  sheet.getRange('A3').setValue('Nama Dashboard').setFontWeight('bold').setFontSize(10);
  sheet.getRange('B3').setValue('Candramawa Apps').setFontWeight('bold').setBackground('#FFF9E6');

  var headers = ['Brand Name', 'Spreadsheet ID / URL', 'Nama Sheet Content Plan', 'Aktif (Ya/Tidak)'];
  sheet.getRange(5, 1, 1, 4).setValues([headers]).setFontWeight('bold').setBackground('#EAF0FB');

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 120);

  sheet.setFrozenRows(5);
  ss.setActiveSheet(sheet);
  ui.alert('Sheet "Brand Sources" berhasil dibuat.');
}

// =====================================================================
// WIZARD: BUAT SHEET CONTENT PLAN BARU
// =====================================================================

/**
 * Wizard untuk membuat sheet "Content Plan" siap pakai untuk sebuah brand —
 * baik di spreadsheet BARU (dibuatkan otomatis) maupun spreadsheet yang SUDAH ADA
 * (tinggal paste ID/URL-nya) — lalu otomatis mendaftarkan brand tsb ke sheet "Brand Sources"
 * di spreadsheet database ini.
 */
function createContentPlanSheet() {
  var ui = SpreadsheetApp.getUi();

  // 1. Nama brand
  var brandResp = ui.prompt('Brand Baru', 'Nama Brand:', ui.ButtonSet.OK_CANCEL);
  if (brandResp.getSelectedButton() !== ui.Button.OK) return;
  var brandName = brandResp.getResponseText().trim();
  if (!brandName) {
    ui.alert('Nama brand tidak boleh kosong.');
    return;
  }

  // 2. Spreadsheet baru atau pakai yang sudah ada?
  var modeResp = ui.alert(
    'Lokasi Spreadsheet',
    'Buat spreadsheet Google Sheets BARU untuk brand "' + brandName + '"?\n\n' +
    '(Pilih "No" kalau mau pakai spreadsheet yang sudah ada)',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (modeResp === ui.Button.CANCEL || modeResp === ui.Button.CLOSE) return;

  var isNewSpreadsheet = (modeResp === ui.Button.YES);
  var targetSs;
  var targetIdOrUrl;

  if (isNewSpreadsheet) {
    targetSs = SpreadsheetApp.create(brandName + ' - Content Plan');
    targetIdOrUrl = targetSs.getId();
  } else {
    var idResp = ui.prompt('Spreadsheet Existing', 'Paste Spreadsheet ID atau URL-nya:', ui.ButtonSet.OK_CANCEL);
    if (idResp.getSelectedButton() !== ui.Button.OK) return;
    var rawIdOrUrl = idResp.getResponseText().trim();
    if (!rawIdOrUrl) {
      ui.alert('Spreadsheet ID/URL tidak boleh kosong.');
      return;
    }
    try {
      targetSs = SpreadsheetApp.openById(extractSpreadsheetId_(rawIdOrUrl));
    } catch (err) {
      ui.alert('Gagal membuka spreadsheet: ' + err.message);
      return;
    }
    targetIdOrUrl = rawIdOrUrl;
  }

  // 3. Nama sheet Content Plan
  var sheetNameResp = ui.prompt(
    'Nama Sheet',
    'Nama sheet Content Plan (kosongkan untuk pakai default "' + DEFAULT_CONTENT_PLAN_SHEET + '"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (sheetNameResp.getSelectedButton() !== ui.Button.OK) return;
  var sheetName = sheetNameResp.getResponseText().trim() || DEFAULT_CONTENT_PLAN_SHEET;

  // Kalau sheet dengan nama itu sudah ada di spreadsheet target, konfirmasi reset dulu
  var existingSheet = targetSs.getSheetByName(sheetName);
  if (existingSheet) {
    var resetResp = ui.alert(
      'Sheet "' + sheetName + '" sudah ada',
      'Reset ke format default? Data yang sudah ada di sheet ini akan hilang.',
      ui.ButtonSet.YES_NO
    );
    if (resetResp !== ui.Button.YES) return;
    targetSs.deleteSheet(existingSheet);
  }

  // 4. Bangun struktur sheet Content Plan
  var sheet = targetSs.insertSheet(sheetName, 0);
  buildContentPlanSheetLayout_(sheet, brandName);

  // Kalau spreadsheet baru dibuat dari nol, bersihkan sheet default "Sheet1" yang kosong
  if (isNewSpreadsheet) {
    var defaultSheet = targetSs.getSheetByName('Sheet1');
    if (defaultSheet && targetSs.getSheets().length > 1) {
      targetSs.deleteSheet(defaultSheet);
    }
  }

  // 5. Daftarkan otomatis ke sheet "Brand Sources" di spreadsheet database ini
  var registerError = null;
  try {
    registerBrandToSources_(brandName, targetIdOrUrl, sheetName);
  } catch (err) {
    registerError = err.message;
  }

  var successMsg = 'Sheet Content Plan untuk "' + brandName + '" sudah dibuat.';
  if (isNewSpreadsheet) successMsg += '\n\nSpreadsheet baru: ' + targetSs.getUrl();
  if (registerError) {
    successMsg += '\n\n⚠️ Tapi GAGAL mendaftarkan otomatis ke Brand Sources: ' + registerError +
      '\nSilakan isi manual: ' + brandName + ' | ' + targetIdOrUrl + ' | ' + sheetName + ' | Ya';
  } else {
    successMsg += '\n\nBrand ini juga sudah otomatis terdaftar di sheet "Brand Sources".';
  }

  ui.alert('Berhasil', successMsg, ui.ButtonSet.OK);
}

/**
 * Membangun struktur sheet Content Plan: judul, catatan, header di baris CONTENT_HEADER_ROW,
 * dropdown validasi untuk kolom Status, dan freeze header — siap langsung dibaca
 * oleh fetchRawContentPlanData_().
 */
function buildContentPlanSheetLayout_(sheet, brandName) {
  var headers = [
    FIELD_HEADERS.publishDate,
    FIELD_HEADERS.status,
    FIELD_HEADERS.title,
    FIELD_HEADERS.pilar,
    FIELD_HEADERS.format,
    FIELD_HEADERS.channel,
    FIELD_HEADERS.reference,
    FIELD_HEADERS.brief,
    FIELD_HEADERS.owner,
    FIELD_HEADERS.multimediaCreator,
    FIELD_HEADERS.caption,
    FIELD_HEADERS.multimedia
  ];
  var colCount = headers.length;

  sheet.getRange(1, 1, 1, colCount).merge()
    .setValue('⚙️ CONTENT PLAN — ' + brandName)
    .setFontWeight('bold').setFontSize(14)
    .setFontColor('#FFFFFF').setBackground('#1F3864')
    .setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, colCount).merge()
    .setValue('Data di sheet ini otomatis ditarik oleh Candramawa Apps. Jangan ubah posisi header di baris ' + CONTENT_HEADER_ROW + '.')
    .setFontColor('#595959').setFontSize(10).setWrap(true);

  sheet.getRange(3, 1, 1, colCount).merge()
    .setValue('📌 Isi data konten mulai baris ' + CONTENT_FIRST_DATA_ROW + ' ke bawah.')
    .setFontColor('#8A8A8A').setFontSize(9).setFontStyle('italic');

  sheet.getRange(CONTENT_HEADER_ROW, 1, 1, colCount)
    .setValues([headers])
    .setFontWeight('bold').setBackground('#EAF0FB');

  // Dropdown validasi untuk kolom Status, biar penulisan status konsisten
  var statusOptions = ['Draft', 'Review', 'Revisi', 'Ready to Post', 'Scheduled', 'Published', 'Cancelled'];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();
  var statusColIndex = headers.indexOf(FIELD_HEADERS.status) + 1;
  sheet.getRange(CONTENT_FIRST_DATA_ROW, statusColIndex, 495, 1).setDataValidation(statusRule);

  sheet.setColumnWidths(1, colCount, 150);
  sheet.setColumnWidth(headers.indexOf(FIELD_HEADERS.title) + 1, 250);
  sheet.setColumnWidth(headers.indexOf(FIELD_HEADERS.reference) + 1, 200);
  sheet.setColumnWidth(headers.indexOf(FIELD_HEADERS.brief) + 1, 200);
  sheet.setColumnWidth(headers.indexOf(FIELD_HEADERS.caption) + 1, 280);
  sheet.setColumnWidth(headers.indexOf(FIELD_HEADERS.multimedia) + 1, 200);

  sheet.setFrozenRows(CONTENT_HEADER_ROW);
}

/**
 * Menambahkan 1 baris baru ke sheet "Brand Sources" di spreadsheet database ini
 * (tempat Code.gs ini ter-bind), dipakai setelah createContentPlanSheet() berhasil.
 */
function registerBrandToSources_(brandName, idOrUrl, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCES_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('Sheet "' + SOURCES_SHEET_NAME + '" belum ada. Jalankan "Buat/Reset Sheet Brand Sources" dulu.');
  }

  var row = SOURCES_FIRST_ROW;
  while (String(sourceSheet.getRange(row, 1).getValue() || '').trim() !== '') {
    row++;
  }

  sourceSheet.getRange(row, 1, 1, 4).setValues([[brandName, idOrUrl, sheetName, 'Ya']]);
}

// =====================================================================
// SETUP SHEET: USER MANAGEMENT
// =====================================================================

/**
 * Membuat/reset sheet "User Management" — daftar user yang boleh login ke dashboard.
 * ⚠️ Password disimpan plain text, jadi batasi akses spreadsheet database ini
 * hanya untuk admin (jangan share ke tim secara umum).
 */
function setupUserManagementSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var existing = ss.getSheetByName(USER_SHEET_NAME);

  if (existing) {
    var resp = ui.alert(
      'Sheet "User Management" sudah ada',
      'Reset ke tampilan default? Data user yang sudah diisi akan hilang.',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }

  var sheet = ss.insertSheet(USER_SHEET_NAME);
  var headers = ['Nama Lengkap', 'Username', 'Email', 'WhatsApp', 'Password'];
  var colCount = headers.length;

  sheet.getRange(1, 1, 1, colCount).merge()
    .setValue('⚙️ USER MANAGEMENT — CANDRAMAWA APPS')
    .setFontWeight('bold').setFontSize(14)
    .setFontColor('#FFFFFF').setBackground('#1F3864')
    .setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, colCount).merge()
    .setValue('Daftar user yang bisa login ke dashboard. ⚠️ Password disimpan plain text — batasi akses spreadsheet ini hanya untuk admin.')
    .setFontColor('#B3261E').setFontSize(10).setWrap(true).setFontWeight('bold');

  sheet.getRange(USER_HEADER_ROW, 1, 1, colCount)
    .setValues([headers])
    .setFontWeight('bold').setBackground('#EAF0FB');

  // Paksa kolom WhatsApp & Password sebagai teks, biar angka nol di depan
  // (mis. nomor WA "0812...") atau password full-angka gak otomatis diubah Sheets jadi angka.
  sheet.getRange(USER_FIRST_ROW, 4, 495, 2).setNumberFormat('@');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);

  sheet.setFrozenRows(USER_HEADER_ROW);
  ss.setActiveSheet(sheet);
  ui.alert('Sheet "User Management" berhasil dibuat. Isi minimal 1 baris user untuk mulai bisa login.');
}

// =====================================================================
// CRUD CONTENT ITEM — Add / Update / Delete (dipanggil dari popup di web app)
// =====================================================================

/**
 * Cari lokasi (spreadsheet + nama sheet) Content Plan untuk 1 brand aktif,
 * berdasarkan sheet "Brand Sources". Dipakai bareng oleh add/update/delete.
 */
function resolveBrandTarget_(brandName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCES_SHEET_NAME);
  if (!sourceSheet) return null;

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < SOURCES_FIRST_ROW) return null;

  var raw = sourceSheet.getRange(SOURCES_FIRST_ROW, 1, lastRow - SOURCES_FIRST_ROW + 1, 4).getValues();
  for (var i = 0; i < raw.length; i++) {
    var rowBrand = String(raw[i][0] || '').trim();
    var idOrUrl = String(raw[i][1] || '').trim();
    var sheetName = String(raw[i][2] || '').trim() || DEFAULT_CONTENT_PLAN_SHEET;
    var aktif = String(raw[i][3] || '').trim().toLowerCase();

    if (rowBrand.toLowerCase() === brandName.toLowerCase() && idOrUrl) {
      if (aktif === 'tidak' || aktif === 'no' || aktif === 'false' || aktif === '0') return null;
      return { idOrUrl: idOrUrl, sheetName: sheetName };
    }
  }
  return null;
}

/**
 * Buka spreadsheet+sheet Content Plan milik 1 brand. Melempar Error kalau gagal,
 * ditangkap oleh pemanggil (addContentItem/updateContentItem/deleteContentItem).
 */
function openBrandContentSheet_(brandName) {
  var target = resolveBrandTarget_(brandName);
  if (!target) {
    throw new Error('Brand "' + brandName + '" tidak ditemukan atau sedang nonaktif.');
  }

  var targetSs = SpreadsheetApp.openById(extractSpreadsheetId_(target.idOrUrl));
  var targetSheet = targetSs.getSheetByName(target.sheetName);
  if (!targetSheet) {
    throw new Error('Sheet "' + target.sheetName + '" tidak ditemukan di spreadsheet brand "' + brandName + '".');
  }
  return targetSheet;
}

/** Tulis 1 field ke array baris, berdasarkan posisi kolomnya di colIndex. Aman kalau kolomnya gak ada di sheet ini. */
function writeFieldToRow_(rowArray, colIndex, fieldKey, value) {
  var idx = colIndex[fieldKey];
  if (idx === undefined) return;
  rowArray[idx] = (value === undefined || value === null) ? '' : String(value).trim();
}

/** Cari baris kosong pertama (berdasarkan kolom Title) buat taruh konten baru. */
function findNextContentRow_(sheet, colIndex) {
  var titleColIndex = (colIndex.title !== undefined) ? colIndex.title : 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < CONTENT_FIRST_DATA_ROW) return CONTENT_FIRST_DATA_ROW;

  var numRows = lastRow - CONTENT_FIRST_DATA_ROW + 1;
  var titleValues = sheet.getRange(CONTENT_FIRST_DATA_ROW, titleColIndex + 1, numRows, 1).getValues();
  for (var i = 0; i < titleValues.length; i++) {
    if (String(titleValues[i][0] || '').trim() === '') {
      return CONTENT_FIRST_DATA_ROW + i;
    }
  }
  return lastRow + 1;
}

/**
 * Tambah 1 konten baru ke sheet Content Plan brand terkait.
 * Wajib diisi: brand, publishDate, title — field lain opsional.
 */
function addContentItem(token, payload) {
  var session = validateSession_(token);
  if (!session) return { ok: false, error: 'Sesi tidak valid atau sudah habis. Silakan login ulang.' };

  payload = payload || {};
  var brandName = String(payload.brand || '').trim();
  var publishDateRaw = String(payload.publishDate || '').trim();
  var title = String(payload.title || '').trim();

  if (!brandName || !publishDateRaw || !title) {
    return { ok: false, error: 'Brand, Tanggal Publish, dan Judul wajib diisi.' };
  }

  var targetSheet;
  try {
    targetSheet = openBrandContentSheet_(brandName);
  } catch (err) {
    return { ok: false, error: err.message };
  }

  var lastCol = targetSheet.getLastColumn();
  var headerRowValues = targetSheet.getRange(CONTENT_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var colIndex = buildColIndex_(headerRowValues);

  if (colIndex.title === undefined || colIndex.publishDate === undefined) {
    return { ok: false, error: 'Struktur header sheet brand ini gak lengkap (kolom Publish Date/Title gak ketemu).' };
  }

  var targetRow = findNextContentRow_(targetSheet, colIndex);

  var rowArray = new Array(lastCol).fill('');
  writeFieldToRow_(rowArray, colIndex, 'publishDate', publishDateRaw);
  writeFieldToRow_(rowArray, colIndex, 'title', title);
  writeFieldToRow_(rowArray, colIndex, 'status', payload.status);
  writeFieldToRow_(rowArray, colIndex, 'pilar', payload.pilar);
  writeFieldToRow_(rowArray, colIndex, 'format', payload.format);
  writeFieldToRow_(rowArray, colIndex, 'channel', payload.channel);
  writeFieldToRow_(rowArray, colIndex, 'reference', payload.reference);
  writeFieldToRow_(rowArray, colIndex, 'brief', payload.brief);
  writeFieldToRow_(rowArray, colIndex, 'owner', payload.owner);
  writeFieldToRow_(rowArray, colIndex, 'multimediaCreator', payload.multimediaCreator);
  writeFieldToRow_(rowArray, colIndex, 'caption', payload.caption);
  writeFieldToRow_(rowArray, colIndex, 'multimedia', payload.multimedia);

  targetSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowArray]);

  return { ok: true, row: targetRow };
}

/**
 * Update 1 baris konten yang sudah ada, ditarget lewat brand + nomor baris asli
 * (dikirim balik oleh frontend dari data yang sebelumnya diambil getContentPlanData).
 */
function updateContentItem(token, payload) {
  var session = validateSession_(token);
  if (!session) return { ok: false, error: 'Sesi tidak valid atau sudah habis. Silakan login ulang.' };

  payload = payload || {};
  var brandName = String(payload.brand || '').trim();
  var rowNumber = parseInt(payload.row, 10);
  var publishDateRaw = String(payload.publishDate || '').trim();
  var title = String(payload.title || '').trim();

  if (!brandName || !rowNumber || !publishDateRaw || !title) {
    return { ok: false, error: 'Brand, Tanggal Publish, dan Judul wajib diisi.' };
  }

  var targetSheet;
  try {
    targetSheet = openBrandContentSheet_(brandName);
  } catch (err) {
    return { ok: false, error: err.message };
  }

  if (rowNumber < CONTENT_FIRST_DATA_ROW || rowNumber > targetSheet.getLastRow()) {
    return { ok: false, error: 'Baris data tidak ditemukan (mungkin sudah berubah). Silakan Refresh Data lalu coba lagi.' };
  }

  var lastCol = targetSheet.getLastColumn();
  var headerRowValues = targetSheet.getRange(CONTENT_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var colIndex = buildColIndex_(headerRowValues);

  if (colIndex.title === undefined || colIndex.publishDate === undefined) {
    return { ok: false, error: 'Struktur header sheet brand ini gak lengkap (kolom Publish Date/Title gak ketemu).' };
  }

  var rowArray = targetSheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  writeFieldToRow_(rowArray, colIndex, 'publishDate', publishDateRaw);
  writeFieldToRow_(rowArray, colIndex, 'title', title);
  writeFieldToRow_(rowArray, colIndex, 'status', payload.status);
  writeFieldToRow_(rowArray, colIndex, 'pilar', payload.pilar);
  writeFieldToRow_(rowArray, colIndex, 'format', payload.format);
  writeFieldToRow_(rowArray, colIndex, 'channel', payload.channel);
  writeFieldToRow_(rowArray, colIndex, 'reference', payload.reference);
  writeFieldToRow_(rowArray, colIndex, 'brief', payload.brief);
  writeFieldToRow_(rowArray, colIndex, 'owner', payload.owner);
  writeFieldToRow_(rowArray, colIndex, 'multimediaCreator', payload.multimediaCreator);
  writeFieldToRow_(rowArray, colIndex, 'caption', payload.caption);
  writeFieldToRow_(rowArray, colIndex, 'multimedia', payload.multimedia);

  targetSheet.getRange(rowNumber, 1, 1, lastCol).setValues([rowArray]);

  return { ok: true };
}

/**
 * Hapus 1 konten. Isi baris DIKOSONGKAN (bukan sheet.deleteRow), sengaja begitu supaya
 * nomor baris item lain di sheet yang sama gak ikut geser — soalnya browser lain yang
 * sesi datanya belum di-refresh masih pegang nomor baris versi lama buat item lain.
 * Baris kosong ini juga otomatis kepakai lagi buat konten baru berikutnya (lihat findNextContentRow_).
 */
function deleteContentItem(token, payload) {
  var session = validateSession_(token);
  if (!session) return { ok: false, error: 'Sesi tidak valid atau sudah habis. Silakan login ulang.' };

  payload = payload || {};
  var brandName = String(payload.brand || '').trim();
  var rowNumber = parseInt(payload.row, 10);

  if (!brandName || !rowNumber) {
    return { ok: false, error: 'Data tidak lengkap untuk menghapus konten ini.' };
  }

  var targetSheet;
  try {
    targetSheet = openBrandContentSheet_(brandName);
  } catch (err) {
    return { ok: false, error: err.message };
  }

  if (rowNumber < CONTENT_FIRST_DATA_ROW || rowNumber > targetSheet.getLastRow()) {
    return { ok: false, error: 'Baris data tidak ditemukan (mungkin sudah dihapus). Silakan Refresh Data.' };
  }

  targetSheet.getRange(rowNumber, 1, 1, targetSheet.getLastColumn()).clearContent();

  return { ok: true };
}
