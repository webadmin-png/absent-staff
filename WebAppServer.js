// ═══════════════════════════════════════════════════════════════════════
// WEBAPP.JS — Form absensi berbasis Web App untuk akses staf
//
// Deploy: Extensions → Apps Script → Deploy → New deployment
//         Type: Web App
//         Execute as: User accessing the web app  (USER_ACCESSING)
//         Access: Anyone with a Google Account    (ANYONE)
//
// Berisi:
//   doGet()           — serve halaman HTML form
//   getDataHariIni()  — baca data baris hari ini milik user yang login
//   submitAbsensi()   — simpan input form ke sheet
//   _webIsLocked()    — cek apakah baris sudah dikunci (≥30 menit sejak pulang)
//   _webFmtTime()     — format nilai waktu dari sheet ke string "HH:mm"
//   _webSetTime()     — tulis string "HH:mm" ke sel sheet
// ═══════════════════════════════════════════════════════════════════════

// ── doGet — Entry point Web App ───────────────────────────────────────
// Dipanggil browser saat user membuka URL Web App.
// Mengembalikan halaman HTML dari file WebApp.html.
function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('Absensi — ' + CONFIG.NAMA_INSTANSI)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── getDataHariIni — Baca data baris hari ini milik user ──────────────
// Dipanggil client lewat google.script.run.getDataHariIni()
// USER_ACCESSING: getEffectiveUser() selalu mengembalikan email user yang login
// Return: objek { nama, divisi, email, tanggal, hari, rowIdx, locked, fields }
function getDataHariIni() {
  _loadSettings();
  const email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!email) throw new Error(
    'Akun Google tidak terdeteksi.\n' +
    'Pastikan Anda sudah login dan Web App sudah di-redeploy setelah update manifest.'
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan. Pastikan script terikat ke spreadsheet.');
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) throw new Error('Sheet Master_Data tidak ditemukan.');

  // Cari user di Master_Data — divisi di-uppercase agar cocok dengan nama sheet
  const masterRows = master.getRange('A4:D200').getValues();
  let userInfo = null;
  for (const r of masterRows) {
    if (
      String(r[2]).trim().toLowerCase() === email &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    ) {
      userInfo = {
        divisi: String(r[0]).trim().toUpperCase(),
        nama  : String(r[1]).trim(),
        email
      };
      break;
    }
  }
  if (!userInfo) throw new Error(
    'Email tidak ditemukan di Master_Data atau akun tidak aktif.\n' +
    'Email terdeteksi: ' + email + '\n' +
    'Hubungi HRD untuk pendaftaran.'
  );

  const sheet = getSheetAktifDivisi(userInfo.divisi);
  if (!sheet) throw new Error(
    'Sheet divisi "' + userInfo.divisi + '" tidak ditemukan.\n' +
    'Pastikan sheet bulan ini sudah dibuat oleh HRD.'
  );

  const today   = getToday();
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) throw new Error('Baris hari ini belum tersedia. Hubungi HRD.');

  const data  = sheet.getRange(4, 1, lastRow - 3, TOTAL_COL).getValues();
  let rowData = null;
  let rowIdx  = -1;

  for (let i = 0; i < data.length; i++) {
    if (
      isSameDate(data[i][0], today) &&
      String(data[i][COL_NAMA - 1]).trim() === userInfo.nama
    ) {
      rowData = data[i];
      rowIdx  = i + 4; // 1-indexed sheet row
      break;
    }
  }
  if (!rowData) throw new Error(
    'Baris hari ini belum tersedia.\n' +
    'Kemungkinan trigger belum jalan. Hubungi HRD.'
  );

  const pulangVal = rowData[COL_PULANG - 1];

  return {
    nama   : userInfo.nama,
    divisi : userInfo.divisi,
    email  : email,
    tanggal: Utilities.formatDate(today, CONFIG.TIMEZONE, 'dd/MM/yyyy'),
    hari   : String(rowData[COL_HARI - 1]),
    rowIdx : rowIdx,
    locked : _webIsLocked(pulangVal),
    fields : {
      status          : String(rowData[COL_STATUS       - 1] || ''),
      masuk           : _webFmtTime(rowData[COL_MASUK    - 1]),
      ist1Mulai       : _webFmtTime(rowData[COL_IST1_M   - 1]),
      ist1Selesai     : _webFmtTime(rowData[COL_IST1_S   - 1]),
      ist2Mulai       : _webFmtTime(rowData[COL_IST2_M   - 1]),
      ist2Selesai     : _webFmtTime(rowData[COL_IST2_S   - 1]),
      pulang          : _webFmtTime(pulangVal),
      keterangan      : String(rowData[COL_KETERANGAN    - 1] || ''),
      catatanTelat    : String(rowData[COL_TELAT         - 1] || ''),
      catatanPulangAwal: String(rowData[COL_PULANG_AWAL  - 1] || ''),
    },
    plan        : String(rowData[COL_PLAN - 1] || ''),
    planOptions : CONFIG.PLAN_JAM,
  };
}

// ── submitAbsensi — Simpan input form ke sheet ────────────────────────
// Dipanggil client lewat google.script.run.submitAbsensi(payload)
//
// Dua tipe payload:
//   { rowIdx, type: 'stamp',      action: 'masuk'|'ist1Mulai'|..., time: 'HH:mm' }
//   { rowIdx, type: 'tidakHadir', status: 'Izin'|'Sakit'|'Alpha' }
//
// Verifikasi ulang email di server — tidak percaya data dari client.
function submitAbsensi(payload) {
  _loadSettings();
  const email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!email) throw new Error('Akun Google tidak terdeteksi.');

  // Serialisasi akses spreadsheet agar tidak terjadi race condition
  // saat banyak staf submit bersamaan (misal 30 orang jam 07:30)
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(10000); // tunggu max 10 detik
  if (!acquired) {
    throw new Error('Server sedang sibuk, coba lagi dalam beberapa detik.');
  }

  try {
    return _doSubmitAbsensi(email, payload);
  } finally {
    lock.releaseLock();
  }
}

function _doSubmitAbsensi(email, payload) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) throw new Error('Master_Data tidak ditemukan.');

  // Verifikasi user dari Master_Data (bukan dari payload client)
  const masterRows = master.getRange('A4:D200').getValues();
  let userInfo = null;
  for (const r of masterRows) {
    if (
      String(r[2]).trim().toLowerCase() === email &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    ) {
      userInfo = {
        divisi: String(r[0]).trim().toUpperCase(),
        nama  : String(r[1]).trim()
      };
      break;
    }
  }
  if (!userInfo) throw new Error('Akses ditolak.');

  const sheet = getSheetAktifDivisi(userInfo.divisi);
  if (!sheet) throw new Error('Sheet divisi tidak ditemukan.');

  const row = payload.rowIdx;
  if (!row || row < 4) throw new Error('Data baris tidak valid.');

  // Pastikan baris ini memang milik user yang login
  const rowEmail = String(sheet.getRange(row, COL_EMAIL).getValue()).trim().toLowerCase();
  if (rowEmail !== email) throw new Error('Akses ditolak — baris bukan milik Anda.');

  // Cek lock dari sheet langsung (bukan dari client) agar tidak bisa di-bypass
  const currentPulang = sheet.getRange(row, COL_PULANG).getValue();
  if (_webIsLocked(currentPulang)) {
    throw new Error('Baris sudah terkunci (30 menit setelah jam pulang).');
  }

  if (payload.type === 'tidakHadir') {
    // ── Lapor tidak hadir: tulis status, kosongkan semua waktu, simpan keterangan ──
    const allowed = ['Izin', 'Sakit', 'Alpha'];
    if (!allowed.includes(payload.status)) throw new Error('Status tidak valid.');
    sheet.getRange(row, COL_STATUS).setValue(payload.status);

    // Audit log: catat pembersihan jam (delete) untuk setiap kolom yang sebelumnya terisi
    const TIME_COLS = [COL_MASUK, COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_PULANG];
    const oldTimes = TIME_COLS.map(c => sheet.getRange(row, c).getValue());
    TIME_COLS.forEach((c, idx) => {
      sheet.getRange(row, c).clearContent();
      _logAuditJam({
        email: email, sumber: 'webapp', targetSheet: sheet.getName(), row: row,
        kolomLabel: _kolomJamLabel(c), nilaiLama: oldTimes[idx], nilaiBaru: '',
      });
    });
    // Keterangan — string bebas, boleh kosong
    const keterangan = typeof payload.keterangan === 'string'
      ? payload.keterangan.trim().substring(0, 500) // cap 500 karakter
      : '';
    sheet.getRange(row, COL_KETERANGAN).setValue(keterangan);

  } else if (payload.type === 'stamp') {
    // ── Catat satu aksi: tulis jam ke kolom yang sesuai ──
    const ACTION_MAP = {
      masuk      : COL_MASUK,
      ist1Mulai  : COL_IST1_M,
      ist1Selesai: COL_IST1_S,
      ist2Mulai  : COL_IST2_M,
      ist2Selesai: COL_IST2_S,
      pulang     : COL_PULANG,
    };
    const col = ACTION_MAP[payload.action];
    if (!col) throw new Error('Aksi tidak valid: ' + payload.action);
    if (!payload.time) throw new Error('Jam tidak boleh kosong.');

    // Baca nilai lama SEBELUM overwrite untuk audit log
    const oldVal = sheet.getRange(row, col).getValue();
    _webSetTime(sheet, row, col, payload.time);
    _logAuditJam({
      email: email, sumber: 'webapp', targetSheet: sheet.getName(), row: row,
      kolomLabel: _kolomJamLabel(col), nilaiLama: oldVal, nilaiBaru: payload.time,
    });

    // Stamp masuk → otomatis set Status = Hadir
    if (payload.action === 'masuk') {
      sheet.getRange(row, COL_STATUS).setValue('Hadir');
    }

    // Simpan catatan telat ke COL_TELAT, catatan pulang awal ke COL_PULANG_AWAL
    if (payload.action === 'masuk' &&
        typeof payload.keterangan === 'string' && payload.keterangan.trim()) {
      sheet.getRange(row, COL_TELAT).setValue(payload.keterangan.trim().substring(0, 500));
    }
    if (payload.action === 'pulang' &&
        typeof payload.keterangan === 'string' && payload.keterangan.trim()) {
      sheet.getRange(row, COL_PULANG_AWAL).setValue(payload.keterangan.trim().substring(0, 500));
    }

    // Stamp masuk/pulang → simpan status UPS & PC (replace, bukan append)
    if (payload.action === 'masuk' || payload.action === 'pulang') {
      const label = payload.action === 'masuk' ? 'ON' : 'OFF';
      const ups   = payload.ups  ? 'UPS:' + label : 'UPS:-';
      const pc    = payload.pc   ? 'PC:'  + label : 'PC:-';
      const newSegment = ups + ' ' + pc;

      const existing = String(sheet.getRange(row, COL_DEVICE).getValue() || '');
      const slots    = _parseDeviceStatus(existing);
      if (payload.action === 'masuk')  slots.masuk  = newSegment;
      else                              slots.pulang = newSegment;
      sheet.getRange(row, COL_DEVICE).setValue(_renderDeviceStatus(slots));
    }

  } else if (payload.type === 'plan') {
    // ── Pilih plan jam kerja ──
    const allowed = CONFIG.PLAN_JAM;
    if (!allowed.includes(payload.plan)) throw new Error('Pilihan plan tidak valid.');
    sheet.getRange(row, COL_PLAN).setValue(payload.plan);

  } else if (payload.type === 'deleteJam') {
    // ── Hapus satu kolom jam (untuk koreksi salah input) ──
    const ACTION_MAP = {
      masuk      : COL_MASUK,
      ist1Mulai  : COL_IST1_M,
      ist1Selesai: COL_IST1_S,
      ist2Mulai  : COL_IST2_M,
      ist2Selesai: COL_IST2_S,
      pulang     : COL_PULANG,
    };
    const col = ACTION_MAP[payload.action];
    if (!col) throw new Error('Aksi tidak valid: ' + payload.action);

    const oldVal = sheet.getRange(row, col).getValue();
    sheet.getRange(row, col).clearContent();

    // Kalau yang dihapus adalah Masuk dan tidak ada jam lain → reset Status juga
    if (payload.action === 'masuk') {
      const otherTimes = [COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_PULANG]
        .map(c => sheet.getRange(row, c).getValue())
        .some(v => v !== '' && v !== null);
      if (!otherTimes) sheet.getRange(row, COL_STATUS).clearContent();
    }

    // Hapus device segment yang bersangkutan (masuk → ON segment, pulang → OFF segment)
    if (payload.action === 'masuk' || payload.action === 'pulang') {
      const existing = String(sheet.getRange(row, COL_DEVICE).getValue() || '');
      const slots    = _parseDeviceStatus(existing);
      if (payload.action === 'masuk')  slots.masuk  = '';
      else                              slots.pulang = '';
      sheet.getRange(row, COL_DEVICE).setValue(_renderDeviceStatus(slots));
    }

    _logAuditJam({
      email: email, sumber: 'webapp', targetSheet: sheet.getName(), row: row,
      kolomLabel: _kolomJamLabel(col), nilaiLama: oldVal, nilaiBaru: '',
    });

  } else {
    throw new Error('Tipe submit tidak valid.');
  }

  Logger.log('✓ Web submit: ' + email + ' baris ' + row +
    ' type=' + payload.type +
    (payload.action ? ' action=' + payload.action : '') +
    (payload.status ? ' status=' + payload.status : ''));
  return { ok: true };
}

// ── _webIsLocked — Cek apakah baris sudah dikunci ────────────────────
// Baris dianggap terkunci jika jam pulang sudah diisi ≥30 menit yang lalu.
// Mendukung tiga format nilai pulang dari Sheets: Date, string "HH:mm", angka fraksi.
function _webIsLocked(pulangVal) {
  if (!pulangVal && pulangVal !== 0) return false;
  const now = new Date();
  let jamPulang;
  if (pulangVal instanceof Date) {
    jamPulang = new Date(
      now.getFullYear(), now.getMonth(), now.getDate(),
      pulangVal.getHours(), pulangVal.getMinutes(), 0
    );
  } else if (typeof pulangVal === 'string' && pulangVal.includes(':')) {
    const [h, m] = pulangVal.split(':');
    jamPulang = new Date(
      now.getFullYear(), now.getMonth(), now.getDate(),
      parseInt(h), parseInt(m), 0
    );
  } else if (typeof pulangVal === 'number' && pulangVal > 0) {
    // Nilai fraksi hari dari Sheets (contoh: 0.333 = 08:00)
    const totalMenit = Math.round(pulangVal * 1440);
    jamPulang = new Date(
      now.getFullYear(), now.getMonth(), now.getDate(),
      Math.floor(totalMenit / 60), totalMenit % 60, 0
    );
  } else {
    return false;
  }
  if (jamPulang > now) jamPulang.setDate(jamPulang.getDate() - 1);
  return (now - jamPulang) / 60000 >= 30;
}

// ── _webFmtTime — Format nilai waktu dari Sheets ke "HH:mm" ──────────
// Sheets menyimpan waktu sebagai fraksi hari (0.333 = 08:00),
// objek Date, atau kadang string "HH:mm". Fungsi ini menormalisasi ketiganya.
function _webFmtTime(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'HH:mm');
  }
  if (typeof val === 'number' && val > 0) {
    const totalMenit = Math.round(val * 1440);
    const h = Math.floor(totalMenit / 60);
    const m = totalMenit % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  if (typeof val === 'string' && val.includes(':')) return val;
  return '';
}

// ── _webSetTime — Tulis string "HH:mm" ke sel sheet ──────────────────
// Sheets akan auto-konversi string waktu menjadi time serial.
// Jika timeStr kosong, sel dikosongkan.
function _webSetTime(sheet, row, col, timeStr) {
  const cell = sheet.getRange(row, col);
  if (!timeStr) {
    cell.clearContent();
    return;
  }
  cell.setValue(timeStr).setNumberFormat('HH:mm');
}

// ── Device status helpers ────────────────────────────────────────────
// Format kolom T (DEVICE):
//   "<masuk-segment> | <pulang-segment>"
//   contoh: "UPS:ON PC:ON | UPS:OFF PC:OFF"
// Masuk segment selalu mengandung "ON" (atau hanya "-" kalau dua-duanya unchecked)
// Pulang segment selalu mengandung "OFF" (atau hanya "-")
// Parser pakai presence ON/OFF dulu, fallback ke posisi (slot 0 = masuk, 1 = pulang).

function _parseDeviceStatus(existing) {
  const parts = String(existing || '').split('|').map(s => s.trim()).filter(Boolean);
  let masuk = '', pulang = '';
  const ambiguous = [];
  for (const p of parts) {
    const hasOn  = /\bON\b/.test(p);
    const hasOff = /\bOFF\b/.test(p);
    if (hasOn && !hasOff)      masuk  = p;
    else if (hasOff && !hasOn) pulang = p;
    else                        ambiguous.push(p); // "UPS:- PC:-" — pakai posisi
  }
  // Fallback posisi untuk segment ambigu (semua dash)
  for (let i = 0; i < ambiguous.length; i++) {
    if (!masuk)       masuk  = ambiguous[i];
    else if (!pulang) pulang = ambiguous[i];
  }
  return { masuk, pulang };
}

function _renderDeviceStatus(slots) {
  const out = [];
  if (slots.masuk)  out.push(slots.masuk);
  if (slots.pulang) out.push(slots.pulang);
  return out.join(' | ');
}
