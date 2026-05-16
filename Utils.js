// ═══════════════════════════════════════════════════════════════════════
// UTILS.JS — Fungsi utilitas yang dipakai di seluruh project
// Tidak ada logika bisnis di sini — hanya helper murni.
// ═══════════════════════════════════════════════════════════════════════

// Kembalikan date hari ini pada jam 12:00 (menghindari ambiguitas timezone)
function getToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0);
}

// Cek apakah dua nilai Date jatuh pada hari yang sama (abaikan jam)
function isSameDate(val, today) {
  if (!val || !(val instanceof Date)) return false;
  const d = new Date(val.getFullYear(),   val.getMonth(),   val.getDate());
  const t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return d.getTime() === t.getTime();
}

// Cek apakah val sudah lewat dari today (strictly before)
function isPast(val, today) {
  if (!val || !(val instanceof Date)) return false;
  const d = new Date(val.getFullYear(),   val.getMonth(),   val.getDate());
  const t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return d.getTime() < t.getTime();
}

// Cari sheet divisi untuk bulan tertentu (default: hari ini)
// Contoh: "WEB_Apr_2026". Fallback ke sheet tanpa tanggal ("WEB").
function getSheetAktifDivisi(divisi, date) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const ref       = (date instanceof Date) ? date : new Date();
  const namaBulan = Utilities.formatDate(ref, CONFIG.TIMEZONE, 'MMM_yyyy');
  return ss.getSheetByName(divisi + '_' + namaBulan)
      || ss.getSheetByName(divisi)
      || null;
}

// Ambil info user (divisi, nama, email) dari Master_Data berdasarkan sesi aktif
// Dipakai untuk fungsi menu/stamp di Google Sheets (bukan Web App)
function getInfoUser() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const email  = Session.getEffectiveUser().getEmail();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) throw new Error('Sheet ' + CONFIG.SHEET_MASTER + ' tidak ditemukan.');

  const data = master.getRange('A4:D200').getValues().filter(r => r[0] !== '');
  const row  = data.find(r =>
    String(r[2]).trim().toLowerCase() === email.toLowerCase()
  );

  if (!row) throw new Error(
    'Email ' + email + ' tidak terdaftar di Master_Data.\n' +
    'Hubungi HRD untuk mendaftarkan email kamu.'
  );

  return {
    divisi: String(row[0]).trim(),
    nama  : String(row[1]).trim(),
    email : String(row[2]).trim(),
  };
}

// Temukan nomor baris (1-indexed) milik `nama` untuk hari ini di sheet
// Kembalikan -1 jika tidak ditemukan
function cariBarisSaya(sheet, nama) {
  const today = getToday();
  const data  = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) {
    const tgl       = data[i][0];
    const namaBaris = String(data[i][COL_NAMA - 1]).trim();
    if (namaBaris === nama && isSameDate(tgl, today)) return i + 1;
  }
  return -1;
}

// Tampilkan daftar semua nama sheet — berguna saat setup awal
function cekNamaSheet() {
  const names = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets().map(s => '"' + s.getName() + '"');
  SpreadsheetApp.getUi().alert('Sheet yang ada:\n\n' + names.join('\n'));
}

// Konversi jam desimal ke string "HH:MM"
// Contoh: 7.5 → "07:30",  1.25 → "01:15",  -0.5 → "-00:30"
function decimalToHHMM(decimal) {
  if (!decimal || isNaN(decimal) || decimal === 0) return '00:00';
  const totalMenit = Math.round(Math.abs(decimal) * 60);
  const jam        = Math.floor(totalMenit / 60);
  const menit      = totalMenit % 60;
  const sign       = decimal < 0 ? '-' : '';
  return sign + String(jam).padStart(2, '0') + ':' + String(menit).padStart(2, '0');
}

// Parse nilai kolom L/M/N/O (time fraction hari) → jam desimal
// Mendukung: Date object, number (fraction), string "HH:MM", string "7j 30m"
function parseTimeFraction(val) {
  if (val === null || val === undefined || val === '' || val === '—') return 0;

  if (val instanceof Date) {
    return val.getHours() + val.getMinutes() / 60;
  }
  if (typeof val === 'number') {
    return val * 24;  // fraction hari → jam
  }

  const str = String(val).trim();
  if (str === '' || str === '—') return 0;

  // Format "7j 30m" atau "7j"
  if (str.includes('j')) {
    const jamMatch   = str.match(/(\d+(?:\.\d+)?)j/);
    const menitMatch = str.match(/(\d+)m/);
    return (jamMatch   ? parseFloat(jamMatch[1])   : 0) +
           (menitMatch ? parseInt(menitMatch[1]) / 60 : 0);
  }

  // Format "07:30"
  if (str.includes(':')) {
    const parts = str.split(':');
    if (parts.length >= 2) return parseInt(parts[0]) + parseInt(parts[1]) / 60;
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num * 24;
}

// Parse string jam dari kolom sheet → jam desimal
// Mendukung: "7j 30m", "07:30", "0.5j", number langsung
function parseHHMM(val) {
  if (!val || val === '—' || val === '') return 0;
  if (typeof val === 'number') return val;

  const str = String(val).trim();
  if (str === '—' || str === '') return 0;

  if (str.includes('j')) {
    const jamMatch   = str.match(/(\d+)j/);
    const menitMatch = str.match(/(\d+)m/);
    return (jamMatch   ? parseInt(jamMatch[1])   : 0) +
           (menitMatch ? parseInt(menitMatch[1]) / 60 : 0);
  }

  if (str.includes(':')) {
    const parts = str.split(':');
    if (parts.length === 2) return parseInt(parts[0]) + parseInt(parts[1]) / 60;
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

// ─── AUDIT LOG ────────────────────────────────────────────────────────
// Catat perubahan jam absensi (input/hapus jam) ke sheet '_AuditLog'
// per spreadsheet divisi. Sheet dibuat lazy saat log pertama, di-hide
// dari staf, dan di-protect (owner + admin only).

function _getAuditSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('_AuditLog');
  const isNew = !sheet;

  if (isNew) {
    sheet = ss.insertSheet('_AuditLog');
    sheet.appendRow([
      'Timestamp','Email','Sumber','Sheet','Row','Kolom','Aksi','Nilai Lama','Nilai Baru'
    ]);
    sheet.getRange(1,1,1,9)
      .setBackground('#178232').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 9, 130);
    try { sheet.hideSheet(); } catch(e) {}

    // Protect: hanya owner + admin yang bisa edit (sistem append via service account owner)
    try {
      const prot = sheet.protect();
      prot.setDescription('Audit log — append-only oleh sistem');
      prot.setWarningOnly(false);
      prot.removeEditors(prot.getEditors());
      prot.addEditor(Session.getEffectiveUser());
      for (const adminEmail of (CONFIG.ADMIN_EMAILS || [])) {
        try { prot.addEditor(adminEmail); } catch(e) {}
      }
    } catch(e) { Logger.log('⚠ Proteksi _AuditLog gagal: ' + e.message); }
  }

  // Format kolom (idempotent — aman dipanggil setiap kali, juga memperbaiki
  // sheet lama yang dibuat sebelum fix HH:mm display).
  // Sheets auto-convert string "09:29" → fraction-of-day; pakai format HH:mm
  // supaya nilai (apapun penyimpanannya) ditampilkan sebagai jam.
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm:ss');
  sheet.getRange('H:I').setNumberFormat('HH:mm');

  // Format level KOLOM tidak override cell yang sudah punya format sendiri
  // (mis. cell yang ditulis appendRow dengan auto-detect format).
  // Force re-apply ke seluruh range data H:I yang sudah ada.
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
    sheet.getRange(2, 8, lastRow - 1, 2).setNumberFormat('HH:mm');
  }

  return sheet;
}

// Map nomor kolom jam → label readable. Return null untuk kolom non-jam.
function _kolomJamLabel(col) {
  if (typeof COL_MASUK === 'undefined') return null;
  switch (col) {
    case COL_MASUK:  return 'Masuk';
    case COL_IST1_M: return 'Ist1 Mulai';
    case COL_IST1_S: return 'Ist1 Selesai';
    case COL_IST2_M: return 'Ist2 Mulai';
    case COL_IST2_S: return 'Ist2 Selesai';
    case COL_PULANG: return 'Pulang';
    default: return null;
  }
}

// Format nilai jam (Date / fraction-of-day / string) → "HH:mm" atau "" jika kosong
function _formatJamLog(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'HH:mm');
  }
  if (typeof val === 'number') {
    const totalMin = Math.round(val * 24 * 60);
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  return String(val).trim();
}

// Tentukan aksi dari nilai lama vs baru (string-formatted)
function _aksiJam(lama, baru) {
  if (!lama && baru)  return 'create';
  if (lama && !baru)  return 'delete';
  if (lama && baru && lama !== baru) return 'update';
  return null;  // tidak ada perubahan
}

// Append satu baris ke audit log. Tidak pernah throw — kegagalan log
// tidak boleh memutus alur utama (submit absen, onEdit, dll).
function _logAuditJam(opts) {
  try {
    const lama = _formatJamLog(opts.nilaiLama);
    const baru = _formatJamLog(opts.nilaiBaru);
    const aksi = _aksiJam(lama, baru);
    if (!aksi) return;  // tidak ada perubahan — skip

    const sheet = _getAuditSheet();
    sheet.appendRow([
      new Date(),
      opts.email || '',
      opts.sumber || '',           // 'webapp' | 'sheet'
      opts.targetSheet || '',
      opts.row || '',
      opts.kolomLabel || '',
      aksi,
      lama,
      baru,
    ]);
  } catch(e) {
    Logger.log('⚠ _logAuditJam gagal: ' + e.message);
  }
}

// Konversi nomor kolom (1-indexed) ke huruf Excel
// Contoh: 1 → "A",  26 → "Z",  27 → "AA"
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col    = Math.floor((col - 1) / 26);
  }
  return letter;
}
