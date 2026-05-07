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

// Cari sheet divisi bulan ini (contoh: "WEB_Apr_2026")
// Fallback ke sheet tanpa tanggal ("WEB") jika tidak ditemukan
function getSheetAktifDivisi(divisi) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const namaBulan = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MMM_yyyy');
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
