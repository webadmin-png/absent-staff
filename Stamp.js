// ═══════════════════════════════════════════════════════════════════════
// STAMP.JS — Aksi staf via menu Google Sheets
//
// Berisi:
//   keBarisHariIni() — navigasi ke baris hari ini
//   stampMasuk/Pulang/Istirahat*() — catat jam ke kolom yang sesuai
//   doStamp()        — inti logika pencatatan jam
//   cekRekapSaya()   — tampilkan ringkasan kehadiran bulan ini
// ═══════════════════════════════════════════════════════════════════════

// ── Navigasi ke baris hari ini ────────────────────────────────────────
// Pindahkan kursor ke sel Status (kolom E) baris user hari ini
function keBarisHariIni() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const user  = getInfoUser();
    const sheet = getSheetAktifDivisi(user.divisi);
    if (!sheet) throw new Error('Sheet divisi tidak ditemukan. Jalankan appendHariIni() dulu.');

    ss.setActiveSheet(sheet);
    const targetRow = cariBarisSaya(sheet, user.nama);

    if (targetRow === -1) {
      SpreadsheetApp.getUi().alert(
        '⚠ Baris hari ini belum tersedia.\n\n' +
        'Kemungkinan trigger belum jalan.\n' +
        'Minta HRD jalankan appendHariIni().'
      );
      return;
    }

    sheet.setActiveRange(sheet.getRange(targetRow, COL_STATUS));
    SpreadsheetApp.getUi().alert(
      '📍 Sudah ke baris kamu!\n\n' +
      'Nama  : ' + user.nama   + '\n' +
      'Divisi: ' + user.divisi + '\n' +
      'Sheet : ' + sheet.getName()
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

// ── Shortcut stamp — masing-masing memanggil doStamp() ───────────────
function stampMasuk()       { doStamp(COL_MASUK,  'Masuk',               'Hadir'); }
function stampIst1Mulai()   { doStamp(COL_IST1_M, 'Istirahat 1 Mulai',   null); }
function stampIst1Selesai() { doStamp(COL_IST1_S, 'Istirahat 1 Selesai', null); }
function stampIst2Mulai()   { doStamp(COL_IST2_M, 'Istirahat 2 Mulai',   null); }
function stampIst2Selesai() { doStamp(COL_IST2_S, 'Istirahat 2 Selesai', null); }
function stampPulang()      { doStamp(COL_PULANG,  'Pulang',              null); }

// ── doStamp — Inti pencatatan jam ─────────────────────────────────────
// Tulis jam sekarang ke kolom `kolom` pada baris hari ini milik user.
// Jika kolom sudah terisi, tanya konfirmasi sebelum menimpa.
// `statusValue` — jika diisi, juga update kolom Status (misal "Hadir" saat masuk)
function doStamp(kolom, labelAksi, statusValue) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const user      = getInfoUser();
    const sheet     = getSheetAktifDivisi(user.divisi);
    if (!sheet) throw new Error('Sheet divisi tidak ditemukan.');

    const targetRow = cariBarisSaya(sheet, user.nama);
    if (targetRow === -1) {
      SpreadsheetApp.getUi().alert('⚠ Baris hari ini belum tersedia. Hubungi HRD.');
      return;
    }

    const jam      = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm');
    const existing = sheet.getRange(targetRow, kolom).getValue();

    if (existing !== '') {
      const ui   = SpreadsheetApp.getUi();
      const resp = ui.alert(
        '⚠ ' + labelAksi + ' sudah terisi: ' + existing,
        'Timpa dengan jam sekarang (' + jam + ')?',
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    }

    sheet.getRange(targetRow, kolom).setValue(jam);
    if (statusValue) sheet.getRange(targetRow, COL_STATUS).setValue(statusValue);

    ss.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(targetRow, kolom));

    SpreadsheetApp.getUi().alert(
      '✅ ' + labelAksi + ' tercatat!\n\n' +
      'Jam   : ' + jam        + '\n' +
      'Nama  : ' + user.nama  + '\n' +
      'Divisi: ' + user.divisi
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

// ── Rekap kehadiran pribadi ───────────────────────────────────────────
// Hitung dan tampilkan ringkasan status absensi user bulan ini
function cekRekapSaya() {
  try {
    const user  = getInfoUser();
    const sheet = getSheetAktifDivisi(user.divisi);
    if (!sheet) throw new Error('Sheet divisi tidak ditemukan.');

    const today = getToday();
    const data  = sheet.getDataRange().getValues();
    let hadir = 0, sakit = 0, izin = 0, alpha = 0, belum = 0;

    for (let i = 3; i < data.length; i++) {
      const tgl       = data[i][0];
      const namaBaris = String(data[i][COL_NAMA   - 1]).trim();
      const status    = String(data[i][COL_STATUS - 1]).trim();

      if (namaBaris !== user.nama) continue;
      if (!(tgl instanceof Date)) continue;

      const tglD = new Date(tgl.getFullYear(), tgl.getMonth(), tgl.getDate());
      if (tglD.getTime() > today.getTime()) continue;

      if      (status === 'Hadir') hadir++;
      else if (status === 'Sakit') sakit++;
      else if (status === 'Izin')  izin++;
      else if (status === 'Alpha') alpha++;
      else belum++;
    }

    const total = hadir + sakit + izin + alpha + belum;
    const pct   = total > 0 ? Math.round(hadir / total * 100) : 0;

    SpreadsheetApp.getUi().alert(
      '📊 Rekap Absensi Bulan Ini\n' +
      user.nama + ' — ' + user.divisi + '\n' +
      '─────────────────────────\n' +
      '✅ Hadir    : ' + hadir + ' hari\n' +
      '🤒 Sakit    : ' + sakit + ' hari\n' +
      '📝 Izin     : ' + izin  + ' hari\n' +
      '❌ Alpha    : ' + alpha + ' hari\n' +
      '⏳ Belum isi: ' + belum + ' hari\n' +
      '─────────────────────────\n' +
      'Kehadiran: ' + pct + '%'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}
