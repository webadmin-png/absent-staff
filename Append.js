// ═══════════════════════════════════════════════════════════════════════
// APPEND.JS — Penambahan baris harian + tampilan sheet
// TEST 3
// Berisi:
//   appendHariIni()      — append baris staf untuk hari ini (jaring pengaman 06:00)
//   appendBesok()        — append baris staf untuk besok (trigger 22:00)
//   _appendUntukTanggal()— inti append untuk satu tanggal (dipakai keduanya)
//   highlightHariIni()   — warnai baris hari ini (kuning) dan lewat (abu)
//   groupByToday()       — collapse baris hari lama, buka hari ini
// ═══════════════════════════════════════════════════════════════════════

// ── appendHariIni — Append baris staf untuk HARI INI ──────────────────
// Dipanggil otomatis pukul 06:00 via trigger (jaring pengaman), atau manual
// oleh HRD. Inti logikanya ada di _appendUntukTanggal(); di sini hanya guard
// admin, lock anti-paralel, dan notifikasi UI.
function appendHariIni() {
  _requireAdmin();

  // Lock: cegah eksekusi paralel (trigger + admin manual klik bersamaan)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⚠ appendHariIni: tidak bisa acquire lock dalam 30s — skip.');
    return;
  }

  try {
    const r = _appendUntukTanggal(getToday());
    try {
      SpreadsheetApp.getUi().alert(
        '✅ Append hari ini selesai!\n' +
        r.tglStr + ' (' + r.namaHari + ')\n\n' +
        r.hasil.join('\n')
      );
    } catch(e) {
      // Dipanggil dari trigger — tidak ada UI
    }
  } finally {
    lock.releaseLock();
  }
}

// ── appendBesok — Append baris staf untuk BESOK (hari ini + 1) ─────────
// Dipanggil otomatis pukul 22:00 via trigger, atau manual oleh HRD.
// Tujuannya: baris untuk esok hari sudah tersedia SEBELUM tengah malam,
// sehingga staf shift malam/dini hari (akses 00:00–06:00) tidak menemui
// error "Baris hari ini belum tersedia" di web app.
//
// Jika besok menyeberang ke bulan baru, sheet bulan itu mungkin belum ada
// (trigger buatSheetBulanBaru baru jalan tanggal 1 jam 05:00). Karena itu
// kita pastikan dulu sheet bulan target dibuat sebelum append.
function appendBesok() {
  _requireAdmin();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⚠ appendBesok: tidak bisa acquire lock dalam 30s — skip.');
    return;
  }

  try {
    const today    = getToday();
    const tomorrow = new Date(
      today.getFullYear(), today.getMonth(), today.getDate() + 1, 12, 0, 0
    );

    // Lintas bulan → pastikan sheet bulan besok ada lebih dulu.
    if (tomorrow.getMonth() !== today.getMonth()
        || tomorrow.getFullYear() !== today.getFullYear()) {
      const dibuat = _buatSheetBulanUntuk(tomorrow);
      Logger.log('appendBesok: siapkan sheet bulan baru →\n' + dibuat.join('\n'));
    }

    const r = _appendUntukTanggal(tomorrow);
    try {
      SpreadsheetApp.getUi().alert(
        '✅ Append besok selesai!\n' +
        r.tglStr + ' (' + r.namaHari + ')\n\n' +
        r.hasil.join('\n')
      );
    } catch(e) {
      // Dipanggil dari trigger — tidak ada UI
    }
  } finally {
    lock.releaseLock();
  }
}

// ── _appendUntukTanggal — Inti append baris untuk satu tanggal ────────
// Dipakai bersama oleh appendHariIni() (target = hari ini) dan appendBesok()
// (target = besok). Untuk tiap divisi:
//   1. Ambil daftar staf aktif dari Master_Data
//   2. Cek duplikat (skip jika tanggal target sudah ada di sheet)
//   3. Append baris kosong per staf dengan formula L–O
//   4. Pasang validasi dan proteksi baris baru
// CATATAN: tidak melakukan guard admin / lock — itu tanggung jawab pemanggil.
// Return: { tglStr, namaHari, hasil[] } untuk notifikasi UI.
function _appendUntukTanggal(targetDate) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) {
    Logger.log('❌ Master_Data tidak ditemukan.');
    return { tglStr: '', namaHari: '', hasil: ['❌ Master_Data tidak ditemukan.'] };
  }

  const today    = getToday();
  const isToday  = isSameDate(targetDate, today);
  const tglStr   = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const namaHari = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, 'EEEE');

  // Ambil semua staf aktif
  const masterData = master.getRange('A4:D200').getValues()
    .filter(r =>
      r[0] !== '' && r[1] !== '' &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    );

  if (masterData.length === 0) {
    Logger.log('Tidak ada staf aktif di Master_Data.');
    return { tglStr, namaHari, hasil: ['⚠ Tidak ada staf aktif di Master_Data'] };
  }

  // Kelompokkan per divisi (case-insensitive)
  const stafPerDivisi = {};
  for (const k of masterData) {
    const div = String(k[0]).trim().toUpperCase();
    if (!stafPerDivisi[div]) stafPerDivisi[div] = [];
    stafPerDivisi[div].push({
      nama : String(k[1]).trim(),
      email: String(k[2]).trim(),
    });
  }

  // Warna baris sesuai posisi tanggal: kuning kalau hari ini, muted kalau
  // bukan (mis. baris besok yang dibuat malam ini). highlightHariIni() akan
  // mewarnai ulang menjadi kuning saat tanggal itu benar-benar tiba.
  const bgLocked  = isToday ? '#FFF9C4' : '#F1EFE8';
  const bgEdit    = isToday ? '#FFF9C4' : '#FFFFFF';
  const bgFormula = isToday ? '#FFF9C4' : '#EEEDFE';

  const hasil = [];

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi, targetDate);
    if (!sheet) {
      hasil.push('⚠ ' + divisi + ': sheet tidak ditemukan');
      continue;
    }

    const staf = stafPerDivisi[divisi.toUpperCase()] || [];
    if (staf.length === 0) {
      hasil.push('⚠ ' + divisi + ': tidak ada staf aktif');
      continue;
    }

    // Konsolidasi proteksi hari lampau — HANYA saat append hari ini.
    // Untuk tanggal masa depan (appendBesok), "sebelum target" mencakup baris
    // hari ini yang TIDAK boleh ikut dikunci. Konsolidasi tetap berjalan harian
    // lewat appendHariIni jam 06:00.
    if (isToday) {
      const hapus = _bersihkanProteksiLama(sheet, today);
      if (hapus > 0) Logger.log('🗑 ' + divisi + ': hapus ' + hapus + ' proteksi hari lalu');
    }

    // Cegah duplikat — skip jika tanggal target sudah ada di sheet
    const existingData = sheet.getLastRow() > 3
      ? sheet.getRange(4, 1, sheet.getLastRow() - 3, 1).getValues()
      : [];

    const sudahAda = existingData.some(r => {
      const tgl = r[0];
      return tgl instanceof Date && isSameDate(tgl, targetDate);
    });

    if (sudahAda) {
      hasil.push('⚠ ' + divisi + ': ' + tglStr + ' sudah ada — skip');
      continue;
    }

    // Siapkan baris baru (semua kolom editable dikosongkan)
    const newRows = staf.map(s => [
      targetDate, // A: Tanggal
      namaHari,   // B: Hari
      s.nama,     // C: Nama
      s.email,    // D: Email
      '', '', '', '', '', '', '',  // E–K: diisi staf
      '', '', '', '',              // L–O: formula (diset di bawah)
      '', '',                      // P–Q: admin only
      '',                          // R: Keterangan tidak hadir
      '',                          // S: Plan (admin only)
      '',                          // T: Device Status
      '',                          // U: Catatan Telat
      '',                          // V: Catatan Pulang Awal
    ]);

    const insertAt = sheet.getLastRow() + 1;
    sheet.getRange(insertAt, 1, newRows.length, TOTAL_COL).setValues(newRows);

    // Format tanggal kolom A
    sheet.getRange(insertAt, 1, newRows.length, 1).setNumberFormat('DD/MM/YYYY');

    // Warna kolom
    sheet.getRange(insertAt, 1,  newRows.length, 4)
      .setBackground(bgLocked).setFontColor('#5F5E5A');  // A:D terkunci
    sheet.getRange(insertAt, 5,  newRows.length, 7)
      .setBackground(bgEdit).setFontColor('#2C2C2A');    // E:K editable
    sheet.getRange(insertAt, 12, newRows.length, 4)
      .setBackground(bgFormula).setFontColor('#534AB7').setFontWeight('bold'); // L:O formula
    sheet.getRange(insertAt, 16, newRows.length, 2)
      .setBackground(bgEdit).setFontColor('#2C2C2A');    // P:Q admin
    sheet.getRange(insertAt, COL_KETERANGAN, newRows.length, 1)
      .setBackground(bgEdit).setFontColor('#2C2C2A');    // R: Keterangan
    sheet.getRange(insertAt, COL_PLAN, newRows.length, 1)
      .setBackground(bgFormula).setFontColor('#085041').setFontWeight('bold'); // S: Plan
    sheet.getRange(insertAt, COL_DEVICE, newRows.length, 1)
      .setBackground(bgEdit).setFontColor('#2C2C2A');    // T: Device
    sheet.getRange(insertAt, COL_TELAT, newRows.length, 2)
      .setBackground(bgEdit).setFontColor('#E65100');    // U:V Catatan Telat/Pulang Awal

    // Border
    sheet.getRange(insertAt, 1, newRows.length, TOTAL_COL)
      .setBorder(true,true,true,true,true,true,
        '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);

    // Pasang formula per baris
    _pasangFormulaBaris(sheet, insertAt, newRows.length);

    // Format kolom jam sebagai [h]:mm
    sheet.getRange(insertAt, COL_EFEKTIF,     newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_REGULAR_JAM, newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT1,         newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT2,         newRows.length, 1).setNumberFormat('[h]:mm');

    setupValidasiBaris(sheet, insertAt, newRows.length);
    proteksiBarisBaru(sheet, divisi, insertAt, newRows.length);

    hasil.push('✓ ' + divisi + ' (' + sheet.getName() + '): ' +
               staf.length + ' staf — ' + tglStr);
    Logger.log('Append selesai: ' + divisi + ' — ' + tglStr);
  }

  groupByToday();
  highlightHariIni();

  Logger.log('_appendUntukTanggal selesai: ' + tglStr + '\n' + hasil.join('\n'));

  return { tglStr, namaHari, hasil };
}

// ── appendTanggalBaru — Append hari berikutnya setelah tanggal terakhir ─
// Dipanggil manual oleh admin dari menu ⚙ Admin.
// Untuk tiap sheet divisi: cari tanggal terakhir di kolom A,
// lalu append baris untuk tanggal berikutnya (lastDate + 1 hari).
// Berguna ketika trigger harian gagal atau admin perlu mengejar hari yang terlewat.
function appendTanggalBaru() {
  _requireAdmin();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { return; }

  // Lock: cegah eksekusi paralel dengan appendHariIni / instance kedua
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert('⚠ Append sedang berjalan. Coba lagi beberapa detik.');
    return;
  }

  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ambil staf aktif dari Master_Data
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) { ui.alert('❌ Sheet Master_Data tidak ditemukan.'); return; }

  const stafPerDivisi = {};
  master.getRange('A4:D200').getValues()
    .filter(r => r[0] !== '' && r[1] !== '' && String(r[3]).trim().toUpperCase() === 'TRUE')
    .forEach(r => {
      const div = String(r[0]).trim().toUpperCase();
      if (!stafPerDivisi[div]) stafPerDivisi[div] = [];
      stafPerDivisi[div].push({ nama: String(r[1]).trim(), email: String(r[2]).trim() });
    });

  // Step 1: preview — tentukan tanggal yang akan di-append per divisi
  const rencana = [];
  const preview = [];

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) { preview.push('⚠ ' + divisi + ': sheet tidak ditemukan'); continue; }

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) { preview.push('⚠ ' + divisi + ': belum ada data'); continue; }

    // Cari tanggal terakhir dari baris paling bawah ke atas
    const colA = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
    let lastDate = null;
    for (let i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] instanceof Date) { lastDate = colA[i][0]; break; }
    }
    if (!lastDate) { preview.push('⚠ ' + divisi + ': tidak ada tanggal valid'); continue; }

    // Tanggal baru = lastDate + 1 hari
    const nextDate = new Date(
      lastDate.getFullYear(), lastDate.getMonth(), lastDate.getDate() + 1, 12, 0, 0
    );
    const nextNorm = new Date(nextDate.getFullYear(), nextDate.getMonth(), nextDate.getDate());
    const nextStr  = Utilities.formatDate(nextDate, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const namaHari = Utilities.formatDate(nextDate, CONFIG.TIMEZONE, 'EEEE');

    // Jika nextDate menyeberang ke bulan baru, sheet target harus sheet bulan itu
    // — bukan sheet bulan lama yang sedang kita baca lastDate-nya.
    let targetSheet = sheet;
    if (nextDate.getMonth() !== lastDate.getMonth()
        || nextDate.getFullYear() !== lastDate.getFullYear()) {
      targetSheet = getSheetAktifDivisi(divisi, nextDate);
      if (!targetSheet || targetSheet === sheet) {
        const bulanTarget = Utilities.formatDate(nextDate, CONFIG.TIMEZONE, 'MMM yyyy');
        preview.push('⚠ ' + divisi + ': sheet bulan ' + bulanTarget +
                     ' belum ada — jalankan "Buat Sheet Bulan Baru" dulu');
        continue;
      }
    }

    // Cek duplikat — pakai data dari targetSheet (bisa beda sheet jika lintas bulan)
    const targetLastRow = targetSheet.getLastRow();
    const targetColA = targetLastRow > 3
      ? targetSheet.getRange(4, 1, targetLastRow - 3, 1).getValues()
      : [];
    const sudahAda = targetColA.some(r => {
      const d = r[0];
      if (!(d instanceof Date)) return false;
      return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() === nextNorm.getTime();
    });
    if (sudahAda) { preview.push('⚠ ' + divisi + ': ' + nextStr + ' sudah ada'); continue; }

    const staf = stafPerDivisi[divisi.toUpperCase()] || [];
    if (staf.length === 0) { preview.push('⚠ ' + divisi + ': tidak ada staf aktif'); continue; }

    rencana.push({ divisi, sheet: targetSheet, nextDate, namaHari, nextStr, staf });
    preview.push('✓ ' + divisi + ' (' + targetSheet.getName() + '): append ' +
                 nextStr + ' (' + namaHari + ') — ' + staf.length + ' staf');
  }

  if (rencana.length === 0) {
    ui.alert('⚠ Tidak ada yang perlu di-append.\n\n' + preview.join('\n'));
    return;
  }

  // Konfirmasi sebelum eksekusi
  const konfirmasi = ui.alert(
    '📅 Konfirmasi Append Tanggal Baru',
    preview.join('\n') + '\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi !== ui.Button.YES) return;

  // Step 2: eksekusi append per divisi
  const hasil = [];
  const today  = getToday();

  for (const { divisi, sheet, nextDate, namaHari, nextStr, staf } of rencana) {
    const auto = (CONFIG.AUTO_ABSENSI || {})[divisi.toUpperCase()] || null;

    const newRows = staf.map(s => [
      nextDate,
      namaHari,
      s.nama,
      s.email,
      auto ? (auto.status      || '') : '',
      auto ? (auto.masuk       || '') : '',
      auto ? (auto.ist1Mulai   || '') : '',
      auto ? (auto.ist1Selesai || '') : '',
      auto ? (auto.ist2Mulai   || '') : '',
      auto ? (auto.ist2Selesai || '') : '',
      auto ? (auto.pulang      || '') : '',
      '', '', '', '',  // L–O: formula
      '', '',          // P–Q: admin only
      '',              // R: Keterangan tidak hadir
      '',              // S: Plan (admin only)
      '',              // T: Device Status
      '',              // U: Catatan Telat
      '',              // V: Catatan Pulang Awal
    ]);

    const insertAt = sheet.getLastRow() + 1;
    sheet.getRange(insertAt, 1, newRows.length, TOTAL_COL).setValues(newRows);
    sheet.getRange(insertAt, 1, newRows.length, 1).setNumberFormat('DD/MM/YYYY');

    if (auto) {
      [COL_MASUK, COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_PULANG]
        .forEach(col => sheet.getRange(insertAt, col, newRows.length, 1).setNumberFormat('HH:mm'));
    }

    // Warna sesuai posisi tanggal relatif terhadap hari ini
    const isToday   = isSameDate(nextDate, today);
    const bgLocked  = isToday ? '#FFF9C4' : '#F1EFE8';
    const bgEdit    = isToday ? '#FFF9C4' : '#FFFFFF';
    const bgFormula = isToday ? '#FFF9C4' : '#EEEDFE';

    sheet.getRange(insertAt, 1,  newRows.length, 4).setBackground(bgLocked).setFontColor('#5F5E5A');
    sheet.getRange(insertAt, 5,  newRows.length, 7).setBackground(bgEdit).setFontColor('#2C2C2A');
    sheet.getRange(insertAt, 12, newRows.length, 4).setBackground(bgFormula).setFontColor('#534AB7').setFontWeight('bold');
    sheet.getRange(insertAt, 16, newRows.length, 2).setBackground(bgEdit).setFontColor('#2C2C2A');
    sheet.getRange(insertAt, COL_KETERANGAN, newRows.length, 1).setBackground(bgEdit).setFontColor('#2C2C2A');
    sheet.getRange(insertAt, COL_TELAT, newRows.length, 2)
      .setBackground(bgEdit).setFontColor('#E65100');  // U:V Catatan Telat/Pulang Awal
    sheet.getRange(insertAt, 1,  newRows.length, TOTAL_COL)
      .setBorder(true,true,true,true,true,true,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);

    _pasangFormulaBaris(sheet, insertAt, newRows.length);
    sheet.getRange(insertAt, COL_EFEKTIF,     newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_REGULAR_JAM, newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT1,         newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT2,         newRows.length, 1).setNumberFormat('[h]:mm');

    setupValidasiBaris(sheet, insertAt, newRows.length);
    proteksiBarisBaru(sheet, divisi, insertAt, newRows.length);

    hasil.push('✓ ' + divisi + ': ' + nextStr + ' (' + namaHari + ') — ' + staf.length + ' staf');
    Logger.log('appendTanggalBaru: ' + divisi + ' — ' + nextStr);
  }

  groupByToday();
  highlightHariIni();
  ui.alert('✅ Append selesai!\n\n' + hasil.join('\n'));
  } finally {
    lock.releaseLock();
  }
}

// ── highlightHariIni — Warnai baris berdasarkan tanggal ───────────────
// Kuning = hari ini,  Abu = sudah lewat
function highlightHariIni() {
  const today = getToday();

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) continue;

    const data    = sheet.getDataRange().getValues();
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    for (let r = 4; r <= lastRow; r++) {
      const val = data[r - 1][0];
      if (!val || !(val instanceof Date)) continue;

      // Hari Minggu → seluruh baris merah. Dicek paling awal supaya menimpa
      // warna kuning (hari ini) / abu (lewat). Kolom B menyimpan 'Sunday'.
      const hari = String(data[r - 1][1]).trim();
      if (hari === 'Sunday') {
        sheet.getRange(r, 1, 1, TOTAL_COL)
          .setBackground('#FFCDD2').setFontColor('#B71C1C');
        continue;
      }

      if (isSameDate(val, today)) {
        sheet.getRange(r, 1,  1, 4).setBackground('#FFF9C4');
        sheet.getRange(r, 5,  1, 7).setBackground('#FFF9C4');
        sheet.getRange(r, 12, 1, 1).setBackground('#FFF9C4');
        sheet.getRange(r, 13, 1, 5).setBackground('#FFF9C4');
      } else if (isPast(val, today)) {
        sheet.getRange(r, 1,  1, 4).setBackground('#F1EFE8');
        sheet.getRange(r, 5,  1, 7).setBackground('#F8F8F8');
        sheet.getRange(r, 12, 1, 1).setBackground('#EEEDFE');
        sheet.getRange(r, 13, 1, 5).setBackground('#F8F8F8');
      }
    }
  }
  Logger.log('Highlight selesai.');
}

// ── groupByToday — Collapse baris lama, buka baris hari ini ──────────
// TargetSheet opsional — jika null, proses semua sheet divisi aktif
function groupByToday(targetSheet) {
  const today  = getToday();
  const sheets = targetSheet
    ? [targetSheet]
    : CONFIG.DIVISI.map(d => getSheetAktifDivisi(d)).filter(s => s);

  for (const sheet of sheets) {
    const data    = sheet.getDataRange().getValues();
    const lastCol = sheet.getLastColumn();

    // Reset grouping lama
    try { sheet.expandAllRowGroups(); } catch(e) {}
    try {
      const rng = sheet.getDataRange();
      for (let i = 0; i < 3; i++) {
        try { rng.shiftRowGroupDepth(-1); } catch(e) { break; }
      }
    } catch(e) {}

    // Temukan baris pertama hari ini
    let firstRowToday = -1;
    for (let i = 3; i < data.length; i++) {
      const val = data[i][0];
      if (!val || !(val instanceof Date)) continue;
      if (isSameDate(val, today)) { firstRowToday = i + 1; break; }
    }

    const groupEnd = firstRowToday === -1
      ? sheet.getLastRow()
      : firstRowToday - 1;

    if (groupEnd >= 4) {
      const groupStart  = 4;
      const groupLength = groupEnd - groupStart + 1;
      sheet.getRange(groupStart, 1, groupLength, lastCol).shiftRowGroupDepth(1);
    }

    try { sheet.collapseAllRowGroups(); } catch(e) {}

    Logger.log(
      sheet.getName() + ': group baris 4–' + groupEnd +
      (firstRowToday > -1 ? ', hari ini mulai baris ' + firstRowToday : ' (semua)')
    );
  }
}

// ── _pasangFormulaBaris — Private: set formula L, M, N, O ─────────────
// Dipanggil oleh appendHariIni() dan generateFullMonth()
function _pasangFormulaBaris(sheet, startRow, numRows) {
  for (let r = startRow; r < startRow + numRows; r++) {
    const a=`A${r}`, b=`B${r}`, e=`E${r}`, f=`F${r}`,
          g=`G${r}`, h=`H${r}`, i=`I${r}`, j=`J${r}`,
          k=`K${r}`, l=`L${r}`, p=`P${r}`, s=`S${r}`, q=`Q${r}`;

    // Durasi K-F dengan dukungan lintas tengah malam:
    // jika pulang (K) < masuk (F), berarti clock-out di hari berikutnya → tambah 1 (24 jam)
    const kf = `IF(${k}<${f},${k}+1-${f},${k}-${f})`;

    // L: Jam Efektif (fraction hari)
    // Aturan khusus Minggu: walau status="Hadir", jika PLAN (S) kosong → 0
    // (mencegah jam efektif terhitung tanpa plan kerja eksplisit di hari libur)
    sheet.getRange(r, COL_EFEKTIF).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(AND(${b}="Sunday",${q}=""),0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF(AND(${g}<>"",${h}<>""),` +
        `IF(AND(${i}<>"",${j}<>""),` +
          `${kf}-(${h}-${g})-(${j}-${i}),` +
          `${kf}-(${h}-${g})),` +
      `IF(AND(${i}<>"",${j}<>""),` +
          `${kf}-(${j}-${i}),` +
          `${kf})))))`
    );

    // M: Regular Hours
    // Precedence:
    //  1. Status="Red Day" → REGULAR_DAYS jam (tetap, tidak peduli hari)
    //  2. NOTE ∈ paid-off list (RED DAY/VACATION PAID/FLEX DAY/dst) → REGULAR_DAYS jam (full)
    //  3. Saturday normal → cap REGULAR_DAYS jika Jam Efektif >= SATURDAY, else Jam Efektif
    //  4. Weekday normal → cap REGULAR_DAYS jika Jam Efektif >= REGULAR_DAYS, else Jam Efektif
    sheet.getRange(r, COL_REGULAR_JAM).setFormula(
      `=IF(${e}="Red Day",${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
      `IF(OR(${p}="RED DAY",${p}="RED DAY DOUBLE",${p}="SAVING DAY RED DAY/SUNDAY",${p}="SWAP RED DAY",${p}="VACATION PAID",${p}="FLEX DAY",${p}="ADDITIONAL PAID",${p}="MATERNITY LEAVE",${p}="SICK PAID"),` +
        `IF(${b}="Saturday",${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24),` +
      `IF(${b}="Saturday",` +
        `IF(${l}>=${CONFIG.DAYS_HOUR.SATURDAY}/24,` +
          `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l}),` +
      `IF(${l}>=${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
        `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l}))))`
    );

    // N: OT 1 (maks 1 jam di atas regular)
    // Sama dengan L: Minggu tanpa PLAN → 0 (tidak hitung OT tanpa plan eksplisit)
    sheet.getRange(r, COL_OT1).setFormula(
      `=IF(${e}<>"Hadir",0,IF(AND(${b}="Sunday",${q}=""),0,IF(OR(${f}="",${k}=""),0,IF((${kf}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))<=IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24,0,MIN(1/24,(${kf}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))-IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24)))))`
    );

    // O: OT 2 (di atas OT 1)
    // Sama dengan L/N: Minggu tanpa PLAN → 0
    sheet.getRange(r, COL_OT2).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(AND(${b}="Sunday",${q}=""),0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF((${kf}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))` +
        `<=(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24,` +
      `0,` +
      `${kf}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24))))`
    );
  }
}

// ── perbaruiRumusSemuaBaris — Pasang ulang rumus L, M, N, O ke baris lama ─
// Dipanggil manual oleh admin dari menu ⚙ Admin ketika rumus di Append.js
// diperbarui (misal: dukungan lintas tengah malam, aturan Minggu, dll).
// Operasi hanya mengganti rumus — data jam (E–K) dan kolom lain tidak disentuh.
//
// Cakupan:
//   - Sheet aktif (yang sedang dibuka admin), ATAU
//   - Semua sheet divisi (semua bulan) — sesuai pilihan admin
function perbaruiRumusSemuaBaris() {
  _requireAdmin();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { return; }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert('⚠ Sistem sedang sibuk. Coba lagi beberapa detik.');
    return;
  }

  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const activeName  = activeSheet.getName();

    // Pilih scope: sheet aktif saja, atau semua sheet divisi
    const scopeRes = ui.alert(
      '🔄 Perbarui Rumus L, M, N, O',
      'Pilih cakupan update:\n\n' +
      '• YES  → SEMUA sheet divisi (semua bulan)\n' +
      '• NO   → Hanya sheet aktif: "' + activeName + '"\n' +
      '• CANCEL → Batal',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (scopeRes === ui.Button.CANCEL || scopeRes === ui.Button.CLOSE) return;

    // Kumpulkan sheet target
    const targets = [];
    if (scopeRes === ui.Button.YES) {
      const divisiPrefixes = CONFIG.DIVISI.map(d => d.toUpperCase());
      for (const s of ss.getSheets()) {
        const nameUp  = s.getName().toUpperCase();
        const lastRow = s.getLastRow();
        if (lastRow < 4) continue;
        const matched = divisiPrefixes.some(p => nameUp === p || nameUp.startsWith(p + '_'));
        if (matched) targets.push({ sheet: s, rows: lastRow - 3 });
      }
    } else {
      const lastRow = activeSheet.getLastRow();
      if (lastRow < 4) { ui.alert('Sheet aktif belum punya data.'); return; }
      // Validasi: sheet aktif harus sheet divisi
      const nameUp = activeName.toUpperCase();
      const matched = CONFIG.DIVISI.map(d => d.toUpperCase())
        .some(p => nameUp === p || nameUp.startsWith(p + '_'));
      if (!matched) {
        ui.alert('Sheet aktif "' + activeName + '" bukan sheet divisi.');
        return;
      }
      targets.push({ sheet: activeSheet, rows: lastRow - 3 });
    }

    if (targets.length === 0) {
      ui.alert('Tidak ada sheet divisi yang ditemukan.');
      return;
    }

    const totalRows = targets.reduce((sum, t) => sum + t.rows, 0);
    const preview   = targets.map(t => '• ' + t.sheet.getName() + ': ' + t.rows + ' baris').join('\n');

    const konfirmasi = ui.alert(
      'Konfirmasi Perbarui Rumus',
      'Akan memperbarui rumus L, M, N, O untuk ' + totalRows +
        ' baris di ' + targets.length + ' sheet:\n\n' +
      preview + '\n\n' +
      'Data jam masuk/pulang/istirahat TIDAK akan berubah.\n' +
      'Lanjutkan?',
      ui.ButtonSet.YES_NO
    );
    if (konfirmasi !== ui.Button.YES) return;

    // Eksekusi — pasang rumus per sheet
    const hasil = [];
    for (const t of targets) {
      try {
        _pasangFormulaBaris(t.sheet, 4, t.rows);
        // Re-apply format jam — jaga konsisten
        t.sheet.getRange(4, COL_EFEKTIF,     t.rows, 1).setNumberFormat('[h]:mm');
        t.sheet.getRange(4, COL_REGULAR_JAM, t.rows, 1).setNumberFormat('[h]:mm');
        t.sheet.getRange(4, COL_OT1,         t.rows, 1).setNumberFormat('[h]:mm');
        t.sheet.getRange(4, COL_OT2,         t.rows, 1).setNumberFormat('[h]:mm');
        hasil.push('✓ ' + t.sheet.getName() + ': ' + t.rows + ' baris');
        Logger.log('perbaruiRumusSemuaBaris: ' + t.sheet.getName() + ' — ' + t.rows + ' baris');
      } catch(e) {
        hasil.push('❌ ' + t.sheet.getName() + ': ' + e.message);
        Logger.log('perbaruiRumusSemuaBaris error pada ' + t.sheet.getName() + ': ' + e.message);
      }
    }

    ui.alert('✅ Perbarui rumus selesai!\n\n' + hasil.join('\n'));
  } finally {
    lock.releaseLock();
  }
}

// ── _bersihkanProteksiLama — Konsolidasi proteksi hari lalu ──────────
// Dipanggil di awal appendHariIni() sebelum menambah baris baru.
// Strategi:
//   1. Temukan batas baris hari lalu (baris 4 s/d firstTodayRow-1)
//   2. Hapus semua range protection individual di area hari lalu
//   3. Buat SATU proteksi konsolidasi yang mengcover seluruh area hari lalu
//      → Baris hari lalu tetap terkunci (owner + admin only), hemat kuota protect
// Header (baris 1–3) dan proteksi hari ini tidak disentuh.
// Return: jumlah proteksi individual yang dihapus (untuk logging).
function _bersihkanProteksiLama(sheet, today) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return 0;

  // Temukan baris pertama yang merupakan hari ini
  const dates = sheet.getRange(4, COL_TANGGAL, lastRow - 3, 1).getValues();
  let firstTodayRow = lastRow + 1; // default: belum ada baris hari ini
  for (let i = 0; i < dates.length; i++) {
    if (dates[i][0] instanceof Date && isSameDate(dates[i][0], today)) {
      firstTodayRow = i + 4; // 1-indexed
      break;
    }
  }

  const lastPastRow = firstTodayRow - 1;
  if (lastPastRow < 4) return 0; // tidak ada baris hari lalu sama sekali

  // Hapus semua range protection yang seluruhnya berada di area hari lalu
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let removed = 0;
  for (const prot of protections) {
    const range  = prot.getRange();
    const pStart = range.getRow();
    const pEnd   = range.getLastRow();

    if (pEnd <= 3) continue;                    // lewati header
    if (pStart >= 4 && pEnd <= lastPastRow) {   // seluruhnya di area hari lalu
      prot.remove();
      removed++;
    }
  }

  // Ganti dengan SATU proteksi konsolidasi (owner + admin only)
  // → baris hari lalu tetap read-only meski grup dibuka
  const owner     = Session.getEffectiveUser();
  const pastRange = sheet.getRange(4, 1, lastPastRow - 3, TOTAL_COL);
  const newProt   = pastRange.protect();
  newProt.setDescription('Hari lalu — terkunci (konsolidasi)');
  newProt.setWarningOnly(false);
  newProt.removeEditors(newProt.getEditors());
  newProt.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { newProt.addEditor(adminEmail); } catch(e) {}
  }

  return removed;
}
