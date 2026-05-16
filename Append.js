// ═══════════════════════════════════════════════════════════════════════
// APPEND.JS — Penambahan baris harian + tampilan sheet
// TEST 3
// Berisi:
//   appendHariIni()    — inti fungsi harian: tambah baris staf ke sheet divisi
//   highlightHariIni() — warnai baris hari ini (kuning) dan lewat (abu)
//   groupByToday()     — collapse baris hari lama, buka hari ini
// ═══════════════════════════════════════════════════════════════════════

// ── appendHariIni — Core function harian ──────────────────────────────
// Dipanggil otomatis pukul 06:00 via trigger, atau manual oleh HRD.
// Untuk setiap divisi:
//   1. Ambil daftar staf aktif dari Master_Data
//   2. Cek duplikat (skip jika hari ini sudah ada)
//   3. Append baris kosong per staf dengan formula L–O
//   4. Pasang validasi dan proteksi baris baru
function appendHariIni() {
  _requireAdmin();

  // Lock: cegah eksekusi paralel (trigger 06:00 + admin manual klik bersamaan)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⚠ appendHariIni: tidak bisa acquire lock dalam 30s — skip.');
    return;
  }

  try {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) {
    Logger.log('❌ Master_Data tidak ditemukan.');
    return;
  }

  const today    = getToday();
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const namaHari = Utilities.formatDate(today, CONFIG.TIMEZONE, 'EEEE');

  // Ambil semua staf aktif
  const masterData = master.getRange('A4:D200').getValues()
    .filter(r =>
      r[0] !== '' && r[1] !== '' &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    );

  if (masterData.length === 0) {
    Logger.log('Tidak ada staf aktif di Master_Data.');
    return;
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

  const hasil = [];

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) {
      hasil.push('⚠ ' + divisi + ': sheet tidak ditemukan');
      continue;
    }

    const staf = stafPerDivisi[divisi.toUpperCase()] || [];
    if (staf.length === 0) {
      hasil.push('⚠ ' + divisi + ': tidak ada staf aktif');
      continue;
    }

    // Hapus proteksi hari lalu sebelum append — hemat kuota protect Google Sheets
    const hapus = _bersihkanProteksiLama(sheet, today);
    if (hapus > 0) Logger.log('🗑 ' + divisi + ': hapus ' + hapus + ' proteksi hari lalu');

    // Cegah duplikat — skip jika hari ini sudah ada di sheet
    const existingData = sheet.getLastRow() > 3
      ? sheet.getRange(4, 1, sheet.getLastRow() - 3, 1).getValues()
      : [];

    const sudahAda = existingData.some(r => {
      const tgl = r[0];
      return tgl instanceof Date && isSameDate(tgl, today);
    });

    if (sudahAda) {
      hasil.push('⚠ ' + divisi + ': hari ini sudah ada — skip');
      continue;
    }

    // Siapkan baris baru (semua kolom editable dikosongkan)
    const newRows = staf.map(s => [
      today,    // A: Tanggal
      namaHari, // B: Hari
      s.nama,   // C: Nama
      s.email,  // D: Email
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
      .setBackground('#FFF9C4').setFontColor('#5F5E5A');  // A:D terkunci
    sheet.getRange(insertAt, 5,  newRows.length, 7)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // E:K editable
    sheet.getRange(insertAt, 12, newRows.length, 4)
      .setBackground('#FFF9C4').setFontColor('#534AB7').setFontWeight('bold'); // L:O formula
    sheet.getRange(insertAt, 16, newRows.length, 2)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // P:Q admin
    sheet.getRange(insertAt, COL_KETERANGAN, newRows.length, 1)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // R: Keterangan
    sheet.getRange(insertAt, COL_PLAN, newRows.length, 1)
      .setBackground('#FFF9C4').setFontColor('#085041').setFontWeight('bold'); // S: Plan
    sheet.getRange(insertAt, COL_DEVICE, newRows.length, 1)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');                       // T: Device
    sheet.getRange(insertAt, COL_TELAT, newRows.length, 2)
      .setBackground('#FFF9C4').setFontColor('#E65100');                       // U:V Catatan Telat/Pulang Awal

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
               staf.length + ' staf — ' + todayStr);
    Logger.log('Append selesai: ' + divisi + ' — ' + todayStr);
  }

  groupByToday();
  highlightHariIni();

  Logger.log('appendHariIni selesai: ' + todayStr + '\n' + hasil.join('\n'));

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Append hari ini selesai!\n' +
      todayStr + ' (' + namaHari + ')\n\n' +
      hasil.join('\n')
    );
  } catch(e) {
    // Dipanggil dari trigger — tidak ada UI
  }
  } finally {
    lock.releaseLock();
  }
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
          k=`K${r}`, l=`L${r}`, p=`P${r}`;

    // L: Jam Efektif (fraction hari)
    sheet.getRange(r, COL_EFEKTIF).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF(AND(${g}<>"",${h}<>""),` +
        `IF(AND(${i}<>"",${j}<>""),` +
          `${k}-${f}-(${h}-${g})-(${j}-${i}),` +
          `${k}-${f}-(${h}-${g})),` +
      `IF(AND(${i}<>"",${j}<>""),` +
          `${k}-${f}-(${j}-${i}),` +
          `${k}-${f}))))`
    );

    // M: Regular Hours
    // Jika NOTE mengandung "PAID" (mis. VACATION PAID, SICK PAID), langsung 7 jam
    sheet.getRange(r, COL_REGULAR_JAM).setFormula(
      `=IF(AND(ISNUMBER(SEARCH("PAID",${p})),NOT(ISNUMBER(SEARCH("UNPAID",${p})))),${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
      `IF(${b}="Saturday",` +
        `IF(${l}>=${CONFIG.DAYS_HOUR.SATURDAY}/24,` +
          `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l}),` +
      `IF(${l}>=${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
        `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l})))`
    );

    // N: OT 1 (maks 1 jam di atas regular)
    sheet.getRange(r, COL_OT1).setFormula(
      `=IF(${e}<>"Hadir",0,IF(OR(${f}="",${k}=""),0,IF((${k}-${f}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))<=IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24,0,MIN(1/24,(${k}-${f}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))-IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24))))`
    );

    // O: OT 2 (di atas OT 1)
    sheet.getRange(r, COL_OT2).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF((${k}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))` +
        `<=(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24,` +
      `0,` +
      `${k}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24)))`
    );
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
