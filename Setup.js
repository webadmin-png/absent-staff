// ═══════════════════════════════════════════════════════════════════════
// SETUP.JS — Inisialisasi struktur sheet dan proteksi
//
// Berisi:
//   buatSheetBulanBaru()   — buat sheet kosong untuk bulan berjalan
//   setupProteksiMaster()  — kunci sheet Master_Data
//   setupValidasiBaris()   — pasang dropdown & validasi format jam per baris baru
//   setupValidasi()        — pasang validasi ke seluruh sheet yang sudah ada data
//   proteksiBarisBaru()    — proteksi range E:O per staf + P:Q khusus admin
// ═══════════════════════════════════════════════════════════════════════

// ── buatSheetBulanBaru — Buat sheet divisi bulan ini ──────────────────
// Buat struktur header (baris 1–3) + formatting untuk setiap divisi.
// Baris data (staf) akan diisi oleh appendHariIni() setelah sheet dibuat.
function buatSheetBulanBaru() {
  _requireAdmin();
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const now       = new Date();
  const namaBulan = Utilities.formatDate(now, CONFIG.TIMEZONE, 'MMM_yyyy');
  const hasil     = [];

  for (const divisi of CONFIG.DIVISI) {
    const namaSheet = divisi + '_' + namaBulan;

    if (ss.getSheetByName(namaSheet)) {
      hasil.push('⚠ ' + namaSheet + ' sudah ada — skip');
      continue;
    }

    const sheet = ss.insertSheet(namaSheet);
    sheet.setTabColor('#1D9E75');
    sheet.setHiddenGridlines(true);

    // Baris 1: Judul
    sheet.getRange(1, 1, 1, TOTAL_COL).merge()
      .setValue('ABSENSI ' + divisi + ' — ' +
        Utilities.formatDate(now, CONFIG.TIMEZONE, 'MMMM yyyy').toUpperCase())
      .setBackground('#0F6E56').setFontColor('#FFFFFF')
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);

    // Baris 2: Legenda warna
    const legends = [
      [1,  4, 'ABU = sudah lewat',    '#F1EFE8', '#5F5E5A'],
      [5,  4, 'PUTIH = bisa diedit',  '#FFFFFF',  '#2C2C2A'],
      [9,  4, 'UNGU = formula auto',  '#EEEDFE',  '#534AB7'],
      [13, 10, 'KUNING = hari ini',    '#FFF9C4',  '#633806'],
    ];
    for (const [startCol, span, text, bg, fg] of legends) {
      sheet.getRange(2, startCol, 1, span).merge()
        .setValue(text).setBackground(bg).setFontColor(fg)
        .setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true,true,true,true,false,false,
          '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(2, 16);

    // Baris 3: Header kolom
    const headers = [
      ['Tanggal',                           '#1D9E75', '#FFFFFF'],
      ['Hari',                              '#1D9E75', '#FFFFFF'],
      ['Nama',                              '#1D9E75', '#FFFFFF'],
      ['Email',                             '#1D9E75', '#FFFFFF'],
      ['Status ▾',                          '#E1F5EE', '#085041'],
      ['Masuk',                             '#E1F5EE', '#085041'],
      ['Ist. Pertama\nMulai',               '#E1F5EE', '#085041'],
      ['Ist. Pertama\nSelesai',             '#E1F5EE', '#085041'],
      ['Ist. Kedua\nMulai',                 '#E1F5EE', '#085041'],
      ['Ist. Kedua\nSelesai',               '#E1F5EE', '#085041'],
      ['Pulang',                            '#E1F5EE', '#085041'],
      ['Jam Efektif 🔒',                    '#1D9E75', '#FFFFFF'],
      ['Regular Hours',                     '#E1F5EE', '#085041'],
      ['OT 1',                              '#E1F5EE', '#085041'],
      ['OT 2',                              '#E1F5EE', '#085041'],
      ['NOTE',                              '#E1F5EE', '#085041'],
      ['SUNDAY/RED DAY\nFILL: DOUBLE/SWAP', '#E1F5EE', '#085041'],
      ['KETERANGAN\nTIDAK HADIR',           '#E1F5EE', '#085041'],
      ['PLAN',                              '#1D9E75', '#FFFFFF'],
      ['UPC / PC\nSTATUS',                  '#E1F5EE', '#085041'],
      ['CATATAN\nTELAT',                    '#FFF3E0', '#E65100'],
      ['CATATAN\nPULANG AWAL',              '#FFF3E0', '#E65100'],
    ];
    for (let col = 0; col < headers.length; col++) {
      const [text, bg, fg] = headers[col];
      sheet.getRange(3, col + 1)
        .setValue(text).setBackground(bg).setFontColor(fg)
        .setFontWeight('bold').setFontSize(9)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true,true,true,true,false,false,
          '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(3, 44);

    // Lebar kolom A–R
    const colWidths = [90,80,130,180,70,70,90,90,90,90,70,100,60,60,60,120,140,180,120,200,180,180];
    colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(3);

    // Kunci header — hanya owner yang bisa edit
    const headerProt = sheet.getRange(1, 1, 3, TOTAL_COL).protect();
    headerProt.setDescription('Header — hanya owner yang bisa edit');
    headerProt.setWarningOnly(false);
    headerProt.removeEditors(headerProt.getEditors());
    headerProt.addEditor(Session.getEffectiveUser());

    hasil.push('✓ ' + namaSheet + ' dibuat (kosong — baris diisi appendHariIni)');
  }

  appendHariIni();

  const msg = '✅ Sheet bulan baru selesai!\n\n' +
    hasil.join('\n') + '\n\n' +
    'Baris hari ini sudah di-append otomatis.\n' +
    'Trigger harian akan append baris setiap pagi 06:00.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ── setupProteksiMaster — Kunci sheet Master_Data ─────────────────────
// Hapus proteksi lama lalu kunci seluruh sheet untuk owner + admin
// (ADMIN_EMAILS diambil dinamis dari spreadsheet Settings via _loadSettings).
function setupProteksiMaster() {
  _loadSettings();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) return;

  master.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(p => p.remove());

  const prot = master.protect();
  prot.setDescription('Hanya HRD/owner yang bisa edit Master_Data');
  prot.setWarningOnly(false);
  prot.removeEditors(prot.getEditors());
  prot.addEditor(Session.getEffectiveUser());
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { prot.addEditor(adminEmail); } catch(e) {}
  }
  Logger.log('Proteksi Master_Data selesai dengan ' +
             CONFIG.ADMIN_EMAILS.length + ' admin.');
}

// ── setupValidasiBaris — Pasang dropdown & validasi per baris baru ────
// Dipanggil setiap kali ada baris baru di-append atau sheet di-generate
function setupValidasiBaris(sheet, startRow, numRows) {
  if (!sheet || numRows <= 0) return;

  // E: Status (dropdown)
  sheet.getRange(startRow, COL_STATUS, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Hadir', 'Sakit', 'Izin', 'Alpha', 'Red Day'], true)
      .setHelpText('Pilih: Hadir / Sakit / Izin / Alpha / Red Day')
      .setAllowInvalid(false).build()
  );

  // F: Masuk — wajib diisi dengan format HH:MM
  sheet.getRange(startRow, COL_MASUK, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=ISNUMBER(TIMEVALUE(F' + startRow + '))')
      .setHelpText('Format jam: HH:MM — contoh 07:30')
      .setAllowInvalid(false).build()
  );

  // G–K: opsional, jika diisi harus HH:MM
  ['G','H','I','J','K'].forEach((col, idx) => {
    sheet.getRange(startRow, COL_IST1_M + idx, numRows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied(
          '=OR(' + col + startRow + '="",ISNUMBER(TIMEVALUE(' + col + startRow + ')))'
        )
        .setHelpText('Kosongkan jika tidak ada. Format: HH:MM')
        .setAllowInvalid(false).build()
    );
  });

  // P: NOTE (admin only — dropdown)
  sheet.getRange(startRow, COL_NOTE, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'HALF DAY', 'HALF DAY RED DAY', 'RED DAY', 'RED DAY DOUBLE',
        'SAVING DAY RED DAY/SUNDAY', 'SWAP RED DAY', 'VACATION PAID',
        'FLEX DAY', 'ADDITIONAL PAID', 'MATERNITY LEAVE',
        'SICK PAID', 'SICK UNPAID', 'DAY OFF UNPAID',
      ], true)
      .setHelpText('Pilih jenis keterangan hari khusus')
      .setAllowInvalid(false).build()
  );

  // Q: SUNDAY/RED DAY (admin only — dropdown)
  sheet.getRange(startRow, COL_SUNDAY, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['SWAP', 'DOUBLE', 'HALF DAY SUNDAY'], true)
      .setHelpText('Pilih: SWAP / DOUBLE / HALF DAY SUNDAY')
      .setAllowInvalid(false).build()
  );

  // S: PLAN (admin only — dropdown dari CONFIG.PLAN_JAM)
  sheet.getRange(startRow, COL_PLAN, numRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.PLAN_JAM, true)
      .setHelpText('Pilih plan jam kerja hari ini')
      .setAllowInvalid(true).build()
  );
}

// ── setupValidasi — Pasang validasi ke seluruh sheet ──────────────────
// Dipakai untuk sheet yang sudah ada datanya (retroactive)
function setupValidasi(sheet) {
  if (!sheet) return;
  const lastRow  = sheet.getLastRow();
  const dataRows = lastRow - 3;
  if (dataRows <= 0) return;
  setupValidasiBaris(sheet, 4, dataRows);
  Logger.log('Validasi selesai: ' + sheet.getName());
}

// ── proteksiBarisBaru — Proteksi range per staf + admin-only P:Q ──────
// Langkah:
//   1. Hapus proteksi range lama yang overlap dengan baris baru
//   2. Buat proteksi E:O per staf (hanya email staf + owner yang bisa edit)
//   3. Buat proteksi P:Q khusus admin (owner + ADMIN_EMAILS)
function proteksiBarisBaru(sheet, divisi, startRow, numRows) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) return;

  const owner  = Session.getEffectiveUser();
  const endRow = startRow + numRows - 1;

  // Step 1: Hapus proteksi lama yang overlap
  const existingProtRanges = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (const prot of existingProtRanges) {
    const pr  = prot.getRange();
    const ps  = pr.getRow();
    const pe  = pr.getLastRow();
    if (ps <= endRow && pe >= startRow) {
      prot.remove();
      Logger.log('🗑 Hapus proteksi lama: baris ' + ps + '–' + pe);
    }
  }

  // Step 2: Ambil staf divisi dari Master_Data
  const masterData = master.getRange('A4:D200').getValues()
    .filter(r =>
      r[0] !== '' &&
      String(r[0]).trim().toUpperCase() === divisi.trim().toUpperCase() &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    );

  if (masterData.length === 0) {
    Logger.log('⚠ Tidak ada staf aktif untuk divisi: ' + divisi);
    return;
  }

  // Step 3: Kelompokkan nomor baris per nama staf
  const newData       = sheet.getRange(startRow, 1, numRows, TOTAL_COL).getValues();
  const barisPerOrang = {};
  for (let i = 0; i < newData.length; i++) {
    const nama = String(newData[i][COL_NAMA - 1]).trim();
    if (!nama) continue;
    if (!barisPerOrang[nama]) barisPerOrang[nama] = [];
    barisPerOrang[nama].push(startRow + i);
  }

  // Step 4: Proteksi E:O per staf (staf hanya bisa edit barisnya sendiri)
  let berhasil = 0;
  for (const k of masterData) {
    const nama  = String(k[1]).trim();
    const email = String(k[2]).trim();
    if (!nama || !email) continue;

    const baris = barisPerOrang[nama];
    if (!baris || baris.length === 0) {
      Logger.log('⚠ ' + nama + ': tidak ada baris di range ' + startRow + '–' + endRow);
      continue;
    }

    const range = sheet.getRange(
      baris[0], COL_STATUS,
      baris.length, COL_OT2 - COL_STATUS + 1  // E sampai O
    );

    const prot = range.protect();
    prot.setDescription(nama + ' — ' + email +
      ' (baris ' + baris[0] + '–' + baris[baris.length - 1] + ')');
    prot.setWarningOnly(false);
    prot.removeEditors(prot.getEditors());
    prot.addEditor(owner);

    try {
      prot.addEditor(email);
      berhasil++;
      Logger.log('✓ Proteksi: ' + nama + ' baris ' + baris[0] + '–' + baris[baris.length - 1]);
    } catch(e) {
      Logger.log('⚠ Gagal tambah editor ' + email + ': ' + e.message);
    }

    // Admin juga bisa edit baris staf manapun (E:O)
    for (const adminEmail of CONFIG.ADMIN_EMAILS) {
      try { prot.addEditor(adminEmail); } catch(e) {}
    }
  }

  // Step 5: Proteksi P:Q — hanya admin
  const rangePQ = sheet.getRange(startRow, COL_NOTE, numRows, 2);
  const protPQ  = rangePQ.protect();
  protPQ.setDescription('Admin only — P:Q baris ' + startRow + '–' + endRow);
  protPQ.setWarningOnly(false);
  protPQ.removeEditors(protPQ.getEditors());
  protPQ.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try {
      protPQ.addEditor(adminEmail);
    } catch(err) {
      Logger.log('⚠ Gagal tambah admin P:Q ' + adminEmail + ': ' + err.message);
    }
  }

  Logger.log('✓ Proteksi P:Q (admin only) baris ' + startRow + '–' + endRow);
  Logger.log(divisi + ': proteksi selesai — ' + berhasil + ' orang');
}

// ── perbaruiEditorAdmin — Tambahkan admin baru ke semua proteksi ──────
// Jalankan sekali dari menu Admin setiap kali ada admin baru ditambahkan
// ke CONFIG.ADMIN_EMAILS agar langsung punya akses ke seluruh sheet.
function perbaruiEditorAdmin() {
  _requireAdmin();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const owner = Session.getEffectiveUser();

  let totalSheet = 0, totalProt = 0, totalGagal = 0;

  for (const divisi of CONFIG.DIVISI) {
    const sheets = ss.getSheets().filter(s =>
      s.getName().startsWith(divisi + '_') || s.getName() === divisi
    );

    for (const sheet of sheets) {
      totalSheet++;
      const prots = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

      for (const prot of prots) {
        // Pastikan owner selalu ada
        try { prot.addEditor(owner); } catch(e) {}

        for (const adminEmail of CONFIG.ADMIN_EMAILS) {
          try {
            prot.addEditor(adminEmail);
            totalProt++;
          } catch(e) {
            totalGagal++;
            Logger.log('⚠ Gagal: ' + adminEmail + ' — ' + e.message);
          }
        }
      }
    }
  }

  // Perbarui juga editor spreadsheet Settings eksternal
  try {
    if (CONFIG.SETTINGS_SPREADSHEET_ID) {
      const settingSS = SpreadsheetApp.openById(CONFIG.SETTINGS_SPREADSHEET_ID);
      for (const adminEmail of CONFIG.ADMIN_EMAILS) {
        try { settingSS.addEditor(adminEmail); } catch(e) {}
      }
    }
  } catch(e) {}

  // Re-apply proteksi Master_Data — ini adalah SHEET protection (bukan RANGE),
  // jadi tidak ter-cover loop di atas. Tanpa baris ini, admin baru di Settings
  // tidak akan bisa edit Master_Data sampai setupAwal dijalankan ulang.
  try { setupProteksiMaster(); } catch(e) { Logger.log('⚠ Master_Data: ' + e.message); }

  const msg = '✅ Selesai!\n\n' +
    'Sheet diproses : ' + totalSheet + '\n' +
    'Proteksi update: ' + totalProt + '\n' +
    'Master_Data    : proteksi di-refresh\n' +
    (totalGagal > 0 ? '⚠ Gagal        : ' + totalGagal + '\n' : '') +
    '\nSemua admin di CONFIG.ADMIN_EMAILS sudah ditambahkan.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ── bukaSettings — Buka spreadsheet Settings di tab baru ─────────────
function bukaSettings() {
  _requireAdmin();
  if (!CONFIG.SETTINGS_SPREADSHEET_ID) {
    try {
      SpreadsheetApp.getUi().alert(
        '⚠ Settings belum dikonfigurasi.\n\n' +
        'Minta Owner untuk menjalankan:\n' +
        '🔑 Owner → Inisialisasi Settings\n\n' +
        'Lalu copy Spreadsheet ID ke Config.js dan deploy ulang.'
      );
    } catch(e) {}
    return;
  }
  try {
    const url = SpreadsheetApp.openById(CONFIG.SETTINGS_SPREADSHEET_ID).getUrl();
    SpreadsheetApp.getUi().showModelessDialog(
      HtmlService.createHtmlOutput(
        '<p style="font-family:sans-serif;font-size:14px">Klik link di bawah untuk membuka Settings:</p>' +
        '<a href="' + url + '" target="_blank" style="font-size:14px">' + url + '</a>'
      ).setWidth(520).setHeight(80),
      '⚙️ Buka Settings'
    );
  } catch(e) {
    try { SpreadsheetApp.getUi().alert('❌ Gagal membuka Settings: ' + e.message); } catch(err) {}
  }
}

// ── setupSettings — Buat spreadsheet Settings terpusat ────────────────
// Dijalankan SEKALI oleh Owner dari menu 🔑 Owner → Inisialisasi Settings.
// Membuat Google Spreadsheet baru yang terpisah dari spreadsheet divisi.
// Satu file Settings digunakan bersama oleh semua divisi.
//
// Setelah selesai: copy Spreadsheet ID yang tampil → paste ke Config.js
// → jalankan deploy.sh agar semua divisi membaca dari file yang sama.
function setupSettings() {
  const owner = Session.getEffectiveUser();

  // Buat spreadsheet baru
  const newSS  = SpreadsheetApp.create(CONFIG.NAMA_INSTANSI + ' — Settings');
  const sheet  = newSS.getActiveSheet();
  sheet.setName('Settings');
  sheet.setTabColor('#0F6E56');
  sheet.setHiddenGridlines(true);

  // Baris 1: Header
  sheet.getRange(1, 1, 1, 3)
    .setValues([['KEY', 'VALUE', 'KETERANGAN']])
    .setBackground('#0F6E56').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);

  // Baris 2–5: Data setting operasional
  const rows = [
    [
      'ADMIN_EMAILS',
      CONFIG.ADMIN_EMAILS.join(', '),
      'Email admin/HRD — pisahkan dengan koma'
    ],
    [
      'JAM_REMINDER',
      CONFIG.JAM_REMINDER,
      'Jam reminder (format 24 jam, contoh: 17 = pukul 17:00)'
    ],
    [
      'SELISIH_MENIT_LOCK',
      CONFIG.SELISIH_MENIT_LOCK,
      'Menit setelah jam pulang diisi hingga baris dikunci (contoh: 30)'
    ],
    [
      'PLAN_JAM',
      CONFIG.PLAN_JAM.join(', '),
      'Pilihan shift — pisahkan dengan koma (contoh: 08:00 - 17:00, 09:00 - 18:00)'
    ],
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Format kolom KEY
  sheet.getRange(2, 1, rows.length, 1)
    .setBackground('#f0f2f5').setFontWeight('bold').setFontColor('#555');

  // Format kolom VALUE — ini yang diedit admin
  sheet.getRange(2, 2, rows.length, 1)
    .setBackground('#ffffff').setFontColor('#1a1a1a').setFontSize(11);

  // Format kolom KETERANGAN
  sheet.getRange(2, 3, rows.length, 1)
    .setBackground('#fafafa').setFontColor('#999').setFontSize(10)
    .setFontStyle('italic');

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 400);
  sheet.setRowHeights(2, rows.length, 32);
  sheet.setFrozenRows(1);

  // Proteksi kolom KEY dan KETERANGAN — hanya VALUE yang boleh diedit
  const protKey = sheet.getRange(1, 1, rows.length + 1, 1).protect();
  protKey.setDescription('KEY — jangan diubah');
  protKey.setWarningOnly(false);
  protKey.removeEditors(protKey.getEditors());
  protKey.addEditor(owner);

  const protKet = sheet.getRange(1, 3, rows.length + 1, 1).protect();
  protKet.setDescription('KETERANGAN — read only');
  protKet.setWarningOnly(false);
  protKet.removeEditors(protKet.getEditors());
  protKet.addEditor(owner);

  // Tambah semua admin sebagai editor spreadsheet Settings
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { newSS.addEditor(adminEmail); } catch(e) {}
  }

  const id  = newSS.getId();
  const url = newSS.getUrl();

  Logger.log('✓ Settings spreadsheet dibuat. ID: ' + id);
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Settings spreadsheet berhasil dibuat!\n\n' +
      '━━━ SPREADSHEET ID (copy ini) ━━━\n' +
      id + '\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      'Langkah selanjutnya:\n' +
      '1. Copy ID di atas\n' +
      '2. Buka Config.js → paste ke SETTINGS_SPREADSHEET_ID\n' +
      '3. Jalankan: bash deploy.sh\n\n' +
      'Setelah itu admin bisa edit settings langsung dari:\n' + url
    );
  } catch(e) {}
}
