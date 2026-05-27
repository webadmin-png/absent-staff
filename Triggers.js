// ═══════════════════════════════════════════════════════════════════════
// TRIGGERS.JS — Event trigger dan setup sistem
//
// Berisi:
//   onEditInstalled()      — guard proteksi edit per email (simple/installable trigger)
//   onOpen()      — buat menu "Absensi Saya" saat file dibuka
//   setupTrigger()— daftarkan semua trigger harian ke Apps Script
//   setupAwal()   — inisialisasi satu kali: proteksi + trigger + sheet + append
// ═══════════════════════════════════════════════════════════════════════

// ── onEdit — Guard proteksi edit ──────────────────────────────────────
// Mencegah staf mengedit kolom yang bukan haknya:
//   - Kolom A–D terkunci (tanggal, hari, nama, email)
//   - Kolom L (Jam Efektif) terkunci — formula otomatis
//   - Kolom P–Q (NOTE, SUNDAY) — hanya admin
//   - Kolom lain: hanya bisa edit baris milik sendiri (cocokkan email)
function _requireAdmin() {
  _loadSettings();
  let email = '';
  try { email = Session.getEffectiveUser().getEmail().trim().toLowerCase(); } catch(e) {}
  if (!email) {
    try { email = Session.getActiveUser().getEmail().trim().toLowerCase(); } catch(e) {}
  }
  if (!email) throw new Error(
    'Email tidak terdeteksi. Pastikan Anda login dengan akun Google.'
  );
  const isAdmin = CONFIG.ADMIN_EMAILS.map(a => a.toLowerCase()).includes(email);
  if (!isAdmin) throw new Error(
    '❌ Akses ditolak — fitur ini hanya untuk admin.\nEmail Anda: ' + email
  );
}




function onEditInstalled(e) {
  if (!e) return;
  _loadSettings();

  const sheet     = e.range.getSheet();
  const row       = e.range.getRow();
  const col       = e.range.getColumn();
  const sheetName = sheet.getName();

  // Hanya proses sheet divisi
  const isDivisiSheet = CONFIG.DIVISI.some(div =>
    sheetName === div || sheetName.startsWith(div + '_')
  );
  if (!isDivisiSheet) return;

  // Baris 1–3 adalah header — selalu tolak edit langsung
  if (row <= 3) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    return;
  }

  // Kolom A–D terkunci untuk semua staf
  if (col < COL_EDIT_START) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    try { SpreadsheetApp.getUi().alert('❌ Kolom ini terkunci.'); } catch(err) {}
    return;
  }

  // Kolom L (Jam Efektif) — formula otomatis, tidak boleh diubah
  if (col === COL_EFEKTIF) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    try { SpreadsheetApp.getUi().alert('❌ Kolom Jam Efektif dihitung otomatis.'); } catch(err) {}
    return;
  }

  // Kolom di luar batas edit (seharusnya tidak ada, tapi sebagai safety)
  if (col > COL_EDIT_END) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    return;
  }

  // Deteksi email user — getEffectiveUser() di simple trigger, fallback getActiveUser()
  let emailUser = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!emailUser) {
    emailUser = Session.getActiveUser().getEmail().trim().toLowerCase();
  }

  // Admin boleh edit segalanya
  const isAdmin = CONFIG.ADMIN_EMAILS.map(e => e.toLowerCase()).includes(emailUser);
  if (isAdmin) {
    _maybeLogJamEdit(e, emailUser, sheet, row, col);
    Logger.log('✓ Admin edit: ' + emailUser + ' baris ' + row);
    return;
  }

  // Kolom P (NOTE) dan Q (SUNDAY/RED DAY) — hanya admin
  if (col === COL_NOTE || col === COL_SUNDAY) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    try { SpreadsheetApp.getUi().alert('❌ Kolom ini hanya bisa diedit oleh admin.'); } catch(err) {}
    return;
  }

  // Jika email tidak terdeteksi (trigger tanpa izin), biarkan lewat
  if (!emailUser) {
    Logger.log('⚠ Email tidak bisa didapat — skip validasi baris ' + row);
    return;
  }

  // Staf hanya bisa edit baris milik mereka sendiri
  const emailBaris = String(sheet.getRange(row, COL_EMAIL).getValue())
    .trim().toLowerCase();
  if (!emailBaris) return;

  if (emailUser !== emailBaris) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    try {
      SpreadsheetApp.getUi().alert(
        '❌ Kamu hanya bisa edit baris milikmu sendiri.\n\n' +
        'Email kamu : ' + (emailUser || '(tidak terdeteksi)') + '\n' +
        'Email baris: ' + emailBaris
      );
    } catch(err) {}
    return;
  }

  _maybeLogJamEdit(e, emailUser, sheet, row, col);
  Logger.log('✓ Edit valid: ' + emailUser + ' baris ' + row);
}

// ── _maybeLogJamEdit — Catat ke audit log kalau yang diedit adalah kolom jam ──
// Dipanggil di akhir alur onEditInstalled untuk edit yang sudah lolos guard.
// Hanya log untuk kolom F-K (Masuk..Pulang). Edit kolom lain di-skip.
function _maybeLogJamEdit(e, emailUser, sheet, row, col) {
  const label = _kolomJamLabel(col);
  if (!label) return;  // bukan kolom jam — abaikan
  _logAuditJam({
    email     : emailUser,
    sumber    : 'sheet',
    targetSheet: sheet.getName(),
    row       : row,
    kolomLabel: label,
    nilaiLama : e.oldValue,
    nilaiBaru : e.value,
  });
}


// ── onOpen — Buat menu kustom saat spreadsheet dibuka ─────────────────
// CATATAN: onOpen adalah simple trigger — Google membatasi akses email user
// secara by-design untuk privasi. getEffectiveUser() dan getActiveUser()
// keduanya dapat return kosong untuk user non-owner.
// Solusi: tampilkan menu admin jika email tidak bisa dideteksi (kemungkinan admin)
// atau jika email memang terdaftar di ADMIN_EMAILS.
// Keamanan sesungguhnya dijaga di _requireAdmin() saat fungsi dijalankan.
function onOpen() {
  _loadSettings();
  const ui = SpreadsheetApp.getUi();

  // Coba baca email — di simple trigger sering return '' untuk non-owner
  const email = (
    Session.getEffectiveUser().getEmail() ||
    Session.getActiveUser().getEmail()
  ).trim().toLowerCase();

  const adminList = CONFIG.ADMIN_EMAILS.map(a => a.toLowerCase());

  // Tampilkan menu admin jika:
  //   - email kosong (tidak bisa dideteksi = simple trigger limitation), ATAU
  //   - email memang ada di ADMIN_EMAILS
  // Non-admin yang emailnya terdeteksi tidak akan melihat menu ini.
  const showAdmin = !email || adminList.includes(email);

  // Menu staf — semua pengguna
  // ui.createMenu('📋 Absensi Saya')
  //   .addItem('📍 Ke baris saya hari ini',    'keBarisHariIni')
  //   .addSeparator()
  //   .addItem('✅ Stamp MASUK',               'stampMasuk')
  //   .addItem('☕ Stamp ISTIRAHAT 1 MULAI',   'stampIst1Mulai')
  //   .addItem('▶ Stamp ISTIRAHAT 1 SELESAI',  'stampIst1Selesai')
  //   .addItem('☕ Stamp ISTIRAHAT 2 MULAI',   'stampIst2Mulai')
  //   .addItem('▶ Stamp ISTIRAHAT 2 SELESAI',  'stampIst2Selesai')
  //   .addItem('🏁 Stamp PULANG',              'stampPulang')
  //   .addSeparator()
  //   .addItem('📊 Rekap absensi saya',        'cekRekapSaya')
  //   .addToUi();

  if (showAdmin) {
    ui.createMenu('⚙️ Admin')
      .addItem('⚙️ Buka Settings',                    'bukaSettings')
      .addSeparator()
      .addItem('📅 Buat Sheet Bulan Baru',    'buatSheetBulanBaru')
      .addItem('➕ Append Hari Ini (manual)',          'appendHariIni')
      .addItem('🔄 Perbarui Rumus L–O (Backfill)',     'perbaruiRumusSemuaBaris')
      .addSeparator()
      .addItem('📋 Rekap Per Bulan (Auto-Formula)',           'generateTemplateRekap')
      .addItem('📅 Tarik Data Lintas Bulan',              'buatSheetRentang')
      .addItem('➕ Append Tanggal Baru',               'appendTanggalBaru')
      .addItem('⏰ Setup Trigger',               'setupTrigger')
      
      // .addItem('⚠️ Cek Belum Isi Pulang (manual)',    'cekBelumIsiPulang')
      // .addItem('🔒 Lock Baris Sudah Pulang (manual)', 'lockBarisWebSudahPulang')
      // .addSeparator()
      // .addItem('👤 Perbarui Akses Admin',             'perbaruiEditorAdmin')
      .addToUi();
  }

  // Menu owner — hanya untuk pemilik spreadsheet (setup sekali)
  try {
    const ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail().trim().toLowerCase();
    if (email && email === ownerEmail) {
      ui.createMenu('🔑 Owner')
        .addItem('🔧 Setup Trigger',          'setupTrigger')
        .addItem('🚀 Setup Awal',                'setupAwal')
        .addItem('⚙️ Inisialisasi Settings', 'setupSettings')
        .addToUi();
    }
  } catch(e) {}
}

// ── setupTrigger — Daftarkan semua trigger harian ─────────────────────
// Hapus semua trigger lama, lalu buat ulang:
//   06:00 → appendHariIni()     (append baris staf tiap pagi)
//   17:00 → cekBelumIsiPulang() (reminder staf yang belum isi pulang)
//   onEdit → guard proteksi (installable, lebih reliable dari simple trigger)
function setupTrigger() {
  _loadSettings();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Tanggal 1 tiap bulan jam 05:00 — buat sheet bulan baru sebelum append
  ScriptApp.newTrigger('buatSheetBulanBaru')
    .timeBased().onMonthDay(1).atHour(5).create();

  // Setiap hari jam 06:00 — append baris staf hari ini
  ScriptApp.newTrigger('appendHariIni')
    .timeBased().everyDays(1).atHour(6).create();

  // // Setiap hari jam JAM_REMINDER — reminder staf yang belum isi pulang
  // ScriptApp.newTrigger('cekBelumIsiPulang')
  //   .timeBased().everyDays(1).atHour(CONFIG.JAM_REMINDER).create();

  // Setiap jam — kunci baris 30 menit setelah jam pulang diisi
  ScriptApp.newTrigger('lockBarisWebSudahPulang')
    .timeBased().everyHours(1).create();

  // onEdit installable — guard proteksi per email.
  // newTrigger('onEditInstalled') = nama FUNGSI yang akan dipanggil saat fire.
  // .onEdit()                     = jenis EVENT (edit). Method ini wajib.
  // Pakai nama fungsi 'onEditInstalled' (bukan 'onEdit') supaya Apps Script
  // tidak ikut mem-fire-nya sebagai SIMPLE trigger dengan permission terbatas.
  ScriptApp.newTrigger('onEditInstalled')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit().create();

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Trigger aktif!\n\n' +
      '• Tgl 1 tiap bulan 05:00 — buat sheet bulan baru\n' +
      '• Setiap hari 06:00 — append baris hari ini otomatis\n' +
      '• Setiap jam — lock baris 30 menit setelah pulang\n' +
      '• onEdit — guard proteksi per email'
    );
  } catch(e) {}
}

// ── setupAwal — Inisialisasi satu kali saat pertama setup ─────────────
// Urutan: proteksi Master_Data → trigger → buat sheet bulan ini → append hari ini
// Jalankan sekali oleh HRD/admin saat pertama kali menggunakan sistem
function setupAwal() {
  _requireAdmin();
  try {
    setupProteksiMaster();
    setupTrigger();
    buatSheetBulanBaru();

    SpreadsheetApp.getUi().alert(
      '✅ Setup selesai!\n\n' +
      'Sistem sudah aktif:\n' +
      '• Sheet bulan ini sudah dibuat\n' +
      '• Baris hari ini sudah di-append\n' +
      '• Trigger harian aktif\n' +
      '• onEdit guard aktif\n\n' +
      'Langkah selanjutnya:\n' +
      '1. Share file ke staf (akses: Editor)\n' +
      '2. Minta staf test klik menu "📋 Absensi Saya"\n' +
      '3. Test stamp masuk/pulang'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    Logger.log('Error setupAwal: ' + e.message);
  }
}

function hapusSemuaTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}
