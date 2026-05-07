// ═══════════════════════════════════════════════════════════════════════
// CONFIG.JS — Konfigurasi global dan mapping kolom sheet
//
// STRUKTURAL (ubah hanya lewat kode + deploy):
//   SHEET_MASTER, DIVISI, NAMA_INSTANSI, TIMEZONE, DAYS_HOUR, semua COL_*
//
// OPERASIONAL (bisa diubah langsung dari sheet _Settings tanpa deploy):
//   ADMIN_EMAILS, JAM_REMINDER, SELISIH_MENIT_LOCK, PLAN_JAM
// ═══════════════════════════════════════════════════════════════════════

var CONFIG = {
  // ── Struktural — jangan ubah tanpa deploy ulang ───────────────────
  SHEET_MASTER : 'Master_Data',
  DIVISI       : ['WEB','HR'],

  // ID Google Spreadsheet khusus Settings (satu file untuk semua divisi).
  // Isi setelah Owner menjalankan 🔑 Owner → Inisialisasi Settings.
  SETTINGS_SPREADSHEET_ID : '1J936-7vEejkvIsxUjUGA7CJOIlrSB9tOoxs7qpSKGx0',
  NAMA_INSTANSI   : 'PT InFashion',
  TIMEZONE        : 'Asia/Makassar', // WITA (UTC+8)

  DAYS_HOUR : {
    REGULAR_DAYS : 7,   // Jam kerja normal per hari (senin–jumat)
    SATURDAY     : 5    // Jam kerja normal hari sabtu
  },

  // ── Operasional — nilai default (akan di-override dari sheet _Settings) ──
  ADMIN_EMAILS : [
    'webadmin@wooden-ships.com',
    'hrd@pt-infashion.com',
    'web@pt-infashion.com'
  ],

  JAM_REMINDER       : 17,
  SELISIH_MENIT_LOCK : 30,

  PLAN_JAM : [
    '08:00 - 17:00',
    '09:00 - 18:00',
    '07:00 - 16:00',
    '10:00 - 19:00',
    '08:00 - 13:00',
    '13:00 - 17:00',
  ]
};

// ── _loadSettings — Baca nilai operasional dari spreadsheet Settings ──
// Membuka spreadsheet terpisah via SETTINGS_SPREADSHEET_ID.
// Jika ID kosong atau gagal dibaca, CONFIG tetap pakai nilai default.
// Dipanggil di awal setiap entry point (trigger, menu, web app).
function _loadSettings() {
  try {
    if (!CONFIG.SETTINGS_SPREADSHEET_ID) return;
    const ss    = SpreadsheetApp.openById(CONFIG.SETTINGS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const [key, value] of data) {
      const k = String(key).trim();
      const v = String(value).trim();
      if (!k || !v) continue;

      switch (k) {
        case 'ADMIN_EMAILS': {
          const list = v.split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
          if (list.length > 0) CONFIG.ADMIN_EMAILS = list;
          break;
        }
        case 'JAM_REMINDER': {
          const jam = parseInt(v);
          if (!isNaN(jam) && jam >= 0 && jam <= 23) CONFIG.JAM_REMINDER = jam;
          break;
        }
        case 'SELISIH_MENIT_LOCK': {
          const menit = parseInt(v);
          if (!isNaN(menit) && menit > 0) CONFIG.SELISIH_MENIT_LOCK = menit;
          break;
        }
        case 'PLAN_JAM': {
          const plans = v.split(',').map(s => s.trim()).filter(Boolean);
          if (plans.length > 0) CONFIG.PLAN_JAM = plans;
          break;
        }
      }
    }
    Logger.log('✓ Settings dimuat dari spreadsheet eksternal');
  } catch(e) {
    Logger.log('⚠ _loadSettings gagal, pakai default Config.js: ' + e.message);
  }
}

// ── Mapping kolom sheet (1-indexed) ───────────────────────────────────
// A=1  Tanggal           → terkunci (diisi otomatis)
// B=2  Hari              → terkunci (diisi otomatis)
// C=3  Nama              → terkunci (diisi dari Master_Data)
// D=4  Email             → terkunci (diisi dari Master_Data)
// E=5  Status ▾          → editable staf (dropdown)
// F=6  Masuk             → editable staf (HH:mm)
// G=7  Ist. Pertama Mulai   → editable staf (HH:mm, opsional)
// H=8  Ist. Pertama Selesai → editable staf (HH:mm, opsional)
// I=9  Ist. Kedua Mulai     → editable staf (HH:mm, opsional)
// J=10 Ist. Kedua Selesai   → editable staf (HH:mm, opsional)
// K=11 Pulang            → editable staf (HH:mm)
// L=12 Jam Efektif       → formula otomatis (terkunci)
// M=13 Regular Hours     → formula otomatis (terkunci)
// N=14 OT 1              → formula otomatis (terkunci)
// O=15 OT 2              → formula otomatis (terkunci)
// P=16 NOTE              → editable admin only (dropdown)
// Q=17 SUNDAY/RED DAY    → editable admin only (dropdown)
// R=18 KETERANGAN        → editable staf (teks bebas, diisi saat tidak hadir)
// S=19 PLAN              → editable admin only (dropdown, dari CONFIG.PLAN_JAM)
// T=20 DEVICE STATUS        → editable staf (UPC/PC ON saat masuk, OFF saat pulang)
// U=21 CATATAN TELAT        → editable staf (alasan telat masuk)
// V=22 CATATAN PULANG AWAL  → editable staf (alasan pulang lebih awal)

const TOTAL_COL       = 22; // Jumlah kolom total (A sampai V)

const COL_TANGGAL     = 1;  // A
const COL_HARI        = 2;  // B
const COL_NAMA        = 3;  // C
const COL_EMAIL       = 4;  // D
const COL_STATUS      = 5;  // E — awal kolom editable staf
const COL_MASUK       = 6;  // F
const COL_IST1_M      = 7;  // G
const COL_IST1_S      = 8;  // H
const COL_IST2_M      = 9;  // I
const COL_IST2_S      = 10; // J
const COL_PULANG      = 11; // K
const COL_EFEKTIF     = 12; // L — formula, terkunci
const COL_REGULAR_JAM = 13; // M — formula, terkunci
const COL_OT1         = 14; // N — formula, terkunci
const COL_OT2         = 15; // O — formula, terkunci
const COL_NOTE        = 16; // P — admin only
const COL_SUNDAY      = 17; // Q — admin only
const COL_KETERANGAN  = 18; // R — editable staf (keterangan tidak hadir)
const COL_PLAN        = 19; // S — editable admin only (plan jam kerja)
const COL_DEVICE       = 20; // T — editable staf (UPC/PC status masuk/pulang)
const COL_TELAT        = 21; // U — editable staf (alasan telat masuk)
const COL_PULANG_AWAL  = 22; // V — editable staf (alasan pulang lebih awal)

// Batas kolom yang boleh diedit staf (E–V)
// Kolom L (COL_EFEKTIF) dikecualikan lewat guard di onEdit
// Kolom P–Q (NOTE, SUNDAY) dikecualikan lewat guard di onEdit
const COL_EDIT_START = COL_STATUS;      // E = 5
const COL_EDIT_END   = COL_PULANG_AWAL; // V = 22
