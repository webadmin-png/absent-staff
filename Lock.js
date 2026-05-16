// ═══════════════════════════════════════════════════════════════════════
// LOCK.JS — Penguncian baris dan pengingat jam pulang
//
// Berisi:
//   cekBelumIsiPulang()      — reminder sore untuk staf yang belum isi pulang
//   lockBarisWebSudahPulang() — kunci baris 30 menit setelah jam pulang diisi
// ═══════════════════════════════════════════════════════════════════════

// ── cekBelumIsiPulang — Reminder sore ────────────────────────────────
// Dipanggil via trigger setiap pukul JAM_REMINDER (default 17:00).
// Cari staf yang status = Hadir, sudah isi masuk, tapi belum isi pulang.
function cekBelumIsiPulang() {
  _requireAdmin();
  const today = getToday();
  const hasil = [];

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    for (let i = 3; i < data.length; i++) {
      const tgl    = data[i][0];
      const nama   = String(data[i][COL_NAMA   - 1]).trim();
      const status = String(data[i][COL_STATUS - 1]).trim();
      const masuk  = String(data[i][COL_MASUK  - 1]).trim();
      const pulang = String(data[i][COL_PULANG - 1]).trim();

      if (!isSameDate(tgl, today)) continue;
      if (status !== 'Hadir' || masuk === '' || pulang !== '') continue;
      hasil.push(divisi + ' — ' + nama);
    }
  }

  if (hasil.length === 0) {
    Logger.log('Semua sudah isi jam pulang.');
    return;
  }

  try {
    SpreadsheetApp.getUi().alert(
      '⚠ Belum isi jam pulang (' + hasil.length + ' orang):\n\n' +
      hasil.join('\n')
    );
  } catch(e) {
    Logger.log('Belum isi pulang: ' + hasil.join(', '));
  }
}

// ── lockBarisWebSudahPulang — Kunci baris setelah 30 menit pulang ─────
// Dipanggil via trigger setiap jam.
// Logic per baris:
//   - Status = Hadir DAN kolom K (Pulang) sudah terisi
//   - Selisih waktu sekarang vs jam pulang >= 30 menit
//   → Ganti proteksi menjadi hanya owner + admin (staf tidak bisa edit lagi)
function lockBarisWebSudahPulang() {
  _loadSettings();
  _requireAdmin();
  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) continue;

    const now   = new Date();
    const today = getToday();
    const data  = sheet.getDataRange().getValues();

    // Ambil semua proteksi sekali di luar loop (lebih efisien)
    const allProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const adminEmails    = CONFIG.ADMIN_EMAILS;

    for (let i = 3; i < data.length; i++) {
      const row    = i + 1;
      const tgl    = data[i][0];
      const nama   = String(data[i][COL_NAMA   - 1]).trim();
      const status = String(data[i][COL_STATUS - 1]).trim();
      const pulang = data[i][COL_PULANG - 1];
      const email  = String(data[i][COL_EMAIL  - 1]).trim();

      // Skip baris bukan hari ini
      if (!tgl || !isSameDate(tgl, today)) continue;

      // Skip jika bukan Hadir atau pulang belum diisi
      if (status.toLowerCase() !== 'hadir' || !pulang) continue;

      // Parse jam pulang ke objek Date hari ini
      let jamPulang;
      if (pulang instanceof Date) {
        jamPulang = new Date(
          now.getFullYear(), now.getMonth(), now.getDate(),
          pulang.getHours(), pulang.getMinutes(), 0
        );
      } else if (typeof pulang === 'string' && pulang.includes(':')) {
        const [h, m] = pulang.split(':');
        jamPulang = new Date(
          now.getFullYear(), now.getMonth(), now.getDate(),
          parseInt(h), parseInt(m), 0
        );
      } else {
        Logger.log('❌ Format pulang tidak valid: ' + pulang + ' baris ' + row);
        continue;
      }

      // Koreksi jika jam pulang sudah lewat tengah malam
      if (jamPulang > now) {
        Logger.log('⏳ ' + nama + ' (baris ' + row + '): pulang masih di masa depan (' +
          Utilities.formatDate(jamPulang, CONFIG.TIMEZONE, 'HH:mm') + '), skip lock');
        continue;
      }

      const selisihMenit = (now - jamPulang) / 60000;
      if (selisihMenit < CONFIG.SELISIH_MENIT_LOCK) {
        Logger.log('⏳ ' + nama + ' (baris ' + row + ') belum ' +
          CONFIG.SELISIH_MENIT_LOCK + ' menit (' +
          Math.round(selisihMenit) + ' menit)');
        continue;
      }

      
      // Hapus proteksi lama pada baris ini
      const existingProt = allProtections.find(prot => {
        const r = prot.getRange();
        return row >= r.getRow() && row <= r.getLastRow();
      });
      if (existingProt) {
        Logger.log('🗑 Hapus proteksi lama baris ' + row);
        existingProt.remove();
      }

      // Buat proteksi baru — seluruh baris, hanya admin
      const rowRange = sheet.getRange(row + ':' + row);
      const newProt  = rowRange.protect();
      newProt.setDescription(nama + ' — terkunci ' + now.toTimeString().slice(0,5));
      newProt.removeEditors(newProt.getEditors());

      // Tambah owner + semua admin email
      const owner = Session.getEffectiveUser();
      newProt.addEditor(owner);
      for (const adminEmail of adminEmails) {
        try { newProt.addEditor(adminEmail); } catch(e) {}
      }

      Logger.log('🔒 Terkunci: ' + nama + ' baris ' + row);
    }
  }
}
