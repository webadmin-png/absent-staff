// ═══════════════════════════════════════════════════════════════════════
// REKAP.JS — Perhitungan rekap jam kerja dan laporan gaji
//
// Berisi:
//   buatSheetRekap()         — buat sheet ringkasan jam dari semua divisi
//   generateFullMonth()      — generate sheet absensi satu bulan penuh
//   hitungRekap()            — hitung rekap dari sheet sumber via prompt
//   generateTemplateRekap()  — buat template rekap per divisi dengan rumus Excel
//
//   Helper kalkulasi (dipanggil dari buatSheetRekap & hitungRekap):
//   calculateRegularHours, calculateOt1Hrs, calculateOtAfter1stHrs,
//   calculateSundayRedDayHrs, calculateSundayRedDayOt1stHrs,
//   calculateSundayRedDayOtAfter1stHrs, calculateCapacityHrs,
//   calculateDayRegMealAllowance, calculateSundayRedDayMealAllowance,
//   calculateOtMeal, calculateDeductionSkillBonus
// ═══════════════════════════════════════════════════════════════════════

// ── buatSheetRekap — Buat sheet ringkasan jam semua divisi ─────────────
// Baca semua sheet divisi aktif, akumulasi jam per karyawan,
// tulis ke sheet "Rekap_Absensi"
function buatSheetRekap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rekapSheetName = 'Rekap_Absensi';
  let rekapSheet = ss.getSheetByName(rekapSheetName);

  if (rekapSheet) ss.deleteSheet(rekapSheet);
  rekapSheet = ss.insertSheet(rekapSheetName);
  rekapSheet.setTabColor('#FF5722');
  rekapSheet.setHiddenGridlines(true);

  const columnMapping = {
    'Name': 1,
    'Division': 2,
    'TTL Regular HRS': 3,
    'TTL O/T 1st HRS': 4,
    'TTL O/T After 1st HRS': 5,
    'Bonus TTL O/T After 1st HRS': 6,
    'Sunday/Red Day TTL HRS': 7,
    'Sunday/Red Day TTL O/T After 1st HRS': 8,
    'TTL Working HRS for Capacity': 9,
    'TTL Day Reg Meal Allowance': 10,
    'TTL OT Meal': 11,
    'Deduction Skill Bonus Based on Day Come In': 12,
  };

  const headers = Object.keys(columnMapping);
  rekapSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#178232').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const columnWidths = [150,150,150,150,200,250,200,250,200,200,150,300];
  columnWidths.forEach((w, i) => rekapSheet.setColumnWidth(i + 1, w));
  rekapSheet.setFrozenRows(1);

  // Isi nama + divisi dari Master_Data
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const masterData  = masterSheet.getRange('A4:D200').getValues();
  const employeeDivisionMap = {};
  masterData.forEach(row => {
    const division = String(row[0]).trim();
    const name     = String(row[1]).trim();
    if (division && name) employeeDivisionMap[name] = division;
  });

  const divisionData = Object.entries(employeeDivisionMap)
    .map(([name, division]) => [name, division]);
  rekapSheet.getRange(2, 1, divisionData.length, 2).setValues(divisionData);

  const employeeHoursMap = {};

  CONFIG.DIVISI.forEach(divisi => {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 3; i < data.length; i++) {
      const name             = String(data[i][COL_NAMA        - 1]).trim();
      if (!name) continue;

      const regularHrs       = parseTimeFraction(data[i][COL_REGULAR_JAM - 1]) || 0;
      const ot1Hrs           = parseTimeFraction(data[i][COL_OT1         - 1]) || 0;
      const otAfter1stHrs    = parseTimeFraction(data[i][COL_OT2         - 1]) || 0;
      const sundayRedDayNote = String(data[i][COL_SUNDAY      - 1]).trim();
      const note             = String(data[i][COL_NOTE        - 1]).trim();
      const namaHari         = String(data[i][COL_HARI        - 1]).trim();
      const effectiveHrs     = parseTimeFraction(data[i][COL_EFEKTIF - 1]) || 0;

      if (!employeeHoursMap[name]) {
        employeeHoursMap[name] = {
          ttlRegularHrs: 0, ttlOt1Hrs: 0, ttlOtAfter1stHrs: 0,
          ttlBonusOtAfter1stHrs: 0, ttlSundayRedDayHrs: 0,
          ttlSundayRedDayOt1stHrs: 0, ttlSundayRedDayOtAfter1stHrs: 0,
          ttlWorkingHrsForCapacity: 0, ttlDayRegMealAllowance: 0,
          ttlSundayRedDayMealAllowance: 0, ttlOtMeal: 0,
          deductionSkillBonusBasedOnDayComeIn: 0,
        };
      }

      const emp = employeeHoursMap[name];
      emp.ttlRegularHrs += calculateRegularHours(regularHrs, namaHari, note, sundayRedDayNote);
      emp.ttlOt1Hrs     += calculateOt1Hrs(ot1Hrs, namaHari, note, sundayRedDayNote);

      const { totalOtAfter1stHrs, bonusOtAfter1stHrs } =
        calculateOtAfter1stHrs(ot1Hrs, otAfter1stHrs, namaHari, note, sundayRedDayNote);
      emp.ttlOtAfter1stHrs      += totalOtAfter1stHrs;
      emp.ttlBonusOtAfter1stHrs += bonusOtAfter1stHrs;

      emp.ttlSundayRedDayHrs           += calculateSundayRedDayHrs(regularHrs, namaHari, sundayRedDayNote);
      emp.ttlSundayRedDayOt1stHrs      += calculateSundayRedDayOt1stHrs(ot1Hrs, namaHari, sundayRedDayNote);
      emp.ttlSundayRedDayOtAfter1stHrs += calculateSundayRedDayOtAfter1stHrs(otAfter1stHrs, namaHari, sundayRedDayNote);
      emp.ttlWorkingHrsForCapacity     += calculateCapacityHrs(regularHrs, namaHari, note);
      emp.ttlDayRegMealAllowance       += calculateDayRegMealAllowance(namaHari, sundayRedDayNote, note);
      emp.ttlSundayRedDayMealAllowance += calculateSundayRedDayMealAllowance(namaHari, sundayRedDayNote);
      emp.ttlOtMeal                    += calculateOtMeal(effectiveHrs);
      emp.deductionSkillBonusBasedOnDayComeIn +=
        calculateDeductionSkillBonus(note, sundayRedDayNote);
    }
  });

  const rekapData = rekapSheet.getDataRange().getValues();
  for (const [name, hours] of Object.entries(employeeHoursMap)) {
    let targetRow = -1;
    for (let r = 1; r < rekapData.length; r++) {
      if (String(rekapData[r][0]).trim() === name) { targetRow = r + 1; break; }
    }
    if (targetRow === -1) continue;

    rekapSheet.getRange(targetRow, 3, 1, 10).setValues([[
      decimalToHHMM(hours.ttlRegularHrs),
      decimalToHHMM(hours.ttlOt1Hrs),
      decimalToHHMM(hours.ttlOtAfter1stHrs),
      decimalToHHMM(hours.ttlBonusOtAfter1stHrs),
      decimalToHHMM(hours.ttlSundayRedDayHrs),
      decimalToHHMM(hours.ttlSundayRedDayOtAfter1stHrs),
      decimalToHHMM(hours.ttlWorkingHrsForCapacity),
      hours.ttlDayRegMealAllowance,
      hours.ttlOtMeal,
      hours.deductionSkillBonusBasedOnDayComeIn,
    ]]);
  }

  SpreadsheetApp.getUi().alert('✅ Sheet rekap berhasil dibuat!');
}

// ── generateFullMonth — Generate sheet satu bulan penuh ───────────────
// HRD input bulan + divisi via prompt, lalu buat sheet dengan semua hari
// dan semua staf untuk bulan tersebut
function generateFullMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Step 1: Input bulan
  const inputBulan = ui.prompt(
    '📅 Generate Full Month (1/3)',
    'Masukkan bulan dan tahun\nFormat: MM/YYYY\nContoh: 04/2026',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputBulan.getSelectedButton() !== ui.Button.OK) return;

  const bulanStr = inputBulan.getResponseText().trim();
  if (!bulanStr.match(/^\d{2}\/\d{4}$/)) {
    ui.alert('❌ Format salah. Gunakan MM/YYYY\nContoh: 04/2026'); return;
  }

  const [bulanNum, tahunNum] = bulanStr.split('/').map(Number);
  if (bulanNum < 1 || bulanNum > 12) {
    ui.alert('❌ Bulan harus antara 01 sampai 12.'); return;
  }

  const tglRef     = new Date(tahunNum, bulanNum - 1, 1);
  const namaBulan  = Utilities.formatDate(tglRef, CONFIG.TIMEZONE, 'MMM_yyyy');
  const labelBulan = Utilities.formatDate(tglRef, CONFIG.TIMEZONE, 'MMMM yyyy').toUpperCase();

  // Step 2: Input divisi
  const daftarDivisi = CONFIG.DIVISI.join(', ');
  const inputDivisi  = ui.prompt(
    '📅 Generate Full Month (2/3)',
    'Pilih divisi:\nKetik nama divisi (contoh: WEB)\n' +
    'Atau ketik "SEMUA" untuk semua divisi\n\nDivisi tersedia: ' + daftarDivisi,
    ui.ButtonSet.OK_CANCEL
  );
  if (inputDivisi.getSelectedButton() !== ui.Button.OK) return;

  const inputDivisiStr = inputDivisi.getResponseText().trim().toUpperCase();
  let targetDivisi     = [];

  if (inputDivisiStr === 'SEMUA') {
    targetDivisi = CONFIG.DIVISI;
  } else {
    const divisiValid = CONFIG.DIVISI.find(d => d.toUpperCase() === inputDivisiStr);
    if (!divisiValid) {
      ui.alert('❌ Divisi "' + inputDivisiStr + '" tidak ditemukan.\n\nDivisi tersedia: ' + daftarDivisi);
      return;
    }
    targetDivisi = [divisiValid];
  }

  // Step 3: Konfirmasi
  const konfirmasi = ui.alert(
    '⚠ Konfirmasi Generate Full Month',
    'Bulan  : ' + labelBulan + '\n' +
    'Divisi : ' + targetDivisi.join(', ') + '\n' +
    'Sheet  : ' + targetDivisi.map(d => d + '_' + namaBulan).join(', ') + '\n\n' +
    '⚠ Sheet yang sudah ada akan di-OVERWRITE!\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi !== ui.Button.YES) return;

  // Step 4: Hitung semua hari dalam bulan
  const hariDlmBln = new Date(tahunNum, bulanNum, 0).getDate();
  const hariSemua  = [];
  for (let d = 1; d <= hariDlmBln; d++) {
    hariSemua.push(new Date(tahunNum, bulanNum - 1, d, 12, 0, 0));
  }

  // Step 5: Ambil staf dari Master_Data
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) { ui.alert('❌ Sheet Master_Data tidak ditemukan.'); return; }

  const masterData = master.getRange('A4:D200').getValues()
    .filter(r => r[0] !== '' && r[1] !== '' && String(r[3]).trim().toUpperCase() === 'TRUE');

  const stafPerDivisi = {};
  for (const k of masterData) {
    const div = String(k[0]).trim().toUpperCase();
    if (!stafPerDivisi[div]) stafPerDivisi[div] = [];
    stafPerDivisi[div].push({ nama: String(k[1]).trim(), email: String(k[2]).trim() });
  }

  // Step 6: Generate sheet per divisi
  const hasil = [];
  const today = getToday();

  for (const divisi of targetDivisi) {
    const namaSheet = divisi + '_' + namaBulan;
    const staf      = stafPerDivisi[divisi.toUpperCase()] || [];

    if (staf.length === 0) {
      hasil.push('⚠ ' + divisi + ': tidak ada staf aktif di Master_Data'); continue;
    }

    const sheetLama = ss.getSheetByName(namaSheet);
    if (sheetLama) ss.deleteSheet(sheetLama);
    const sheet = ss.insertSheet(namaSheet);
    sheet.setTabColor('#1D9E75');
    sheet.setHiddenGridlines(true);

    // Baris 1: Judul
    sheet.getRange(1, 1, 1, TOTAL_COL).merge()
      .setValue('ABSENSI ' + divisi + ' — ' + labelBulan)
      .setBackground('#0F6E56').setFontColor('#FFFFFF')
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);

    // Baris 2: Legenda
    const legends = [
      [1,  4, 'ABU = sudah lewat',   '#F1EFE8', '#5F5E5A'],
      [5,  4, 'PUTIH = bisa diedit', '#FFFFFF',  '#2C2C2A'],
      [9,  4, 'UNGU = formula auto', '#EEEDFE',  '#534AB7'],
      [13, 10, 'KUNING = hari ini',   '#FFF9C4',  '#633806'],
    ];
    for (const [startCol, span, text, bg, fg] of legends) {
      sheet.getRange(2, startCol, 1, span).merge()
        .setValue(text).setBackground(bg).setFontColor(fg)
        .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center')
        .setBorder(true,true,true,true,false,false,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(2, 16);

    // Baris 3: Header
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
        .setBorder(true,true,true,true,false,false,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(3, 44);
    const colWidths = [90,80,130,180,70,70,90,90,90,90,70,100,60,60,60,120,140,180,120,200,180,180];
    colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(3);

    // Siapkan semua baris data
    const allRows = [];
    for (const hari of hariSemua) {
      const hariNorm = new Date(hari.getFullYear(), hari.getMonth(), hari.getDate(), 12, 0, 0);
      const namaHari = Utilities.formatDate(hariNorm, CONFIG.TIMEZONE, 'EEEE');
      for (const s of staf) {
        allRows.push([hariNorm, namaHari, s.nama, s.email,
          '','','','','','','',  // E–K: staf
          '','','','',           // L–O: formula
          '','',                 // P–Q: admin
          '', '', '',            // R: Keterangan, S: Plan, T: Device
          '', '']);              // U: Catatan Telat, V: Catatan Pulang Awal
      }
    }

    const startRow = 4;
    sheet.getRange(startRow, 1, allRows.length, TOTAL_COL).setValues(allRows);
    sheet.getRange(startRow, 1, allRows.length, 1).setNumberFormat('DD/MM/YYYY');

    // Warna per baris
    for (let idx = 0; idx < allRows.length; idx++) {
      const r       = idx + startRow;
      const tglBaris = allRows[idx][0];
      const tglNorm  = new Date(tglBaris.getFullYear(), tglBaris.getMonth(), tglBaris.getDate());
      const todayNorm = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const isSunday  = tglNorm.getDay() === 0;
      const isToday   = tglNorm.getTime() === todayNorm.getTime();
      const isPastDay = tglNorm.getTime() < todayNorm.getTime();

      // Hari Minggu → seluruh baris merah (konsisten dengan highlightHariIni).
      // Dicek paling awal supaya menimpa warna kuning(hari ini)/abu(lewat)/putih.
      if (isSunday) {
        sheet.getRange(r, 1, 1, TOTAL_COL)
          .setBackground('#FFCDD2').setFontColor('#B71C1C');
        continue;
      }

      let bgLocked, bgEdit, bgFormula;
      if (isToday) {
        bgLocked = bgEdit = bgFormula = '#FFF9C4';
      } else if (isPastDay) {
        bgLocked = '#F1EFE8'; bgEdit = '#F8F8F8'; bgFormula = '#EEEDFE';
      } else {
        bgLocked = '#F1EFE8'; bgEdit = '#FFFFFF'; bgFormula = '#EEEDFE';
      }
      sheet.getRange(r, 1,  1, 4).setBackground(bgLocked).setFontColor('#5F5E5A');
      sheet.getRange(r, 5,  1, 7).setBackground(bgEdit).setFontColor('#2C2C2A');
      sheet.getRange(r, 12, 1, 1).setBackground(bgFormula).setFontColor('#534AB7').setFontWeight('bold');
      sheet.getRange(r, 13, 1, 3).setBackground(bgFormula).setFontColor('#534AB7').setFontWeight('bold');
      sheet.getRange(r, 16, 1, 3).setBackground(bgEdit).setFontColor('#2C2C2A');             // P:R
      sheet.getRange(r, COL_PLAN,   1, 1).setBackground(bgEdit).setFontColor('#085041').setFontWeight('bold'); // S
      sheet.getRange(r, COL_DEVICE, 1, 1).setBackground(bgEdit).setFontColor('#2C2C2A');    // T
      sheet.getRange(r, COL_TELAT,  1, 2).setBackground(bgEdit).setFontColor('#E65100');    // U:V
    }

    sheet.getRange(startRow, 1, allRows.length, TOTAL_COL)
      .setBorder(true,true,true,true,true,true,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);

    // Formula L–O
    _pasangFormulaBaris(sheet, startRow, allRows.length);

    sheet.getRange(startRow, COL_EFEKTIF,     allRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(startRow, COL_REGULAR_JAM, allRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(startRow, COL_OT1,         allRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(startRow, COL_OT2,         allRows.length, 1).setNumberFormat('[h]:mm');

    setupValidasiBaris(sheet, startRow, allRows.length);

    const headerProt = sheet.getRange(1, 1, 3, TOTAL_COL).protect();
    headerProt.setDescription('Header — hanya owner yang bisa edit');
    headerProt.setWarningOnly(false);
    headerProt.removeEditors(headerProt.getEditors());
    headerProt.addEditor(Session.getEffectiveUser());

    proteksiBarisBaru(sheet, divisi, startRow, allRows.length);
    groupByToday(sheet);

    hasil.push('✓ ' + namaSheet + ' — ' + staf.length + ' staf × ' +
      hariSemua.length + ' hari = ' + allRows.length + ' baris');
    Logger.log('generateFullMonth selesai: ' + namaSheet);
  }

  highlightHariIni();
  ui.alert(
    '✅ Generate Full Month selesai!\n\n' +
    'Bulan  : ' + labelBulan + '\n' +
    'Divisi : ' + targetDivisi.join(', ') + '\n\n' +
    hasil.join('\n')
  );
}

// ── hitungRekap — Hitung rekap dari sheet sumber via prompt ───────────
// HR input nama sheet sumber + nama sheet hasil → script hitung semua kolom
// Bisa dipanggil dari menu (UI) atau dari editor/trigger (tanpa UI).
// Jika dipanggil tanpa UI, gunakan hitungRekapAuto(namaSumber, namaHasil).
function hitungRekap(namaSumber, namaHasil) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Deteksi apakah bisa pakai UI
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (_) {}

  // ── Ambil input: dari parameter atau dari UI prompt ──
  if (!namaSumber || !namaHasil) {
    if (!ui) throw new Error(
      'hitungRekap dipanggil tanpa UI dan tanpa parameter.\n' +
      'Gunakan: hitungRekapAuto("NamaSheetSumber", "NamaSheetHasil")'
    );

    var semuaSheet  = ss.getSheets().map(function(s) { return s.getName(); }).join('\n');
    var inputSumber = ui.prompt(
      '📊 Hitung Rekap (1/2)',
      'Masukkan nama sheet sumber:\n\nSheet tersedia:\n' + semuaSheet,
      ui.ButtonSet.OK_CANCEL
    );
    if (inputSumber.getSelectedButton() !== ui.Button.OK) return;

    namaSumber = inputSumber.getResponseText().trim();

    var inputHasil = ui.prompt(
      '📊 Hitung Rekap (2/2)',
      'Masukkan nama sheet rekap hasil\n(contoh: Rekap_GajiMar2026)\n\n' +
      '⚠ Sheet lama akan di-overwrite.',
      ui.ButtonSet.OK_CANCEL
    );
    if (inputHasil.getSelectedButton() !== ui.Button.OK) return;

    namaHasil = inputHasil.getResponseText().trim();
    if (!namaHasil) { ui.alert('❌ Nama sheet hasil tidak boleh kosong.'); return; }
  }

  const sheetSumber = ss.getSheetByName(namaSumber);
  if (!sheetSumber) {
    var msg = '❌ Sheet "' + namaSumber + '" tidak ditemukan.';
    if (ui) { ui.alert(msg); return; }
    throw new Error(msg);
  }

  const data       = sheetSumber.getDataRange().getValues();
  const totalBaris = data.length - 3;
  if (totalBaris <= 0) {
    var msg2 = '❌ Sheet "' + namaSumber + '" tidak ada datanya.';
    if (ui) { ui.alert(msg2); return; }
    throw new Error(msg2);
  }

  if (ui) {
    var konfirmasi = ui.alert(
      '⚠ Konfirmasi Hitung Rekap',
      'Sheet sumber : ' + namaSumber + '\nSheet hasil  : ' + namaHasil +
      '\nTotal baris  : ' + totalBaris + '\n\nLanjutkan?',
      ui.ButtonSet.YES_NO
    );
    if (konfirmasi !== ui.Button.YES) return;
  }

  let rekapSheet = ss.getSheetByName(namaHasil);
  if (rekapSheet) ss.deleteSheet(rekapSheet);
  rekapSheet = ss.insertSheet(namaHasil);
  rekapSheet.setTabColor('#FF5722');
  rekapSheet.setHiddenGridlines(true);

  const headers = [
    'Name','Division','TTL Regular HRS','TTL O/T 1st HRS',
    'TTL O/T After 1st HRS','Bonus TTL O/T After 1st HRS',
    'Sunday/Red Day TTL HRS','Sunday/Red Day TTL O/T 1st HRS',
    'Sunday/Red Day TTL O/T After 1st HRS','TTL Working HRS for Capacity',
    'TTL Day Reg Meal Allowance','TTL OT Meal',
    'Deduction Skill Bonus Based on Day Come In',
  ];
  rekapSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]).setBackground('#178232').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const columnWidths = [150,150,150,150,200,250,200,250,250,200,200,150,300];
  columnWidths.forEach((w, i) => rekapSheet.setColumnWidth(i + 1, w));
  rekapSheet.setFrozenRows(1);

  const masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const masterData  = masterSheet.getRange('A4:D200').getValues();
  const divisiMap   = {};
  masterData.forEach(row => {
    const div  = String(row[0]).trim();
    const nama = String(row[1]).trim();
    if (div && nama) divisiMap[nama] = div;
  });

  const employeeHoursMap = {};
  for (let i = 3; i < data.length; i++) {
    const name = String(data[i][COL_NAMA - 1]).trim();
    if (!name) continue;

    const namaHari         = String(data[i][COL_HARI   - 1]).trim();
    const note             = String(data[i][COL_NOTE   - 1]).trim();
    const sundayRedDayNote = String(data[i][COL_SUNDAY - 1]).trim();
    const effectiveHrs     = parseTimeFraction(data[i][COL_EFEKTIF     - 1]);
    const regularHrs       = parseTimeFraction(data[i][COL_REGULAR_JAM - 1]);
    const ot1Hrs           = parseTimeFraction(data[i][COL_OT1         - 1]);
    const otAfter1stHrs    = parseTimeFraction(data[i][COL_OT2         - 1]);

    if (!employeeHoursMap[name]) {
      employeeHoursMap[name] = {
        divisi: divisiMap[name] || '—',
        ttlRegularHrs: 0, ttlOt1Hrs: 0, ttlOtAfter1stHrs: 0,
        ttlBonusOtAfter1stHrs: 0, ttlSundayRedDayHrs: 0,
        ttlSundayRedDayOt1stHrs: 0, ttlSundayRedDayOtAfter1stHrs: 0,
        ttlWorkingHrsForCapacity: 0, ttlDayRegMealAllowance: 0,
        ttlSundayRedDayMealAllowance: 0, ttlOtMeal: 0,
        deductionSkillBonusBasedOnDayComeIn: 0,
      };
    }

    const emp = employeeHoursMap[name];
    emp.ttlRegularHrs += calculateRegularHours(regularHrs, namaHari, note, sundayRedDayNote);
    emp.ttlOt1Hrs     += calculateOt1Hrs(ot1Hrs, namaHari, note, sundayRedDayNote);

    const { totalOtAfter1stHrs, bonusOtAfter1stHrs } =
      calculateOtAfter1stHrs(ot1Hrs, otAfter1stHrs, namaHari, note, sundayRedDayNote);
    emp.ttlOtAfter1stHrs      += totalOtAfter1stHrs;
    emp.ttlBonusOtAfter1stHrs += bonusOtAfter1stHrs;

    emp.ttlSundayRedDayHrs           += calculateSundayRedDayHrs(regularHrs, namaHari, sundayRedDayNote);
    emp.ttlSundayRedDayOt1stHrs      += calculateSundayRedDayOt1stHrs(ot1Hrs, namaHari, sundayRedDayNote);
    emp.ttlSundayRedDayOtAfter1stHrs += calculateSundayRedDayOtAfter1stHrs(otAfter1stHrs, namaHari, sundayRedDayNote);
    emp.ttlWorkingHrsForCapacity     += calculateCapacityHrs(regularHrs, namaHari, note);
    emp.ttlDayRegMealAllowance       += calculateDayRegMealAllowance(namaHari, sundayRedDayNote, note);
    emp.ttlSundayRedDayMealAllowance += calculateSundayRedDayMealAllowance(namaHari, sundayRedDayNote);
    emp.ttlOtMeal                    += calculateOtMeal(effectiveHrs);
    emp.deductionSkillBonusBasedOnDayComeIn +=
      calculateDeductionSkillBonus(note, sundayRedDayNote);
  }

  const outputRows = [];
  for (const [name, emp] of Object.entries(employeeHoursMap)) {
    outputRows.push([
      name, emp.divisi,
      decimalToHHMM(emp.ttlRegularHrs), decimalToHHMM(emp.ttlOt1Hrs),
      decimalToHHMM(emp.ttlOtAfter1stHrs), decimalToHHMM(emp.ttlBonusOtAfter1stHrs),
      decimalToHHMM(emp.ttlSundayRedDayHrs), decimalToHHMM(emp.ttlSundayRedDayOt1stHrs),
      decimalToHHMM(emp.ttlSundayRedDayOtAfter1stHrs),
      decimalToHHMM(emp.ttlWorkingHrsForCapacity),
      emp.ttlDayRegMealAllowance, emp.ttlOtMeal,
      emp.deductionSkillBonusBasedOnDayComeIn,
    ]);
  }

  if (outputRows.length === 0) {
    var msg3 = '⚠ Tidak ada data karyawan ditemukan.';
    if (ui) { ui.alert(msg3); return; }
    throw new Error(msg3);
  }

  rekapSheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  for (let r = 2; r <= 1 + outputRows.length; r++) {
    rekapSheet.getRange(r, 1, 1, headers.length)
      .setBackground(r % 2 === 0 ? '#F7FDFB' : '#FFFFFF');
  }
  rekapSheet.getRange(2, 1, outputRows.length, 2).setFontWeight('bold');
  rekapSheet.getRange(1, 1, 1 + outputRows.length, headers.length)
    .setBorder(true,true,true,true,true,true,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);

  var msgDone = '✅ Rekap selesai!\n\n' +
    'Sheet sumber : ' + namaSumber + '\nSheet hasil  : ' + namaHasil +
    '\nTotal baris  : ' + totalBaris + '\nTotal staf   : ' + outputRows.length;
  if (ui) { ui.alert(msgDone); }
  Logger.log(msgDone);
}

// ── generateTemplateRekap — Template rekap dengan rumus Excel ─────────
// Berbeda dari hitungRekap: ini menggunakan SUMIFS/COUNTIFS langsung di sheet
// sehingga rekap otomatis update jika data sumber berubah
function generateTemplateRekap() {
  _requireAdmin();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {
    Logger.log('generateTemplateRekap: tidak dapat dijalankan dari konteks trigger/non-UI.');
    return;
  }

  const semuaSheet  = ss.getSheets().map(s => s.getName()).join('\n');
  const inputSumber = ui.prompt(
    '📊 Generate Template Rekap (1/3)',
    'Masukkan nama sheet sumber:\n\nSheet tersedia:\n' + semuaSheet,
    ui.ButtonSet.OK_CANCEL
  );
  if (inputSumber.getSelectedButton() !== ui.Button.OK) return;

  const namaSumber  = inputSumber.getResponseText().trim();
  const sheetSumber = ss.getSheetByName(namaSumber);
  if (!sheetSumber) { ui.alert('❌ Sheet "' + namaSumber + '" tidak ditemukan.'); return; }

  const inputPeriode = ui.prompt(
    '📊 Generate Template Rekap (2/3)',
    'Masukkan nama periode rekap (contoh: GajiMar2026)\n\n' +
    'Sheet rekap akan dibuat dengan nama:\n' +
    CONFIG.DIVISI.map(d => 'Rekap_' + d + '_[periode]').join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (inputPeriode.getSelectedButton() !== ui.Button.OK) return;

  const periode = inputPeriode.getResponseText().trim();
  if (!periode) { ui.alert('❌ Nama periode tidak boleh kosong.'); return; }

  const inputDivisi = ui.prompt(
    '📊 Generate Template Rekap (3/3)',
    'Pilih divisi (contoh: WEB) atau ketik "SEMUA"\n\nDivisi tersedia: ' + CONFIG.DIVISI.join(', '),
    ui.ButtonSet.OK_CANCEL
  );
  if (inputDivisi.getSelectedButton() !== ui.Button.OK) return;

  const inputDivisiStr = inputDivisi.getResponseText().trim().toUpperCase();
  let targetDivisi     = [];

  if (inputDivisiStr === 'SEMUA') {
    targetDivisi = CONFIG.DIVISI;
  } else {
    const divisiValid = CONFIG.DIVISI.find(d => d.toUpperCase() === inputDivisiStr);
    if (!divisiValid) {
      ui.alert('❌ Divisi "' + inputDivisiStr + '" tidak ditemukan.\n\nDivisi tersedia: ' + CONFIG.DIVISI.join(', '));
      return;
    }
    targetDivisi = [divisiValid];
  }

  const namaSheetHasil = targetDivisi.map(d => 'Rekap_' + d + '_' + periode);
  const konfirmasi = ui.alert(
    '⚠ Konfirmasi Generate Template Rekap',
    'Sheet sumber : ' + namaSumber + '\nPeriode      : ' + periode +
    '\nDivisi       : ' + targetDivisi.join(', ') +
    '\n\nSheet yang akan dibuat:\n' + namaSheetHasil.join('\n') +
    '\n\n⚠ Sheet yang sudah ada akan di-overwrite!\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi !== ui.Button.YES) return;

  const dataSumber  = sheetSumber.getDataRange().getValues();
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const masterData  = masterSheet.getRange('A4:D200').getValues()
    .filter(r => r[0] !== '' && r[1] !== '');

  const divisiMap = {};
  masterData.forEach(r => {
    const div  = String(r[0]).trim();
    const nama = String(r[1]).trim();
    if (div && nama) divisiMap[nama] = div;
  });

  const namaPerDivisi = {};
  for (let i = 3; i < dataSumber.length; i++) {
    const nama = String(dataSumber[i][COL_NAMA - 1]).trim();
    if (!nama) continue;
    const div = divisiMap[nama] || '';
    if (!div) continue;
    if (!namaPerDivisi[div]) namaPerDivisi[div] = [];
    if (!namaPerDivisi[div].includes(nama)) namaPerDivisi[div].push(nama);
  }

  const hasil = [];
  const MAX_OT_FRACTION  = 78 / 24;
  const OT_MEAL_FRACTION = 11 / 24;

  for (const divisi of targetDivisi) {
    const namaSheetRekap = 'Rekap_' + divisi + '_' + periode;
    const daftarNama     = namaPerDivisi[divisi] || [];

    if (daftarNama.length === 0) {
      hasil.push('⚠ ' + divisi + ': tidak ada karyawan ditemukan'); continue;
    }

    const sheetLama = ss.getSheetByName(namaSheetRekap);
    if (sheetLama) ss.deleteSheet(sheetLama);

    const rekapSheet = ss.insertSheet(namaSheetRekap);
    rekapSheet.setTabColor('#FF5722');
    rekapSheet.setHiddenGridlines(true);

    const headers = [
      'Name','Division','TTL Regular HRS','TTL O/T 1st HRS',
      'TTL O/T After 1st HRS','Bonus TTL O/T After 1st HRS',
      'Sunday/Red Day TTL HRS','Sunday/Red Day TTL O/T 1st HRS',
      'Sunday/Red Day TTL O/T After 1st HRS','TTL Working HRS for Capacity',
      'TTL Day Reg Meal Allowance','Sunday/Red Day Meal Allowance',
      'TTL OT Meal','Deduction Skill Bonus Based on Day Come In',
    ];
    rekapSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]).setBackground('#178232').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    const colWidths = [150,100,130,130,150,200,180,200,220,180,180,200,120,280];
    colWidths.forEach((w, i) => rekapSheet.setColumnWidth(i + 1, w));
    rekapSheet.setRowHeight(1, 44);
    rekapSheet.setFrozenRows(1);
    rekapSheet.setFrozenColumns(2);

    const ref = "'" + namaSumber + "'";

    for (let idx = 0; idx < daftarNama.length; idx++) {
      const nama = daftarNama[idx];
      const row  = idx + 2;
      const n    = `A${row}`;

      rekapSheet.getRange(row, 1).setValue(nama);
      rekapSheet.getRange(row, 2).setValue(divisi);

      rekapSheet.getRange(row, 3).setFormula(
        `=SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday")` +
        `+SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!Q:Q,"SWAP")` +
        `+SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!Q:Q,"HALF DAY SUNDAY")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!P:P,"RED DAY DOUBLE")`
      );
      rekapSheet.getRange(row, 4).setFormula(
        `=SUMIFS(${ref}!N:N,${ref}!C:C,${n})` +
        `-SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!B:B,"Sunday")` +
        `-SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!P:P,"RED DAY DOUBLE")` +
        `+SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!Q:Q,"SWAP")`
      );

      const otAfter =
        `(SUMIFS(${ref}!O:O,${ref}!C:C,${n})` +
        `-SUMIFS(${ref}!O:O,${ref}!C:C,${n},${ref}!B:B,"Sunday")` +
        `-SUMIFS(${ref}!O:O,${ref}!C:C,${n},${ref}!P:P,"RED DAY DOUBLE")` +
        `+SUMIFS(${ref}!O:O,${ref}!C:C,${n},${ref}!Q:Q,"SWAP"))`;
      const ot1 =
        `(SUMIFS(${ref}!N:N,${ref}!C:C,${n})` +
        `-SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!B:B,"Sunday")` +
        `-SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!P:P,"RED DAY DOUBLE")` +
        `+SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!Q:Q,"SWAP"))`;
      const bonusOtFormula = `MAX(0,${ot1}+${otAfter}-${MAX_OT_FRACTION})`;

      rekapSheet.getRange(row, 5).setFormula(`=${otAfter}-${bonusOtFormula}`);
      rekapSheet.getRange(row, 6).setFormula(`=${bonusOtFormula}`);
      rekapSheet.getRange(row, 7).setFormula(
        `=SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"Sunday",${ref}!Q:Q,"DOUBLE")`
      );
      rekapSheet.getRange(row, 8).setFormula(
        `=SUMIFS(${ref}!N:N,${ref}!C:C,${n},${ref}!B:B,"Sunday",${ref}!Q:Q,"DOUBLE")`
      );
      rekapSheet.getRange(row, 9).setFormula(
        `=SUMIFS(${ref}!O:O,${ref}!C:C,${n},${ref}!B:B,"Sunday",${ref}!Q:Q,"DOUBLE")`
      );
      rekapSheet.getRange(row, 10).setFormula(
        `=SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"VACATION PAID")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"FLEX DAY")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"ADDITIONAL PAID")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"MATERNITY LEAVE")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"RED DAY")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"SICK PAID")` +
        `-SUMIFS(${ref}!M:M,${ref}!C:C,${n},${ref}!B:B,"<>Sunday",${ref}!P:P,"SICK UNPAID")`
      );
      rekapSheet.getRange(row, 11).setFormula(
        `=COUNTIFS(${ref}!C:C,${n},${ref}!B:B,"<>Sunday")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"SWAP")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"DOUBLE")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"HALF DAY SUNDAY")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"SICK UNPAID")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"SICK UNPAID")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"DAY OFF UNPAID")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"DAY OFF UNPAID")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"HALF DAY")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"HALF DAY")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"RED DAY DOUBLE")` +
        `-COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"RED DAY DOUBLE")`
      );
      rekapSheet.getRange(row, 12).setFormula(
        `=COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"DOUBLE")`
      );
      rekapSheet.getRange(row, 13).setFormula(
        `=COUNTIFS(${ref}!C:C,${n},${ref}!L:L,">"&${OT_MEAL_FRACTION})`
      );
      rekapSheet.getRange(row, 14).setFormula(
        `=COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"SICK PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"VACATION PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"FLEX DAY")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"ADDITIONAL PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"RED DAY")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!P:P,"MATERNITY LEAVE")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"SICK PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"VACATION PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"FLEX DAY")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"ADDITIONAL PAID")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"RED DAY")` +
        `+COUNTIFS(${ref}!C:C,${n},${ref}!Q:Q,"MATERNITY LEAVE")`
      );

      const bg = idx % 2 === 0 ? '#F7FDFB' : '#FFFFFF';
      rekapSheet.getRange(row, 1, 1, headers.length)
        .setBackground(bg).setHorizontalAlignment('center').setVerticalAlignment('middle');
      rekapSheet.getRange(row, 1, 1, 2).setFontWeight('bold').setHorizontalAlignment('left');
      rekapSheet.setRowHeight(row, 22);
    }

    // Format kolom jam [h]:mm
    const totalDataRows = daftarNama.length;
    rekapSheet.getRange(2, 3, totalDataRows, 8).setNumberFormat('[h]:mm');

    // Baris total
    const totalRow = daftarNama.length + 2;
    rekapSheet.getRange(totalRow, 1).setValue('TOTAL');
    rekapSheet.getRange(totalRow, 2).setValue(divisi);
    for (let col = 3; col <= headers.length; col++) {
      const colLetter = columnToLetter(col);
      rekapSheet.getRange(totalRow, col).setFormula(
        `=SUM(${colLetter}2:${colLetter}${totalRow - 1})`
      );
    }
    rekapSheet.getRange(totalRow, 3, 1, 8).setNumberFormat('[h]:mm');
    rekapSheet.getRange(totalRow, 1, 1, headers.length)
      .setBackground('#0F6E56').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center');
    rekapSheet.getRange(1, 1, totalRow, headers.length)
      .setBorder(true,true,true,true,true,true,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);

    hasil.push('✓ ' + namaSheetRekap + ' — ' + daftarNama.length + ' karyawan');
    Logger.log('generateTemplateRekap selesai: ' + namaSheetRekap);
  }

  ui.alert(
    '✅ Template Rekap selesai!\n\n' +
    'Sheet sumber : ' + namaSumber + '\nPeriode      : ' + periode + '\n\n' +
    hasil.join('\n') + '\n\n' +
    'Rumus Excel sudah terhubung ke sheet sumber.\n' +
    'Jika data berubah, rekap otomatis update.'
  );
}

// ── buatSheetRentang — Sheet data mentah berbasis rumus QUERY ─────────
// Input: start date + end date (DD/MM/YYYY) via prompt + nama sheet output.
// Script hanya membuat sheet dan menulis 1 rumus QUERY yang menarik data
// dari semua sheet divisi yang relevan — data otomatis update jika sumber berubah.
function buatSheetRentang() {
  _requireAdmin();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Input tanggal mulai ──────────────────────────────────────────
  const inputStart = ui.prompt(
    '📅 Siapkan Rentang Tanggal (1/3)',
    'Masukkan tanggal MULAI:\nFormat: DD/MM/YYYY\nContoh: 01/03/2026',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputStart.getSelectedButton() !== ui.Button.OK) return;
  const startDate = _parseTanggal(inputStart.getResponseText().trim());
  if (!startDate) { ui.alert('❌ Format tanggal tidak valid. Gunakan DD/MM/YYYY'); return; }

  // ── 2. Input tanggal selesai ────────────────────────────────────────
  const inputEnd = ui.prompt(
    '📅 Siapkan Rentang Tanggal (2/3)',
    'Masukkan tanggal SELESAI:\nFormat: DD/MM/YYYY\nContoh: 30/04/2026',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputEnd.getSelectedButton() !== ui.Button.OK) return;
  const endDate = _parseTanggal(inputEnd.getResponseText().trim());
  if (!endDate) { ui.alert('❌ Format tanggal tidak valid. Gunakan DD/MM/YYYY'); return; }
  if (endDate < startDate) { ui.alert('❌ Tanggal selesai tidak boleh lebih awal dari tanggal mulai.'); return; }

  const labelStart = Utilities.formatDate(startDate, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const labelEnd   = Utilities.formatDate(endDate,   CONFIG.TIMEZONE, 'dd/MM/yyyy');

  // ── 3. Input nama sheet output ──────────────────────────────────────
  const inputNama = ui.prompt(
    '📅 Siapkan Rentang Tanggal (3/3)',
    'Masukkan nama sheet output:\nContoh: Data_Mar-Apr_2026\n\n⚠ Sheet lama akan di-overwrite.',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputNama.getSelectedButton() !== ui.Button.OK) return;
  const namaOutput = inputNama.getResponseText().trim();
  if (!namaOutput) { ui.alert('❌ Nama sheet tidak boleh kosong.'); return; }

  // ── 4. Temukan sheet sumber yang ada ───────────────────────────────
  const bulan        = _getBulanDalamRentang(startDate, endDate);
  const namaSheetSrc = [];

  for (const divisi of CONFIG.DIVISI) {
    for (const { month, year } of bulan) {
      const tglRef    = new Date(year, month - 1, 1);
      const namaBulan = Utilities.formatDate(tglRef, CONFIG.TIMEZONE, 'MMM_yyyy');
      const nama      = divisi + '_' + namaBulan;
      if (ss.getSheetByName(nama)) namaSheetSrc.push(nama);
    }
    // Fallback sheet tanpa suffix bulan
    if (ss.getSheetByName(divisi)) namaSheetSrc.push(divisi);
  }

  if (namaSheetSrc.length === 0) {
    ui.alert(
      '❌ Tidak ada sheet divisi ditemukan untuk rentang:\n' + labelStart + ' — ' + labelEnd +
      '\n\nPastikan sheet mengikuti format: [DIVISI]_[MMM]_[yyyy]\nContoh: WEB_Mar_2026'
    );
    return;
  }

  // ── 5. Susun rumus QUERY ────────────────────────────────────────────
  // Format tanggal untuk klausa WHERE QUERY: date 'YYYY-MM-DD'
  const qStart = Utilities.formatDate(startDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const qEnd   = Utilities.formatDate(endDate,   CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // Gabungkan semua sheet sumber menjadi satu array {Sheet1!A4:T; Sheet2!A4:T; ...}
  // Range A:T = 20 kolom (Tanggal sampai UPC/PC STATUS) — cocok dengan jumlah header
  // di sheet output. TOTAL_COL absensi adalah 22 (A:V); kolom U:V (CATATAN TELAT,
  // CATATAN PULANG AWAL) tidak ikut ditarik agar tampilan ringkas.
  const bagianRange = namaSheetSrc.map(n => "'" + n + "'!A4:T").join('; ');
  const dataArray   = namaSheetSrc.length === 1
    ? "'" + namaSheetSrc[0] + "'!A4:T"
    : '{' + bagianRange + '}';

  const queryStr =
    'SELECT * WHERE Col1 >= date \'' + qStart + '\'' +
    ' AND Col1 <= date \'' + qEnd   + '\'' +
    ' AND Col3 IS NOT NULL' +
    ' ORDER BY Col1 ASC, Col3 ASC';

  const formula = '=IFERROR(QUERY(' + dataArray + ',"' + queryStr + '",0),"— Tidak ada data dalam rentang ini —")';

  // ── 6. Buat sheet output ────────────────────────────────────────────
  // Header kolom output (A:T = 20 kolom, sama dengan range QUERY di atas).
  // Pakai headers.length (bukan TOTAL_COL) karena output sheet ini hanya A:T
  // sementara TOTAL_COL untuk sheet absensi adalah 22 (A:V).
  const headers = [
    'Tanggal','Hari','Nama','Email',
    'Status','Masuk','Ist.1 Mulai','Ist.1 Selesai',
    'Ist.2 Mulai','Ist.2 Selesai','Pulang',
    'Jam Efektif','Regular Hrs','OT 1','OT 2',
    'NOTE','SUNDAY/RED DAY','KETERANGAN TIDAK HADIR','PLAN','UPC/PC STATUS',
  ];

  let outSheet = ss.getSheetByName(namaOutput);
  if (outSheet) ss.deleteSheet(outSheet);
  outSheet = ss.insertSheet(namaOutput);
  outSheet.setTabColor('#1565C0');
  outSheet.setHiddenGridlines(true);

  // Baris 1: judul
  outSheet.getRange(1, 1, 1, headers.length).merge()
    .setValue('DATA ABSENSI  |  ' + labelStart + '  →  ' + labelEnd)
    .setBackground('#0F6E56').setFontColor('#FFFFFF')
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center');
  outSheet.setRowHeight(1, 28);

  // Baris 2: keterangan sheet sumber
  outSheet.getRange(2, 1, 1, headers.length).merge()
    .setValue('Sumber: ' + namaSheetSrc.join(', '))
    .setBackground('#E8F5E9').setFontColor('#2E7D32')
    .setFontSize(8).setFontStyle('italic');
  outSheet.setRowHeight(2, 16);

  // Baris 3: header kolom (sama dengan sheet divisi)
  outSheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#1D9E75').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,true,true,'#B0D9C8',SpreadsheetApp.BorderStyle.SOLID);
  outSheet.setRowHeight(3, 32);
  outSheet.setFrozenRows(3);

  // Lebar kolom
  const colWidths = [90,80,130,180,70,70,90,90,90,90,70,100,60,60,60,120,140,180,120,140];
  colWidths.forEach((w, i) => outSheet.setColumnWidth(i + 1, w));

  // Baris 4: rumus QUERY — satu rumus untuk semua data
  outSheet.getRange(4, 1).setFormula(formula);

  // Pre-format kolom agar tampilan langsung benar saat data muncul
  outSheet.getRange(4, 1,  500, 1).setNumberFormat('DD/MM/YYYY');
  outSheet.getRange(4, 6,  500, 6).setNumberFormat('HH:mm');     // F:K jam
  outSheet.getRange(4, 12, 500, 4).setNumberFormat('[h]:mm');    // L:O jam efektif

  // Warna merah untuk baris hari Minggu. Data diisi lewat rumus QUERY (jumlah
  // baris dinamis), jadi tidak bisa diwarnai manual — pakai conditional
  // formatting yang otomatis berlaku saat kolom B (Hari) = "Sunday".
  const sundayRange = outSheet.getRange('A4:T2000');
  const sundayRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B4="Sunday"')
    .setBackground('#FFCDD2')
    .setFontColor('#B71C1C')
    .setRanges([sundayRange])
    .build();
  const rules = outSheet.getConditionalFormatRules();
  rules.push(sundayRule);
  outSheet.setConditionalFormatRules(rules);

  ui.alert(
    '✅ Sheet berhasil dibuat!\n\n' +
    'Sheet   : ' + namaOutput + '\n' +
    'Rentang : ' + labelStart + ' — ' + labelEnd + '\n\n' +
    '📌 Data diambil otomatis via rumus QUERY.\n' +
    'Jika data sumber diubah, sheet ini langsung update.\n\n' +
    'Sheet sumber (' + namaSheetSrc.length + '):\n' + namaSheetSrc.join('\n')
  );
}

// ── _parseTanggal — Parse "DD/MM/YYYY" → Date ─────────────────────────
function _parseTanggal(str) {
  if (!str || !/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return null;
  const [d, m, y] = str.split('/').map(Number);
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const tgl = new Date(y, m - 1, d, 12, 0, 0);
  // Validasi overflow (misal 31/02 → jadi 03/03)
  if (tgl.getMonth() !== m - 1) return null;
  return tgl;
}

// ── _getBulanDalamRentang — Daftar bulan (month, year) dalam rentang ───
function _getBulanDalamRentang(startDate, endDate) {
  const hasil = [];
  let cur     = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const akhir = new Date(endDate.getFullYear(),   endDate.getMonth(),   1);
  while (cur <= akhir) {
    hasil.push({ month: cur.getMonth() + 1, year: cur.getFullYear() });
    cur.setMonth(cur.getMonth() + 1);
  }
  return hasil;
}


// ═══════════════════════════════════════════════════════════════════════
// Helper kalkulasi jam (dipanggil oleh buatSheetRekap & hitungRekap)
// ═══════════════════════════════════════════════════════════════════════

// Regular Hours = jam kerja normal (kurangi Sunday, tambah SWAP/HALF DAY SUNDAY,
// kurangi RED DAY DOUBLE karena sudah masuk OT)
function calculateRegularHours(regularHrs, namahari, note, sundayRedDayNote) {
  let total = regularHrs;
  if (namahari.toLowerCase() === 'sunday')                  total -= regularHrs;
  if (sundayRedDayNote.toUpperCase() === 'SWAP')            total += regularHrs;
  if (sundayRedDayNote.toUpperCase() === 'HALF DAY SUNDAY') total += regularHrs;
  if (note.toUpperCase() === 'RED DAY DOUBLE')              total -= regularHrs;
  return total;
}

// OT 1st Hours = lembur tier 1 (kurangi Sunday & RED DAY DOUBLE, tambah SWAP)
function calculateOt1Hrs(ot1Hrs, namaHari, note, sundayRedDayNote) {
  let total = ot1Hrs;
  if (namaHari.toLowerCase() === 'sunday')      total -= ot1Hrs;
  if (note.toUpperCase() === 'RED DAY DOUBLE')  total -= ot1Hrs;
  if (sundayRedDayNote.toUpperCase() === 'SWAP') total += ot1Hrs;
  return total;
}

// OT After 1st Hours = lembur setelah tier 1
// Kembalikan { totalOtAfter1stHrs, bonusOtAfter1stHrs } — bonus = kelebihan batas 78 jam
function calculateOtAfter1stHrs(ot1Hrs, otAfter1stHrs, namaHari, note, sundayRedDayNote) {
  let total  = otAfter1stHrs;
  const maxOT = 78;
  if (namaHari.toLowerCase() === 'sunday')       total -= otAfter1stHrs;
  if (note.toUpperCase() === 'RED DAY DOUBLE')   total -= otAfter1stHrs;
  if (sundayRedDayNote.toUpperCase() === 'SWAP') total += otAfter1stHrs;

  const ttlOt1  = calculateOt1Hrs(ot1Hrs, namaHari, note, sundayRedDayNote);
  const totalOT = total + ttlOt1;
  const bonusOt = totalOT > maxOT ? totalOT - maxOT : 0;

  return {
    totalOtAfter1stHrs: total - bonusOt,
    bonusOtAfter1stHrs: bonusOt,
  };
}

// Sunday/Red Day TTL HRS — jam kerja hari Minggu jika DOUBLE
function calculateSundayRedDayHrs(regularHrs, namaHari, sundayRedDayNote) {
  if (namaHari.toLowerCase() === 'sunday' && sundayRedDayNote.toUpperCase() === 'DOUBLE') {
    return regularHrs;
  }
  return 0;
}

// Sunday/Red Day OT 1st HRS — OT tier 1 hari Minggu jika DOUBLE
function calculateSundayRedDayOt1stHrs(ot1Hrs, namaHari, sundayRedDayNote) {
  if (namaHari.toLowerCase() === 'sunday' && sundayRedDayNote.toUpperCase() === 'DOUBLE') {
    return ot1Hrs;
  }
  return 0;
}

// Sunday/Red Day OT After 1st HRS — OT setelah tier 1 hari Minggu jika DOUBLE
function calculateSundayRedDayOtAfter1stHrs(otAfter1stHrs, namaHari, sundayRedDayNote) {
  if (namaHari.toLowerCase() === 'sunday' && sundayRedDayNote.toUpperCase() === 'DOUBLE') {
    return otAfter1stHrs;
  }
  return 0;
}

// Capacity Hours = jam kerja yang dihitung untuk kapasitas produksi
// (kecuali hari Sunday dan note tertentu seperti liburan/sakit)
function calculateCapacityHrs(regularHrs, namaHari, note) {
  if (namaHari.toLowerCase() === 'sunday') return 0;
  const excludedNotes = [
    'VACATION PAID','FLEX DAY','ADDITIONAL PAID','MATERNITY LEAVE',
    'RED DAY','SICK PAID','SICK UNPAID',
  ];
  if (excludedNotes.includes(note.toUpperCase())) return 0;
  return regularHrs;
}

// Day Reg Meal Allowance = jumlah hari dapat uang makan regular
function calculateDayRegMealAllowance(namaHari, sundayRedDayNote, note) {
  let count = 0;
  if (namaHari && namaHari !== '') count += 1;
  if (namaHari.toLowerCase() === 'sunday') count -= 1;
  if (sundayRedDayNote.toUpperCase() === 'SWAP')            count += 1;
  if (sundayRedDayNote.toUpperCase() === 'DOUBLE')          count += 1;
  if (sundayRedDayNote.toUpperCase() === 'HALF DAY SUNDAY') count += 1;

  const deductTriggers = ['SICK UNPAID','DAY OFF UNPAID','HALF DAY','RED DAY DOUBLE'];
  for (const t of deductTriggers) {
    if (note.toUpperCase() === t || sundayRedDayNote.toUpperCase() === t) count -= 1;
  }
  return count;
}

// Sunday/Red Day Meal Allowance = hari dapat uang makan hari Minggu/merah
// Hanya saat Q=DOUBLE (worker benar-benar kerja Sunday/Red Day).
// Sama dengan formula template: =COUNTIFS(Q:Q,"DOUBLE")
function calculateSundayRedDayMealAllowance(namaHari, sundayRedDayNote) {
  if (sundayRedDayNote.toUpperCase() === 'DOUBLE') return 1;
  return 0;
}

// OT Meal = hari dapat uang makan lembur (jam efektif > 11 jam)
function calculateOtMeal(effectiveHrs) {
  return effectiveHrs > 11 ? 1 : 0;
}

// Deduction Skill Bonus = hari yang mengurangi skill bonus
function calculateDeductionSkillBonus(note, sundayRedDayNote) {
  const deductionNotes = [
    'SICK PAID','VACATION PAID','FLEX DAY','ADDITIONAL PAID','RED DAY','MATERNITY LEAVE',
  ];
  let count = 0;
  if (deductionNotes.includes(note.toUpperCase()))             count += 1;
  if (sundayRedDayNote !== '' &&
      deductionNotes.includes(sundayRedDayNote.toUpperCase())) count += 1;
  return count;
}
