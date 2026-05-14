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
  _loadSettings();
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
      [5,  9, 'PUTIH = bisa diedit',  '#FFFFFF',  '#2C2C2A'],
      [14, 4, 'UNGU = formula auto',  '#EEEDFE',  '#534AB7'],
      [18, 6, 'KUNING = hari ini',    '#FFF9C4',  '#633806'],
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
      ['Ist. Ketiga\nMulai',                '#E1F5EE', '#085041'],
      ['Ist. Ketiga\nSelesai',              '#E1F5EE', '#085041'],
      ['Pulang',                            '#E1F5EE', '#085041'],
      ['Jam Efektif 🔒',                    '#1D9E75', '#FFFFFF'],
      ['Regular Hours',                     '#E1F5EE', '#085041'],
      ['OT 1',                              '#E1F5EE', '#085041'],
      ['OT 2',                              '#E1F5EE', '#085041'],
      ['NOTE',                              '#E1F5EE', '#085041'],
      ['SUNDAY/RED DAY\nFILL: DOUBLE/SWAP', '#E1F5EE', '#085041'],
      ['Keterangan\nTidak Hadir',           '#E1F5EE', '#085041'],
      ['Plan',                              '#E1F5EE', '#085041'],
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

    // Lebar kolom A–W
    const colWidths = [90,80,130,180,70,70,90,90,90,90, 90,90, 70,100,60,60,60,120,140,160,80,100,100];
    colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(3);

    // Kunci header — owner dan admin
    const headerProt = sheet.getRange('A1:W3').protect();
    headerProt.setDescription('Header — owner dan admin');
    headerProt.setWarningOnly(false);
    headerProt.removeEditors(headerProt.getEditors());
    headerProt.addEditor(Session.getEffectiveUser());
    for (const adminEmail of CONFIG.ADMIN_EMAILS) {
      try { headerProt.addEditor(adminEmail); } catch(e) {}
    }

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

// ── _setProteksiSettingsJam — Pasang proteksi range A2:C9 ─────────────
// Private helper dipanggil oleh buatSheetSettings() dan perbaruiAksesSettings().
// Range A2:C9 (row jam: MASUK, IST1/2/3 mulai/selesai, PULANG) bisa diedit
// oleh: owner + admin + semua email worker (Master_Data kolom C) +
// semua email asisten (kolom E). Range protection ini override sheet
// protection admin-only di area row jam.
function _setProteksiSettingsJam(sheet) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const owner  = Session.getEffectiveUser();

  // Hapus proteksi range A2:C9 lama (kalau ada) — bikin idempotent
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .find(p => {
      const r = p.getRange();
      return r.getRow() === 2 && r.getLastRow() === 9 &&
             r.getColumn() === 1 && r.getLastColumn() === 3;
    });
  if (existing) existing.remove();

  // Buat proteksi range baru
  const jamProt = sheet.getRange('A2:C9').protect();
  jamProt.setDescription('_Settings jam — admin + worker + asisten');
  jamProt.setWarningOnly(false);
  jamProt.removeEditors(jamProt.getEditors());
  jamProt.addEditor(owner);

  // Tambah admin emails
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { jamProt.addEditor(adminEmail); } catch(e) {
      Logger.log('⚠ Gagal tambah admin ' + adminEmail + ' ke _Settings A2:C9: ' + e.message);
    }
  }

  // Tambah worker + asisten dari Master_Data
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  let countWorker = 0;
  let countAsisten = 0;
  if (master) {
    const masterData = master.getRange('A4:E200').getValues()
      .filter(r => r[0] !== '' && String(r[3]).trim().toUpperCase() === 'TRUE');
    for (const k of masterData) {
      const workerEmail  = String(k[2] || '').trim();
      const asistenEmail = String(k[4] || '').trim();
      if (workerEmail) {
        try {
          jamProt.addEditor(workerEmail);
          countWorker++;
        } catch(e) {
          Logger.log('⚠ Gagal tambah worker ' + workerEmail + ' ke _Settings: ' + e.message);
        }
      }
      if (asistenEmail) {
        try {
          jamProt.addEditor(asistenEmail);
          countAsisten++;
        } catch(e) {
          Logger.log('⚠ Gagal tambah asisten ' + asistenEmail + ' ke _Settings: ' + e.message);
        }
      }
    }
  } else {
    Logger.log('⚠ Master_Data tidak ditemukan — _Settings A2:C9 hanya editable admin');
  }

  Logger.log('✓ _Settings A2:C9: editor = owner + ' +
    CONFIG.ADMIN_EMAILS.length + ' admin + ' +
    countWorker + ' worker + ' + countAsisten + ' asisten');

  return { countWorker, countAsisten };
}

// ── buatSheetSettings — Buat/reset sheet _Settings di spreadsheet ini ──
// Sheet berisi key-value yang dibaca _loadSettings() setiap fungsi jalan.
// Admin cukup ubah kolom VALUE — berlaku di append/trigger berikutnya.
function buatSheetSettings() {
  _requireAdmin();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const divisi = (CONFIG.DIVISI[0] || '').toUpperCase();
  const auto   = (CONFIG.AUTO_ABSENSI || {})[divisi] || {};
  const ui     = SpreadsheetApp.getUi();

  // Konfirmasi jika sheet sudah ada
  const existing = ss.getSheetByName('_Settings');
  if (existing) {
    const res = ui.alert(
      'Sheet _Settings sudah ada.',
      'Reset dan buat ulang? Perubahan yang belum disimpan akan hilang.',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }

  const sheet = ss.insertSheet('_Settings');
  sheet.setTabColor('#F57F17');
  sheet.setHiddenGridlines(true);

  // ── Header ──
  sheet.getRange('A1:C1')
    .setValues([['KEY', 'VALUE', 'Keterangan']])
    .setBackground('#F57F17').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);

  // ── Data ──
  const rows = [
    // Jam auto-absensi
    ['MASUK',              auto.masuk        || '', 'Jam masuk default (HH:mm) — kosongkan jika staf isi sendiri'],
    ['IST1_MULAI',         auto.ist1Mulai    || '', 'Istirahat pertama mulai (HH:mm) — kosongkan jika tidak ada'],
    ['IST1_SELESAI',       auto.ist1Selesai  || '', 'Istirahat pertama selesai (HH:mm)'],
    ['IST2_MULAI',         auto.ist2Mulai    || '', 'Istirahat kedua mulai (HH:mm) — kosongkan jika tidak ada'],
    ['IST2_SELESAI',       auto.ist2Selesai  || '', 'Istirahat kedua selesai (HH:mm)'],
    ['IST3_MULAI',         auto.ist3Mulai    || '', 'Istirahat ketiga mulai (HH:mm) — kosongkan jika tidak ada'],
    ['IST3_SELESAI',       auto.ist3Selesai  || '', 'Istirahat ketiga selesai (HH:mm)'],
    ['PULANG',             auto.pulang       || '', 'Jam pulang default (HH:mm) — kosongkan jika staf isi sendiri'],
    // Separator kosong
    ['', '', ''],
    // Pengaturan operasional
    ['DIVISI',             divisi,                                     'Nama divisi ini — HURUF KAPITAL. Ubah jika spreadsheet di-duplicate untuk divisi baru (contoh: FINANCE)'],
    ['JAM_REMINDER',       String(CONFIG.JAM_REMINDER        || 17),  'Jam pengingat belum isi pulang (angka 0–23)'],
    ['SELISIH_MENIT_LOCK', String(CONFIG.SELISIH_MENIT_LOCK  || 30),  'Menit setelah pulang → baris terkunci otomatis'],
    ['PLAN_JAM',           (CONFIG.PLAN_JAM  || []).join(','),         'Opsi plan jam kerja, pisah koma (contoh: 07:00-16:00,08:00-17:00)'],
    ['ADMIN_EMAILS',       (CONFIG.ADMIN_EMAILS || []).join(','),      'Email admin, pisah koma'],
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Baris 2–8 (MASUK s/d IST3_SELESAI) — kolom VALUE format HH:mm
  // agar Sheets menyimpan sebagai time serial, bukan string, sehingga
  // formula jam kerja di sheet absensi bisa menghitung dengan benar
  sheet.getRange(2, 2, 7, 1).setNumberFormat('HH:mm');

  // ── Styling ──
  const keyRange = sheet.getRange(2, 1, rows.length, 1);
  keyRange.setBackground('#FFF3E0').setFontWeight('bold').setFontSize(10);

  const valRange = sheet.getRange(2, 2, rows.length, 1);
  valRange.setBackground('#FFFFFF').setFontSize(10);

  const ketRange = sheet.getRange(2, 3, rows.length, 1);
  ketRange.setBackground('#FAFAFA').setFontColor('#888').setFontSize(9).setFontStyle('italic');

  // Baris separator (baris 7) — abu
  sheet.getRange(7, 1, 1, 3).setBackground('#F5F5F5');

  // Lebar kolom
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 360);
  sheet.setFrozenRows(1);

  // Border seluruh tabel
  sheet.getRange(1, 1, rows.length + 1, 3)
    .setBorder(true, true, true, true, true, true,
      '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Proteksi: 2 lapis ───────────────────────────────────────────────
  // 1. Sheet-level (admin only) — default-deny untuk seluruh sheet
  // 2. Range A2:C9 (jam) — admin + worker + asisten (override sheet protect)
  const owner = Session.getEffectiveUser();

  const sheetProt = sheet.protect();
  sheetProt.setDescription('_Settings — hanya admin yang bisa edit');
  sheetProt.setWarningOnly(false);
  sheetProt.removeEditors(sheetProt.getEditors());
  sheetProt.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { sheetProt.addEditor(adminEmail); } catch(e) {}
  }

  const { countWorker, countAsisten } = _setProteksiSettingsJam(sheet);

  ui.alert(
    '✅ Sheet _Settings berhasil dibuat!\n\n' +
    'Edit kolom VALUE untuk mengubah setting.\n' +
    'Perubahan berlaku otomatis di append/trigger berikutnya.\n\n' +
    'Akses:\n' +
    '• Row 2–9 (jam): admin + ' + countWorker + ' worker + ' + countAsisten + ' asisten\n' +
    '• Row 1, 10–15: admin only'
  );
}

// ── perbaruiAksesSettings — Refresh editor A2:C9 dari Master_Data ─────
// Dipakai ketika Master_Data berubah (worker baru, asisten ganti, dll.)
// tanpa harus reset seluruh sheet _Settings. Idempotent.
function perbaruiAksesSettings() {
  _requireAdmin();
  _loadSettings();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Settings');
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '⚠ Sheet _Settings tidak ditemukan.\n\n' +
      'Jalankan "⚙️ Buat/Reset Sheet Settings" dulu.'
    );
    return;
  }

  const { countWorker, countAsisten } = _setProteksiSettingsJam(sheet);

  SpreadsheetApp.getUi().alert(
    '✅ Akses _Settings diperbarui!\n\n' +
    'Row 2–9 (jam) editor:\n' +
    '• Owner + ' + CONFIG.ADMIN_EMAILS.length + ' admin\n' +
    '• ' + countWorker + ' worker dari Master_Data\n' +
    '• ' + countAsisten + ' asisten dari Master_Data'
  );
}

// ── setupProteksiMaster — Kunci sheet Master_Data ─────────────────────
// Hapus proteksi lama lalu kunci seluruh sheet untuk owner saja
function setupProteksiMaster() {
  _loadSettings();
  _requireAdmin();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) return;

  master.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(p => p.remove());

  const prot = master.protect();
  prot.setDescription('HRD/owner/admin yang bisa edit Master_Data');
  prot.setWarningOnly(false);
  prot.removeEditors(prot.getEditors());
  prot.addEditor(Session.getEffectiveUser());
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { prot.addEditor(adminEmail); } catch(e) {}
  }
  Logger.log('Proteksi Master_Data selesai.');
}

// ── perbaruiProteksiAdmin — Tambah admin ke semua proteksi yang sudah ada ──
// Dipakai saat ADMIN_EMAILS diperbarui di _Settings agar editor lama
// (Master_Data, _Settings, header sheet bulan) langsung mendapat akses baru.
function perbaruiProteksiAdmin() {
  _requireAdmin();
  _loadSettings();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const owner   = Session.getEffectiveUser();
  const admins  = CONFIG.ADMIN_EMAILS;
  const hasil   = [];

  function _terapkan(prot, label) {
    prot.removeEditors(prot.getEditors());
    prot.addEditor(owner);
    for (const e of admins) {
      try { prot.addEditor(e); } catch(err) {}
    }
    Logger.log('✓ Proteksi diperbarui: ' + label);
  }

  // Master_Data
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (master) {
    master.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = master.protect();
    p.setDescription('HRD/owner/admin yang bisa edit Master_Data');
    p.setWarningOnly(false);
    _terapkan(p, 'Master_Data');
    hasil.push('✓ Master_Data');
  }

  // _Settings
  const settings = ss.getSheetByName('_Settings');
  if (settings) {
    settings.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = settings.protect();
    p.setDescription('_Settings — hanya admin yang bisa edit');
    p.setWarningOnly(false);
    _terapkan(p, '_Settings');
    hasil.push('✓ _Settings');
  }

  // Header semua sheet divisi (baris 1–3)
  const divisiPrefix = CONFIG.DIVISI.map(d => d + '_');
  ss.getSheets().forEach(sheet => {
    const nama = sheet.getName();
    const isDivisi = CONFIG.DIVISI.some(d => nama === d || nama.startsWith(d + '_'));
    if (!isDivisi) return;

    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      if (p.getRange().getRow() <= 3) {
        _terapkan(p, nama + ' header');
      }
    });
    hasil.push('✓ ' + nama + ' (header)');
  });

  SpreadsheetApp.getUi().alert(
    '✅ Proteksi diperbarui!\n\n' +
    'Admin emails:\n' + admins.join('\n') + '\n\n' +
    hasil.join('\n')
  );
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

  // G–M: opsional, jika diisi harus HH:MM (Ist 1/2/3 + Pulang)
  ['G','H','I','J','K','L','M'].forEach((col, idx) => {
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
  // Range diperluas ke kolom E untuk membaca Email Asisten (opsional).
  const masterData = master.getRange('A4:E200').getValues()
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

  // Step 4: Proteksi E:K per staf (staf hanya bisa edit jam absen barisnya sendiri)
  // Kolom L:O (formula) dan P:Q (admin) diproteksi terpisah di bawah
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
      baris.length, COL_PULANG - COL_STATUS + 1  // E sampai K
    );

    const prot = range.protect();
    prot.setDescription(nama + ' — ' + email +
      ' (baris ' + baris[0] + '–' + baris[baris.length - 1] + ')');
    prot.setWarningOnly(false);
    prot.removeEditors(prot.getEditors());
    prot.addEditor(owner);

    // Tambah semua admin email
    for (const adminEmail of CONFIG.ADMIN_EMAILS) {
      try { prot.addEditor(adminEmail); } catch(err) {
        Logger.log('⚠ Gagal tambah admin ' + adminEmail + ': ' + err.message);
      }
    }

    try {
      prot.addEditor(email);
      berhasil++;
      Logger.log('✓ Proteksi: ' + nama + ' baris ' + baris[0] + '–' + baris[baris.length - 1]);
    } catch(e) {
      Logger.log('⚠ Gagal tambah editor ' + email + ': ' + e.message);
    }

    // Email asisten (kolom E Master_Data) — opsional, skip kalau kosong.
    // Mapping 1-ke-1: asisten cuma boleh edit baris worker spesifiknya.
    const asisten = String(k[4] || '').trim();
    if (asisten) {
      try {
        prot.addEditor(asisten);
        Logger.log('✓ Proteksi asisten: ' + nama + ' ← ' + asisten);
      } catch(err) {
        Logger.log('⚠ Gagal tambah asisten ' + asisten + ': ' + err.message);
      }
    }
  }

  // Step 5: Proteksi L:O — kolom formula, hanya owner + admin
  const rangeLO = sheet.getRange(startRow, COL_EFEKTIF, numRows, COL_OT2 - COL_EFEKTIF + 1);
  const protLO  = rangeLO.protect();
  protLO.setDescription('Formula L:O — owner + admin only — baris ' + startRow + '–' + endRow);
  protLO.setWarningOnly(false);
  protLO.removeEditors(protLO.getEditors());
  protLO.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try {
      protLO.addEditor(adminEmail);
    } catch(err) {
      Logger.log('⚠ Gagal tambah admin L:O ' + adminEmail + ': ' + err.message);
    }
  }
  Logger.log('✓ Proteksi L:O (formula, owner + admin) baris ' + startRow + '–' + endRow);

  // Step 6: Proteksi P:Q — hanya admin
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

// ── migrateSheetTambahIst3 — Migrasi sheet 21-col → 23-col ───────────
// Untuk sheet yang dibuat SEBELUM fitur Ist 3 ditambah:
//   1. Insert 2 kolom di posisi 11 (Google Sheets auto-shift formula
//      yang sudah ada — referensi K → otomatis jadi M setelah insert)
//   2. Tulis header "Ist. Ketiga Mulai/Selesai" di K3, L3 dengan
//      formatting sama dengan ist1/ist2
//   3. Set column width K, L = 90
//   4. Update merge baris 1 (judul) ke A1:W1
//   5. Reset legenda baris 2 dengan span posisi baru
//   6. Re-pasang formula baris hari ini (supaya pakai versi flat
//      baru yang subtract ist3)
//   7. Re-run proteksiBarisBaru untuk baris hari ini supaya range
//      proteksi range E:M tepat
// Idempotent: kalau sudah migrated (header K3 berisi "Ist. Ketiga"
// atau lastColumn >= 23), skip dengan log.
function migrateSheetTambahIst3() {
  _loadSettings();
  _requireAdmin();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hasil   = [];

  // Cari sheet divisi (nama persis CONFIG.DIVISI atau prefix DIVISI_)
  const sheets = ss.getSheets().filter(s => {
    const nama = s.getName();
    return CONFIG.DIVISI.some(d => nama === d || nama.startsWith(d + '_'));
  });

  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert('⚠ Tidak ada sheet divisi untuk dimigrasi.');
    return;
  }

  for (const sheet of sheets) {
    const namaSheet = sheet.getName();

    // Idempotency check
    const headerK3 = String(sheet.getRange(3, 11).getValue() || '').toLowerCase();
    if (headerK3.includes('ketiga') || sheet.getLastColumn() >= 23) {
      hasil.push('⏭ ' + namaSheet + ': sudah ter-migrate, skip');
      Logger.log('⏭ ' + namaSheet + ': sudah ter-migrate, skip');
      continue;
    }

    // 1. Insert 2 kolom di posisi 11 (sebelum kolom Pulang lama)
    sheet.insertColumnsBefore(11, 2);

    // 2. Tulis header K3, L3 dengan formatting sama dengan ist1/ist2
    sheet.getRange(3, 11)
      .setValue('Ist. Ketiga\nMulai')
      .setBackground('#E1F5EE').setFontColor('#085041')
      .setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setBorder(true,true,true,true,false,false,
        '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(3, 12)
      .setValue('Ist. Ketiga\nSelesai')
      .setBackground('#E1F5EE').setFontColor('#085041')
      .setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setBorder(true,true,true,true,false,false,
        '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);

    // 3. Set column width K, L = 90
    sheet.setColumnWidth(11, 90);
    sheet.setColumnWidth(12, 90);

    // 4. Update merge baris 1 (judul) — un-merge lalu re-merge A1:W1
    try {
      sheet.getRange(1, 1, 1, 23).breakApart();
    } catch(e) {}
    sheet.getRange(1, 1, 1, 23).merge();

    // 5. Update legenda baris 2 — un-merge dulu, lalu pasang ulang
    try {
      sheet.getRange(2, 1, 1, 23).breakApart();
    } catch(e) {}
    const legends = [
      [1,  4, 'ABU = sudah lewat',    '#F1EFE8', '#5F5E5A'],
      [5,  9, 'PUTIH = bisa diedit',  '#FFFFFF',  '#2C2C2A'],
      [14, 4, 'UNGU = formula auto',  '#EEEDFE',  '#534AB7'],
      [18, 6, 'KUNING = hari ini',    '#FFF9C4',  '#633806'],
    ];
    for (const [startCol, span, text, bg, fg] of legends) {
      sheet.getRange(2, startCol, 1, span).merge()
        .setValue(text).setBackground(bg).setFontColor(fg)
        .setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true,true,true,true,false,false,
          '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    }

    // 6 & 7: Re-pasang formula + re-run proteksi untuk baris hari ini
    const today = getToday();
    const lastRow = sheet.getLastRow();
    let startToday = -1;
    let numToday = 0;
    if (lastRow >= 4) {
      const dates = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        if (dates[i][0] instanceof Date && isSameDate(dates[i][0], today)) {
          if (startToday === -1) startToday = i + 4;
          numToday++;
        }
      }
    }

    if (startToday > 0 && numToday > 0) {
      _pasangFormulaBaris(sheet, startToday, numToday);
      Logger.log('✓ ' + namaSheet + ': formula baris hari ini di-refresh (' +
        startToday + '–' + (startToday + numToday - 1) + ')');

      const divisi = CONFIG.DIVISI.find(d =>
        namaSheet === d || namaSheet.startsWith(d + '_')
      );
      if (divisi) {
        proteksiBarisBaru(sheet, divisi, startToday, numToday);
        Logger.log('✓ ' + namaSheet + ': proteksi baris hari ini di-refresh');
      }
    }

    hasil.push('✓ ' + namaSheet + ': migrated (2 kolom ditambah di K, L)');
    Logger.log('✓ ' + namaSheet + ': migration selesai');
  }

  SpreadsheetApp.getUi().alert(
    '✅ Migrasi Ist 3 selesai!\n\n' + hasil.join('\n')
  );
}
