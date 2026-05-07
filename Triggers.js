// ═══════════════════════════════════════════════════════════════════════
// TRIGGERS.JS — Event trigger dan setup sistem
//
// Berisi:
//   _requireAdmin()   — guard: lempar error jika bukan admin
//   onEdit()          — guard proteksi edit per email
//   onOpen()          — buat menu staf + menu admin saat file dibuka
//   setupTrigger()    — daftarkan semua trigger harian ke Apps Script
//   setupAwal()       — inisialisasi satu kali: proteksi + trigger + sheet
// ═══════════════════════════════════════════════════════════════════════

// ── _requireAdmin — Guard akses admin ────────────────────────────────
// Dipanggil di awal setiap fungsi admin.
// Aman dipakai di menu item (bukan simple trigger) karena OAuth sudah penuh.
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

// ── onEdit — Guard proteksi edit ──────────────────────────────────────
// Mencegah staf mengedit kolom yang bukan haknya:
//   - Kolom A–D terkunci (tanggal, hari, nama, email)
//   - Kolom L (Jam Efektif) terkunci — formula otomatis
//   - Kolom P–Q (NOTE, SUNDAY) — hanya admin
//   - Kolom lain: hanya bisa edit baris milik sendiri (cocokkan email)
function onEdit(e) {
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

  // Kolom di luar batas edit (sebagai safety)
  if (col > COL_EDIT_END) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    return;
  }

  // Deteksi email user
  let emailUser = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!emailUser) {
    emailUser = Session.getActiveUser().getEmail().trim().toLowerCase();
  }

  // Admin boleh edit segalanya
  const isAdmin = CONFIG.ADMIN_EMAILS.map(a => a.toLowerCase()).includes(emailUser);
  if (isAdmin) {
    Logger.log('✓ Admin edit: ' + emailUser + ' baris ' + row);
    return;
  }

  // Kolom P (NOTE) dan Q (SUNDAY/RED DAY) — hanya admin
  if (col === COL_NOTE || col === COL_SUNDAY) {
    e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');
    try { SpreadsheetApp.getUi().alert('❌ Kolom NOTE dan SUNDAY/RED DAY hanya bisa diedit oleh admin.'); } catch(err) {}
    return;
  }

  // Jika email tidak terdeteksi, biarkan lewat
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

  Logger.log('✓ Edit valid: ' + emailUser + ' baris ' + row);
}

// ── onOpen — Buat menu kustom saat spreadsheet dibuka ─────────────────
// Menu "📋 Absensi Saya" tampil untuk semua pengguna.
// Menu "🔧 Admin" hanya muncul jika email terdeteksi sebagai admin.
// Catatan: simple trigger terkadang tidak dapat email non-owner —
// jika email kosong, menu admin ditampilkan dan keamanan dijaga _requireAdmin().
function onOpen() {
  _loadSettings();
  const ui = SpreadsheetApp.getUi();

  // ── Menu staf — selalu tampil ──
  ui.createMenu('📋 Absensi Saya')
    .addItem('📍 Ke baris saya hari ini',    'keBarisHariIni')
    .addSeparator()
    .addItem('✅ Stamp MASUK',               'stampMasuk')
    .addItem('☕ Stamp ISTIRAHAT 1 MULAI',   'stampIst1Mulai')
    .addItem('▶ Stamp ISTIRAHAT 1 SELESAI',  'stampIst1Selesai')
    .addItem('☕ Stamp ISTIRAHAT 2 MULAI',   'stampIst2Mulai')
    .addItem('▶ Stamp ISTIRAHAT 2 SELESAI',  'stampIst2Selesai')
    .addItem('🏁 Stamp PULANG',              'stampPulang')
    .addSeparator()
    .addItem('📊 Rekap absensi saya',        'cekRekapSaya')
    .addSeparator()
    .addItem('🔄 Refresh Menu',              'onOpen')
    .addToUi();

  // ── Deteksi admin (best-effort) ──
  // Keamanan sesungguhnya ada di _requireAdmin() saat fungsi dieksekusi
  let emailUser = '';
  try { emailUser = Session.getEffectiveUser().getEmail().trim().toLowerCase(); } catch(e) {}
  if (!emailUser) {
    try { emailUser = Session.getActiveUser().getEmail().trim().toLowerCase(); } catch(e) {}
  }
  const isAdmin = !emailUser ||
    CONFIG.ADMIN_EMAILS.map(a => a.toLowerCase()).includes(emailUser);

  if (!isAdmin) return;

  // ── Menu admin ──
  ui.createMenu('🔧 Admin')
    .addItem('📅 Buat Sheet Bulan Baru',      'buatSheetBulanBaru')
    .addItem('➕ Append Baris Hari Ini',       'appendHariIni')
    .addSeparator()
    .addItem('📋 Generate Template Rekap',     'generateTemplateRekap')
    .addItem('📅 Data Rentang Tanggal',        'buatSheetRentang')
    .addItem('📊 Rekap Rentang Tanggal',       'rekapRentangTanggal')
    .addSeparator()
    .addItem('⚠️ Cek Belum Isi Pulang',       'cekBelumIsiPulang')
    .addItem('🔒 Lock Baris Sudah Pulang',     'lockBarisWebSudahPulang')
    .addSeparator()
    .addItem('⚙️ Setup Awal (pertama kali)',   'setupAwal')
    .addItem('⏰ Setup Trigger',               'setupTrigger')
    .addItem('🛡️ Setup Proteksi Master',      'setupProteksiMaster')
    .addItem('🔑 Perbarui Akses Admin',        'perbaruiProteksiAdmin')
    .addItem('⚙️ Buat/Reset Sheet Settings',  'buatSheetSettings')
    .addToUi();
}

// ── setupTrigger — Daftarkan semua trigger harian ─────────────────────
function setupTrigger() {
  _requireAdmin();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Tanggal 1 tiap bulan jam 05:00 — buat sheet bulan baru
  ScriptApp.newTrigger('buatSheetBulanBaru')
    .timeBased().onMonthDay(1).atHour(5).create();

  // Setiap hari jam 06:00 — append baris staf hari ini
  ScriptApp.newTrigger('appendHariIni')
    .timeBased().everyDays(1).atHour(6).create();

  // Setiap hari jam JAM_REMINDER — reminder staf yang belum isi pulang
  ScriptApp.newTrigger('cekBelumIsiPulang')
    .timeBased().everyDays(1).atHour(CONFIG.JAM_REMINDER).create();

  // Setiap jam — kunci baris 30 menit setelah jam pulang diisi
  ScriptApp.newTrigger('lockBarisWebSudahPulang')
    .timeBased().everyHours(1).create();

  // onEdit installable — guard proteksi per email
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit().create();

  SpreadsheetApp.getUi().alert(
    '✅ Trigger aktif!\n\n' +
    '• Tgl 1 tiap bulan 05:00 — buat sheet bulan baru\n' +
    '• Setiap hari 06:00 — append baris hari ini otomatis\n' +
    '• Setiap hari ' + CONFIG.JAM_REMINDER + ':00 — reminder belum isi pulang\n' +
    '• Setiap jam — lock baris 30 menit setelah pulang\n' +
    '• onEdit — guard proteksi per email'
  );
}

// ── setupAwal — Inisialisasi satu kali saat pertama setup ─────────────
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
      '2. Minta staf buka file dan test menu "📋 Absensi Saya"\n' +
      '3. Test stamp masuk/pulang langsung di sheet'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    Logger.log('Error setupAwal: ' + e.message);
  }
}
