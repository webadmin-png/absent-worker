// ═══════════════════════════════════════════════════════════════════════
// CONFIG.JS — Konfigurasi global dan mapping kolom sheet
// Ubah nilai di sini untuk menyesuaikan dengan kebutuhan instansi.
// ═══════════════════════════════════════════════════════════════════════

var CONFIG = {
  SHEET_MASTER  : 'Master_Data',
  DIVISI        : ['DEVELOPMENT'],
  JAM_REMINDER  : 17,
  NAMA_INSTANSI : 'PT InFashion',
  TIMEZONE      : 'Asia/Makassar', // WITA (UTC+8)

  // Email yang boleh edit semua baris bebas (HRD, admin, supervisor)
  ADMIN_EMAILS  : [
    'webadmin@wooden-ships.com',
  ],

  // ID Google Spreadsheet terpisah yang berisi sheet "Settings"
  // Diisi setelah menjalankan setupSettings() — jangan kosongkan setelah diisi
  SETTINGS_SPREADSHEET_ID : '',

  // Menit setelah jam pulang diisi → baris dikunci otomatis
  SELISIH_MENIT_LOCK : 30,

  // Opsi plan jam kerja yang tampil di Web App (bisa di-override lewat Settings spreadsheet)
  PLAN_JAM : [
    '07:00 - 16:00',
    '08:00 - 17:00',
    '09:00 - 18:00',
    '10:00 - 19:00',
    '11:00 - 20:00',
  ],

  DAYS_HOUR     : {
    REGULAR_DAYS : 7,   // Jam kerja normal per hari (senin–jumat)
    SATURDAY     : 5    // Jam kerja normal hari sabtu
  },

  // Divisi yang diabsenkan otomatis oleh leader saat append.
  // Jam masuk, istirahat, dst. langsung diisi — leader hanya perlu isi jam pulang.
  // Key  : nama divisi (harus cocok persis dengan TARGETS di deploy.sh, huruf besar)
  // Value: jam yang diisi otomatis (kosongkan string jika tidak ingin diisi)
  // Setiap spreadsheet hanya membaca entry yang cocok dengan DIVISI-nya —
  // aman menambah semua divisi di sini sekaligus.
  AUTO_ABSENSI  : {
    'DEVELOPMENT': {
      status      : 'Hadir',
      masuk       : '07:00',
      ist1Mulai   : '12:00',
      ist1Selesai : '13:00',
      ist2Mulai   : '',
      ist2Selesai : '',
      pulang      : '',
    },
    'WORKER': {
      status      : 'Hadir',
      masuk       : '08:00',
      ist1Mulai   : '12:00',
      ist1Selesai : '13:00',
      ist2Mulai   : '',
      ist2Selesai : '',
      pulang      : '',
    },
    // Tambah divisi baru di sini — nama harus sama dengan di TARGETS deploy.sh
    // 'FINANCE': {
    //   status      : 'Hadir',
    //   masuk       : '09:00',
    //   ist1Mulai   : '12:00',
    //   ist1Selesai : '13:00',
    //   ist2Mulai   : '',
    //   ist2Selesai : '',
    //   pulang      : '',
    // },
  }
};

// ── _loadSettings — Override CONFIG dari sheet _Settings lokal ────────
// Prioritas: sheet _Settings (dalam spreadsheet ini) > Config.js default.
// Dipanggil di awal fungsi yang butuh settings terkini tanpa redeploy.
function _loadSettings() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const localSheet = ss.getSheetByName('_Settings');
    if (!localSheet) return;

    const lastRow = localSheet.getLastRow();
    if (lastRow < 2) return;

    const data = localSheet.getRange(2, 1, lastRow - 1, 2).getValues();

    // Pre-pass: baca DIVISI lebih dulu agar AUTO_ABSENSI pakai key yang benar.
    // Ini memungkinkan duplicate spreadsheet tanpa deploy — cukup ubah DIVISI di _Settings.
    for (const [key, value] of data) {
      const k = String(key).trim();
      const v = String(value).trim().toUpperCase();
      if (k === 'DIVISI' && v) { CONFIG.DIVISI = [v]; break; }
    }

    const divisi = (CONFIG.DIVISI[0] || '').toUpperCase();

    // Pastikan entry AUTO_ABSENSI untuk divisi ini ada sebelum diisi
    if (!CONFIG.AUTO_ABSENSI)          CONFIG.AUTO_ABSENSI = {};
    if (!CONFIG.AUTO_ABSENSI[divisi])  CONFIG.AUTO_ABSENSI[divisi] = {
      status: 'Hadir', masuk: '', ist1Mulai: '', ist1Selesai: '',
      ist2Mulai: '', ist2Selesai: '', pulang: ''
    };

    const autoFieldMap = {
      MASUK        : 'masuk',
      IST1_MULAI   : 'ist1Mulai',
      IST1_SELESAI : 'ist1Selesai',
      IST2_MULAI   : 'ist2Mulai',
      IST2_SELESAI : 'ist2Selesai',
      PULANG       : 'pulang',
    };

    for (const [key, value] of data) {
      const k = String(key).trim();
      const v = String(value).trim();
      if (!k) continue;

      if (autoFieldMap[k] !== undefined) {
        // Gunakan cellToTimeStr agar Date object dari Sheets tidak jadi string "Sat Dec 30 1899..."
        CONFIG.AUTO_ABSENSI[divisi][autoFieldMap[k]] = cellToTimeStr(value);
        continue;
      }

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
// R=18 KETERANGAN        → editable staf (teks bebas)
// S=19 PLAN              → editable staf (dropdown via Web App)

const TOTAL_COL       = 21; // Jumlah kolom total (A sampai U)

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
const COL_KETERANGAN  = 18; // R — editable staf
const COL_PLAN        = 19; // S — editable staf
const COL_TELAT        = 20; // T — editable staf (alasan telat masuk)
const COL_PULANG_AWAL  = 21; // U — editable staf (alasan pulang lebih awal)

// Batas kolom yang boleh diedit staf (E–U)
// Kolom L (COL_EFEKTIF) dikecualikan lewat guard di onEdit
// Kolom P, Q (COL_NOTE, COL_SUNDAY) dikecualikan sebagai admin-only
const COL_EDIT_START = COL_STATUS;  // E = 5
const COL_EDIT_END   = COL_PULANG_AWAL;    // U = 21
