# Istirahat 3 di Template & Settings — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Sisipkan kolom Ist 3 Mulai/Selesai di posisi K/L sheet absensi, shift 11 kolom +2, refactor formula Jam Efektif ke pola flat dengan ist3 subtraction, plus migration function untuk sheet existing.

**Architecture:** Insert kolom di posisi logis (setelah Ist 2, sebelum Pulang). Semua `COL_*` constant dari `COL_PULANG` ke kanan geser +2. Formula L (Jam Efektif) di-refactor ke pola flat (sudah dipakai OT 1/2). Migration function idempotent untuk sheet existing.

**Tech Stack:** Google Apps Script. No local test runner — verifikasi via `node --check` syntax + manual Apps Script + Logger output.

**Spec reference:** [`docs/superpowers/specs/2026-05-14-ist3-template-settings-design.md`](../specs/2026-05-14-ist3-template-settings-design.md)

---

## File Structure

- Modify: `Config.js` — 11 `COL_*` constant shift, 2 new constants, `TOTAL_COL=23`, `AUTO_ABSENSI` field, `_loadSettings` mapping, comment block.
- Modify: `Setup.js` — `buatSheetBulanBaru` (header array, widths, legends, protection), `buatSheetSettings` (2 row baru, format range), `setupValidasiBaris` (loop range). Plus fungsi baru `migrateSheetTambahIst3`.
- Modify: `Append.js` — `newRows` array, `setNumberFormat` array, color band literals → constants, `_pasangFormulaBaris` (formula L refactor flat + tambah ist3 di L/N/O).
- Modify: `Triggers.js` — tambah menu item Migrasi di `onOpen` admin section.

`Lock.js`, `Stamp.js`, `Rekap.js`, `Utils.js` tidak disentuh (semua pakai constant auto-update).

---

## Pre-flight

- [ ] **Step 0a: Verifikasi state working tree**

Run: `git status -- Config.js Setup.js Append.js Triggers.js`
Expected: clean atau hanya perubahan yang related ke task ini. Catat kalau ada perubahan unrelated dan jangan ikut commit.

- [ ] **Step 0b: Verifikasi syntax base files**

Run: `node --check Config.js && node --check Setup.js && node --check Append.js && node --check Triggers.js`
Expected: semua print no output (success).

---

## Task 1: Update `Config.js` — constants, AUTO_ABSENSI, _loadSettings

**Files:**
- Modify: `Config.js:36-65` (AUTO_ABSENSI struktur)
- Modify: `Config.js:105-117` (_loadSettings fallback + autoFieldMap)
- Modify: `Config.js:158-207` (skema comment + constants)

- [ ] **Step 1.1: Tambah field ist3 di AUTO_ABSENSI**

Di [`Config.js:47-77`](../../../Config.js#L47-L77), update kedua object divisi (`DEVELOPMENT` dan `WORKER`) untuk tambah `ist3Mulai` dan `ist3Selesai` antara `ist2Selesai` dan `pulang`:

Cari:
```js
    'DEVELOPMENT': {
      status      : 'Hadir',
      masuk       : '07:00',
      ist1Mulai   : '12:00',
      ist1Selesai : '13:00',
      ist2Mulai   : '',
      ist2Selesai : '',
      pulang      : '',
    },
```

Ganti dengan:
```js
    'DEVELOPMENT': {
      status      : 'Hadir',
      masuk       : '07:00',
      ist1Mulai   : '12:00',
      ist1Selesai : '13:00',
      ist2Mulai   : '',
      ist2Selesai : '',
      ist3Mulai   : '',
      ist3Selesai : '',
      pulang      : '',
    },
```

Cari blok `'WORKER':` (line 57-65) dan lakukan perubahan identik (tambah `ist3Mulai: '', ist3Selesai: '',` sebelum `pulang`).

- [ ] **Step 1.2: Update fallback object di `_loadSettings`**

Di [`Config.js:105-108`](../../../Config.js#L105-L108), cari:
```js
    if (!CONFIG.AUTO_ABSENSI[divisi])  CONFIG.AUTO_ABSENSI[divisi] = {
      status: 'Hadir', masuk: '', ist1Mulai: '', ist1Selesai: '',
      ist2Mulai: '', ist2Selesai: '', pulang: ''
    };
```

Ganti dengan:
```js
    if (!CONFIG.AUTO_ABSENSI[divisi])  CONFIG.AUTO_ABSENSI[divisi] = {
      status: 'Hadir', masuk: '', ist1Mulai: '', ist1Selesai: '',
      ist2Mulai: '', ist2Selesai: '',
      ist3Mulai: '', ist3Selesai: '',
      pulang: ''
    };
```

- [ ] **Step 1.3: Update `autoFieldMap` di `_loadSettings`**

Di [`Config.js:110-117`](../../../Config.js#L110-L117), cari:
```js
    const autoFieldMap = {
      MASUK        : 'masuk',
      IST1_MULAI   : 'ist1Mulai',
      IST1_SELESAI : 'ist1Selesai',
      IST2_MULAI   : 'ist2Mulai',
      IST2_SELESAI : 'ist2Selesai',
      PULANG       : 'pulang',
    };
```

Ganti dengan:
```js
    const autoFieldMap = {
      MASUK        : 'masuk',
      IST1_MULAI   : 'ist1Mulai',
      IST1_SELESAI : 'ist1Selesai',
      IST2_MULAI   : 'ist2Mulai',
      IST2_SELESAI : 'ist2Selesai',
      IST3_MULAI   : 'ist3Mulai',
      IST3_SELESAI : 'ist3Selesai',
      PULANG       : 'pulang',
    };
```

- [ ] **Step 1.4: Update mapping kolom comment + constants**

Di [`Config.js:170-189`](../../../Config.js#L170-L189) (comment "Mapping kolom sheet") dan [`Config.js:191-207`](../../../Config.js#L191-L207) (`const COL_*`), ganti seluruh blok dari line 170 sampai akhir file dengan:

```js
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
// K=11 Ist. Ketiga Mulai    → editable staf (HH:mm, opsional)
// L=12 Ist. Ketiga Selesai  → editable staf (HH:mm, opsional)
// M=13 Pulang            → editable staf (HH:mm)
// N=14 Jam Efektif       → formula otomatis (terkunci)
// O=15 Regular Hours     → formula otomatis (terkunci)
// P=16 OT 1              → formula otomatis (terkunci)
// Q=17 OT 2              → formula otomatis (terkunci)
// R=18 NOTE              → editable admin only (dropdown)
// S=19 SUNDAY/RED DAY    → editable admin only (dropdown)
// T=20 KETERANGAN        → editable staf (teks bebas)
// U=21 PLAN              → editable staf (dropdown via Web App)
// V=22 CATATAN TELAT     → editable staf (alasan telat masuk)
// W=23 CATATAN PULANG AWAL → editable staf (alasan pulang lebih awal)

const TOTAL_COL       = 23; // Jumlah kolom total (A sampai W)

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
const COL_IST3_M      = 11; // K
const COL_IST3_S      = 12; // L
const COL_PULANG      = 13; // M
const COL_EFEKTIF     = 14; // N — formula, terkunci
const COL_REGULAR_JAM = 15; // O — formula, terkunci
const COL_OT1         = 16; // P — formula, terkunci
const COL_OT2         = 17; // Q — formula, terkunci
const COL_NOTE        = 18; // R — admin only
const COL_SUNDAY      = 19; // S — admin only
const COL_KETERANGAN  = 20; // T — editable staf
const COL_PLAN        = 21; // U — editable staf
const COL_TELAT       = 22; // V — editable staf (alasan telat masuk)
const COL_PULANG_AWAL = 23; // W — editable staf (alasan pulang lebih awal)

// Batas kolom yang boleh diedit staf (E–W)
// Kolom N (COL_EFEKTIF) dikecualikan lewat guard di onEdit
// Kolom R, S (COL_NOTE, COL_SUNDAY) dikecualikan sebagai admin-only
const COL_EDIT_START = COL_STATUS;       // E = 5
const COL_EDIT_END   = COL_PULANG_AWAL;  // W = 23
```

- [ ] **Step 1.5: Verifikasi syntax Config.js**

Run: `node --check /Users/webadmin/Documents/Automations/absent-worker/Config.js`
Expected: no output.

- [ ] **Step 1.6: Commit Task 1**

```bash
cd /Users/webadmin/Documents/Automations/absent-worker
git add Config.js
git commit -m "$(cat <<'EOF'
feat(config): tambah Ist 3 ke skema kolom + AUTO_ABSENSI + _loadSettings

- TOTAL_COL 21→23, sisipkan COL_IST3_M=11 dan COL_IST3_S=12
- Geser 11 COL_* constant (COL_PULANG=11→13, COL_EFEKTIF=12→14, dst.)
- AUTO_ABSENSI per divisi: tambah ist3Mulai/ist3Selesai (default kosong)
- _loadSettings autoFieldMap: tambah IST3_MULAI/IST3_SELESAI
- Update comment "Mapping kolom sheet" dengan layout 23 kolom

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Update `Setup.js` — rendering sheet & settings

**Files:**
- Modify: `Setup.js:44-49` (legenda)
- Modify: `Setup.js:61-83` (header array)
- Modify: `Setup.js:97` (column widths)
- Modify: `Setup.js:102` (header protection range)
- Modify: `Setup.js:159-175` (buatSheetSettings rows)
- Modify: `Setup.js:182` (setNumberFormat range)
- Modify: `Setup.js:337` (setupValidasiBaris loop)

- [ ] **Step 2.1: Update legenda baris 2**

Di [`Setup.js:44-49`](../../../Setup.js#L44-L49), cari:
```js
    const legends = [
      [1,  4, 'ABU = sudah lewat',    '#F1EFE8', '#5F5E5A'],
      [5,  4, 'PUTIH = bisa diedit',  '#FFFFFF',  '#2C2C2A'],
      [9,  4, 'UNGU = formula auto',  '#EEEDFE',  '#534AB7'],
      [13, 9, 'KUNING = hari ini',    '#FFF9C4',  '#633806'],
    ];
```

Ganti dengan:
```js
    const legends = [
      [1,  4, 'ABU = sudah lewat',    '#F1EFE8', '#5F5E5A'],
      [5,  9, 'PUTIH = bisa diedit',  '#FFFFFF', '#2C2C2A'],   // E:M (Status sampai Pulang)
      [14, 4, 'UNGU = formula auto',  '#EEEDFE', '#534AB7'],   // N:Q (Jam Efektif sampai OT 2)
      [18, 6, 'KUNING = hari ini',    '#FFF9C4', '#633806'],   // R:W (NOTE sampai CATATAN PULANG AWAL)
    ];
```

- [ ] **Step 2.2: Update header array — sisipkan Ist 3**

Di [`Setup.js:61-83`](../../../Setup.js#L61-L83), cari:
```js
      ['Ist. Kedua\nMulai',                 '#E1F5EE', '#085041'],
      ['Ist. Kedua\nSelesai',               '#E1F5EE', '#085041'],
      ['Pulang',                            '#E1F5EE', '#085041'],
```

Ganti dengan (sisipkan 2 entry baru):
```js
      ['Ist. Kedua\nMulai',                 '#E1F5EE', '#085041'],
      ['Ist. Kedua\nSelesai',               '#E1F5EE', '#085041'],
      ['Ist. Ketiga\nMulai',                '#E1F5EE', '#085041'],
      ['Ist. Ketiga\nSelesai',              '#E1F5EE', '#085041'],
      ['Pulang',                            '#E1F5EE', '#085041'],
```

- [ ] **Step 2.3: Update column widths**

Di [`Setup.js:97`](../../../Setup.js#L97), cari:
```js
    const colWidths = [90,80,130,180,70,70,90,90,90,90,70,100,60,60,60,120,140,160,80,100,100];
```

Ganti dengan (sisipkan `90,90` di index 10, antara ist2 dan pulang):
```js
    const colWidths = [90,80,130,180,70,70,90,90,90,90, 90,90, 70,100,60,60,60,120,140,160,80,100,100];
```

- [ ] **Step 2.4: Update header protection range**

Di [`Setup.js:102`](../../../Setup.js#L102), cari:
```js
    const headerProt = sheet.getRange('A1:S3').protect();
```

Ganti dengan:
```js
    const headerProt = sheet.getRange('A1:W3').protect();
```

- [ ] **Step 2.5: Update `buatSheetSettings` rows**

Di [`Setup.js:159-175`](../../../Setup.js#L159-L175), cari:
```js
  const rows = [
    // Jam auto-absensi
    ['MASUK',              auto.masuk        || '', 'Jam masuk default (HH:mm) — kosongkan jika staf isi sendiri'],
    ['IST1_MULAI',         auto.ist1Mulai    || '', 'Istirahat pertama mulai (HH:mm) — kosongkan jika tidak ada'],
    ['IST1_SELESAI',       auto.ist1Selesai  || '', 'Istirahat pertama selesai (HH:mm)'],
    ['IST2_MULAI',         auto.ist2Mulai    || '', 'Istirahat kedua mulai (HH:mm) — kosongkan jika tidak ada'],
    ['IST2_SELESAI',       auto.ist2Selesai  || '', 'Istirahat kedua selesai (HH:mm)'],
    ['PULANG',             auto.pulang       || '', 'Jam pulang default (HH:mm) — kosongkan jika staf isi sendiri'],
```

Ganti dengan (sisipkan 2 row IST3 sebelum PULANG):
```js
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
```

- [ ] **Step 2.6: Update `setNumberFormat` range untuk row jam**

Di [`Setup.js:182`](../../../Setup.js#L182) (di dalam `buatSheetSettings`), cari:
```js
  // Baris 2–6 (MASUK s/d IST2_SELESAI) — kolom VALUE format HH:mm
  // agar Sheets menyimpan sebagai time serial, bukan string, sehingga
  // formula jam kerja di sheet absensi bisa menghitung dengan benar
  sheet.getRange(2, 2, 5, 1).setNumberFormat('HH:mm');
```

Ganti dengan (5 row jam jadi 7 row — MASUK, IST1_M/S, IST2_M/S, IST3_M/S; PULANG terpisah di bawah tapi pasang format-nya juga):

Wait, perhitungan: dulu 5 row = MASUK, IST1_M, IST1_S, IST2_M, IST2_S (PULANG terpisah). Setelah tambah 2 (IST3_M, IST3_S), jadi 7 row. PULANG masih row ke-8 yang TIDAK kena format ini di kode lama. Mari kita perluas mencakup PULANG juga supaya konsisten: 8 row.

Ganti dengan:
```js
  // Baris 2–9 (MASUK s/d PULANG) — kolom VALUE format HH:mm
  // agar Sheets menyimpan sebagai time serial, bukan string, sehingga
  // formula jam kerja di sheet absensi bisa menghitung dengan benar
  sheet.getRange(2, 2, 8, 1).setNumberFormat('HH:mm');
```

- [ ] **Step 2.7: Update `setupValidasiBaris` loop**

Di [`Setup.js:337`](../../../Setup.js#L337), cari:
```js
  // G–K: opsional, jika diisi harus HH:MM
  ['G','H','I','J','K'].forEach((col, idx) => {
    sheet.getRange(startRow, COL_IST1_M + idx, numRows, 1).setDataValidation(
```

Ganti dengan:
```js
  // G–M: opsional jam istirahat 1/2/3 dan pulang, jika diisi harus HH:MM
  ['G','H','I','J','K','L','M'].forEach((col, idx) => {
    sheet.getRange(startRow, COL_IST1_M + idx, numRows, 1).setDataValidation(
```

- [ ] **Step 2.8: Verifikasi syntax Setup.js**

Run: `node --check /Users/webadmin/Documents/Automations/absent-worker/Setup.js`
Expected: no output.

- [ ] **Step 2.9: Commit Task 2**

```bash
cd /Users/webadmin/Documents/Automations/absent-worker
git add Setup.js
git commit -m "$(cat <<'EOF'
feat(setup): render sheet & _Settings dengan Ist 3

- Header array: sisipkan 'Ist. Ketiga Mulai/Selesai' di posisi 10-11
- Column widths: tambah 90/90 untuk K dan L (Ist 3)
- Legenda baris 2: span dan posisi disesuaikan (E:M editable, N:Q formula, R:W kuning)
- Header protection: A1:S3 → A1:W3
- buatSheetSettings: tambah 2 row IST3_MULAI/IST3_SELESAI sebelum PULANG
- setNumberFormat HH:mm: range 5 row → 8 row (cover semua jam dari MASUK ke PULANG)
- setupValidasiBaris: loop G:M (was G:K) untuk validasi jam ist1/2/3 + pulang

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Update `Append.js` — row insertion, formulas, color band

**Files:**
- Modify: `Append.js:99-116` (newRows array)
- Modify: `Append.js:124-129` (setNumberFormat)
- Modify: `Append.js:131-143` (color bands)
- Modify: `Append.js:264-315` (_pasangFormulaBaris — variabel + 4 formula)

- [ ] **Step 3.1: Update `newRows` array — sisipkan ist3 cells**

Di [`Append.js:99-116`](../../../Append.js#L99-L116), cari:
```js
    const newRows = staf.map(s => [
      today,                                        // A: Tanggal
      namaHari,                                     // B: Hari
      s.nama,                                       // C: Nama
      s.email,                                      // D: Email
      auto ? (auto.status      || '') : '',          // E: Status
      auto ? toSerial(auto.masuk)       : '',        // F: Masuk
      auto ? toSerial(auto.ist1Mulai)   : '',        // G: Ist. 1 Mulai
      auto ? toSerial(auto.ist1Selesai) : '',        // H: Ist. 1 Selesai
      auto ? toSerial(auto.ist2Mulai)   : '',        // I: Ist. 2 Mulai
      auto ? toSerial(auto.ist2Selesai) : '',        // J: Ist. 2 Selesai
      auto ? toSerial(auto.pulang)      : '',        // K: Pulang
      '', '', '', '',                               // L–O: formula (diset di bawah)
      '', '',                                       // P–Q: admin only
      '', '',                                       // R: Keterangan, S: Plan
      '', '',                                       // T: Catatan Telat, U: Catatan Pulang Awal
    ]);
```

Ganti dengan:
```js
    const newRows = staf.map(s => [
      today,                                        // A: Tanggal
      namaHari,                                     // B: Hari
      s.nama,                                       // C: Nama
      s.email,                                      // D: Email
      auto ? (auto.status      || '') : '',          // E: Status
      auto ? toSerial(auto.masuk)       : '',        // F: Masuk
      auto ? toSerial(auto.ist1Mulai)   : '',        // G: Ist. 1 Mulai
      auto ? toSerial(auto.ist1Selesai) : '',        // H: Ist. 1 Selesai
      auto ? toSerial(auto.ist2Mulai)   : '',        // I: Ist. 2 Mulai
      auto ? toSerial(auto.ist2Selesai) : '',        // J: Ist. 2 Selesai
      auto ? toSerial(auto.ist3Mulai)   : '',        // K: Ist. 3 Mulai
      auto ? toSerial(auto.ist3Selesai) : '',        // L: Ist. 3 Selesai
      auto ? toSerial(auto.pulang)      : '',        // M: Pulang
      '', '', '', '',                               // N–Q: formula (diset di bawah)
      '', '',                                       // R–S: admin only (NOTE, SUNDAY/RED DAY)
      '', '',                                       // T: Keterangan, U: Plan
      '', '',                                       // V: Catatan Telat, W: Catatan Pulang Awal
    ]);
```

- [ ] **Step 3.2: Update `setNumberFormat` HH:mm array**

Di [`Append.js:126-128`](../../../Append.js#L126-L128), cari:
```js
      [COL_MASUK, COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_PULANG]
        .forEach(col => sheet.getRange(insertAt, col, newRows.length, 1)
          .setNumberFormat('HH:mm'));
```

Ganti dengan:
```js
      [COL_MASUK, COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_IST3_M, COL_IST3_S, COL_PULANG]
        .forEach(col => sheet.getRange(insertAt, col, newRows.length, 1)
          .setNumberFormat('HH:mm'));
```

- [ ] **Step 3.3: Update color bands — pakai constant supaya tahan shift**

Di [`Append.js:131-143`](../../../Append.js#L131-L143), cari:
```js
    // Warna kolom
    sheet.getRange(insertAt, 1,  newRows.length, 4)
      .setBackground('#FFF9C4').setFontColor('#5F5E5A');  // A:D terkunci
    sheet.getRange(insertAt, 5,  newRows.length, 7)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // E:K editable
    sheet.getRange(insertAt, 12, newRows.length, 4)
      .setBackground('#FFF9C4').setFontColor('#534AB7').setFontWeight('bold'); // L:O formula
    sheet.getRange(insertAt, 16, newRows.length, 2)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // P:Q admin
    sheet.getRange(insertAt, 18, newRows.length, 2)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // R:S keterangan/plan
    sheet.getRange(insertAt, 20, newRows.length, 2)
      .setBackground('#FFF9C4').setFontColor('#E65100');  // T:U catatan telat/pulang awal
```

Ganti dengan (pakai constants supaya tidak break di shift mendatang):
```js
    // Warna kolom — pakai COL_* constant supaya tahan terhadap shift kolom
    sheet.getRange(insertAt, COL_TANGGAL, newRows.length, COL_EMAIL - COL_TANGGAL + 1)
      .setBackground('#FFF9C4').setFontColor('#5F5E5A');  // A:D terkunci
    sheet.getRange(insertAt, COL_STATUS, newRows.length, COL_PULANG - COL_STATUS + 1)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // E:M editable
    sheet.getRange(insertAt, COL_EFEKTIF, newRows.length, COL_OT2 - COL_EFEKTIF + 1)
      .setBackground('#FFF9C4').setFontColor('#534AB7').setFontWeight('bold'); // N:Q formula
    sheet.getRange(insertAt, COL_NOTE, newRows.length, COL_SUNDAY - COL_NOTE + 1)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // R:S admin
    sheet.getRange(insertAt, COL_KETERANGAN, newRows.length, COL_PLAN - COL_KETERANGAN + 1)
      .setBackground('#FFF9C4').setFontColor('#2C2C2A');  // T:U keterangan/plan
    sheet.getRange(insertAt, COL_TELAT, newRows.length, COL_PULANG_AWAL - COL_TELAT + 1)
      .setBackground('#FFF9C4').setFontColor('#E65100');  // V:W catatan telat/pulang awal
```

- [ ] **Step 3.4: Refactor `_pasangFormulaBaris` — variabel + 4 formula**

Di [`Append.js:264-315`](../../../Append.js#L264-L315), ganti seluruh fungsi `_pasangFormulaBaris` dengan versi baru:

```js
// ── _pasangFormulaBaris — Private: set formula N, O, P, Q ─────────────
// Dipanggil oleh appendHariIni() dan generateFullMonth()
//
// Formula L (sekarang N — Jam Efektif) di-refactor ke pola flat:
// pulang - masuk dikurangi durasi setiap istirahat yang diisi (3 istirahat).
// Pola lama nested-IF dihilangkan karena 2³=8 cabang sulit dirawat.
function _pasangFormulaBaris(sheet, startRow, numRows) {
  for (let r = startRow; r < startRow + numRows; r++) {
    const a=`A${r}`, b=`B${r}`, e=`E${r}`, f=`F${r}`,
          g=`G${r}`, h=`H${r}`, i=`I${r}`, j=`J${r}`,
          k=`K${r}`, l=`L${r}`, m=`M${r}`, n=`N${r}`;

    // N: Jam Efektif (fraction hari) — pola flat
    // = pulang - masuk - ist1_dur - ist2_dur - ist3_dur (kalau masing-masing diisi)
    sheet.getRange(r, COL_EFEKTIF).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${m}=""),0,` +
      `${m}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0)))`
    );

    // O: Regular Hours — Red Day langsung dapat 7 jam (hari libur dibayar penuh)
    sheet.getRange(r, COL_REGULAR_JAM).setFormula(
      `=IF(${e}="Red Day",${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
      `IF(${b}="Saturday",` +
        `IF(${n}>=${CONFIG.DAYS_HOUR.SATURDAY}/24,` +
          `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${n}),` +
      `IF(${n}>=${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
        `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${n})))`
    );

    // P: OT 1 (maks 1 jam di atas regular) — flat, subtract 3 istirahat
    sheet.getRange(r, COL_OT1).setFormula(
      `=IF(${e}<>"Hadir",0,IF(OR(${f}="",${m}=""),0,` +
      `IF((${m}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0))` +
        `<=IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24,0,` +
      `MIN(1/24,(${m}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0))` +
        `-IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24))))`
    );

    // Q: OT 2 (di atas OT 1) — flat, subtract 3 istirahat
    sheet.getRange(r, COL_OT2).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${m}=""),0,` +
      `IF((${m}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0))` +
        `<=(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24,` +
      `0,` +
      `${m}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0)` +
        `-(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24)))`
    );
  }
}
```

Catatan penting:
- Local var `k` dan `l` SEKARANG menunjuk ke kolom K (ist3 Mulai) dan L (ist3 Selesai), BUKAN ke pulang/jam-efektif seperti dulu.
- Local var `m` (baru) = pulang. `n` (baru) = jam efektif self-reference di formula M (sekarang O).
- 4 formula sekarang konsisten dengan pola flat: pulang-masuk dikurangi 3 durasi istirahat conditional.

- [ ] **Step 3.5: Verifikasi syntax Append.js**

Run: `node --check /Users/webadmin/Documents/Automations/absent-worker/Append.js`
Expected: no output.

- [ ] **Step 3.6: Commit Task 3**

```bash
cd /Users/webadmin/Documents/Automations/absent-worker
git add Append.js
git commit -m "$(cat <<'EOF'
feat(append): tulis kolom Ist 3 + refactor formula ke pola flat

- newRows: sisipkan 2 cell toSerial(ist3Mulai/Selesai) di posisi 11-12
- setNumberFormat HH:mm: tambah COL_IST3_M, COL_IST3_S ke array
- Color band: ganti literal posisi (5,7), (12,4), (16,2), dst. ke
  COL_* constant supaya tahan shift kolom di masa depan
- _pasangFormulaBaris: refactor 4 formula (N=Jam Efektif, O=Regular,
  P=OT1, Q=OT2) ke pola flat dengan subtract 3 durasi istirahat
  conditional. Pola nested-IF lama dibuang — 2³=8 cabang sulit dirawat.

Hasil matematis identik untuk kasus tanpa ist3 (term ist3 = 0 jika
salah satu sel kosong). Ist3 hanya berpengaruh kalau kedua sel diisi.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Migration function + menu entry

**Files:**
- Modify: `Setup.js` — tambah fungsi baru `migrateSheetTambahIst3()` di akhir file
- Modify: `Triggers.js:168-185` (menu admin)

- [ ] **Step 4.1: Tambah fungsi `migrateSheetTambahIst3()` di Setup.js**

Tambahkan fungsi berikut di akhir file `Setup.js` (setelah `proteksiBarisBaru` selesai):

```js

// ── migrateSheetTambahIst3 — Migrasi sheet existing untuk skema 23 kolom ─
// Idempotent: cek dulu apakah sheet sudah punya kolom Ist 3, kalau sudah → skip.
// Untuk sheet yang masih 21 kolom:
//   1. Insert 2 kolom kosong di posisi 11 (Google Sheets auto-shift formula existing)
//   2. Set header sel K3 dan L3 dengan format Ist 3
//   3. Set column width K dan L
//   4. Re-merge baris 1 (title) dan baris 2 (legenda) ke skema lebar baru
//   5. Re-pasang formula untuk baris hari ini (upgrade ke pola flat baru)
//   6. Re-run proteksiBarisBaru() untuk baris hari ini
//
// HANYA migrate sheet yang nama-nya match pola divisi (DIVISI atau DIVISI_MMM_YYYY).
function migrateSheetTambahIst3() {
  _requireAdmin();
  _loadSettings();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const today  = getToday();
  const hasil  = [];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const divisi = CONFIG.DIVISI.find(d => name === d || name.startsWith(d + '_'));
    if (!divisi) return; // bukan sheet divisi, skip

    // Idempotency: cek header K3 — kalau sudah "Ist. Ketiga", skip
    const headerK = String(sheet.getRange(3, 11).getValue() || '').toLowerCase();
    if (headerK.includes('ist') && headerK.includes('ketiga')) {
      hasil.push('⏭ ' + name + ': sudah punya Ist 3, skip');
      Logger.log('⏭ ' + name + ': sudah migrated');
      return;
    }

    // Step 1: Insert 2 kolom di posisi 11. Sheets auto-shift formula existing.
    sheet.insertColumnsBefore(11, 2);

    // Step 2: Set header K3 dan L3
    [
      [11, 'Ist. Ketiga\nMulai'],
      [12, 'Ist. Ketiga\nSelesai'],
    ].forEach(([col, text]) => {
      sheet.getRange(3, col)
        .setValue(text)
        .setBackground('#E1F5EE').setFontColor('#085041')
        .setFontWeight('bold').setFontSize(9)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, false, false,
          '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    });

    // Step 3: Column widths
    sheet.setColumnWidth(11, 90);
    sheet.setColumnWidth(12, 90);

    // Step 4a: Re-merge baris 1 (title)
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
    const title = String(sheet.getRange(1, 1).getValue() || '');
    sheet.getRange(1, 1, 1, TOTAL_COL).merge()
      .setValue(title)
      .setBackground('#0F6E56').setFontColor('#FFFFFF')
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);

    // Step 4b: Re-merge baris 2 (legenda)
    sheet.getRange(2, 1, 1, sheet.getMaxColumns()).breakApart();
    const legends = [
      [1,  4, 'ABU = sudah lewat',    '#F1EFE8', '#5F5E5A'],
      [5,  9, 'PUTIH = bisa diedit',  '#FFFFFF', '#2C2C2A'],
      [14, 4, 'UNGU = formula auto',  '#EEEDFE', '#534AB7'],
      [18, 6, 'KUNING = hari ini',    '#FFF9C4', '#633806'],
    ];
    for (const [startCol, span, text, bg, fg] of legends) {
      sheet.getRange(2, startCol, 1, span).merge()
        .setValue(text).setBackground(bg).setFontColor(fg)
        .setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, false, false,
          '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(2, 16);

    // Step 5+6: Cari baris hari ini, re-pasang formula + proteksi
    const lastRow = sheet.getLastRow();
    let firstToday = -1, lastToday = -1;
    if (lastRow >= 4) {
      const dates = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        if (dates[i][0] instanceof Date && isSameDate(dates[i][0], today)) {
          if (firstToday === -1) firstToday = i + 4;
          lastToday = i + 4;
        }
      }
    }

    if (firstToday !== -1) {
      const numToday = lastToday - firstToday + 1;
      _pasangFormulaBaris(sheet, firstToday, numToday);
      proteksiBarisBaru(sheet, divisi, firstToday, numToday);
      Logger.log('✓ ' + name + ': re-pasang formula + proteksi baris ' +
        firstToday + '–' + lastToday);
    } else {
      Logger.log('⚠ ' + name + ': tidak ada baris hari ini — skip re-formula');
    }

    hasil.push('✓ ' + name + ': migrated (kolom Ist 3 ditambah)');
  });

  if (hasil.length === 0) {
    hasil.push('⚠ Tidak ada sheet divisi yang ditemukan');
  }

  const msg = '🔄 Migrasi Ist 3 selesai!\n\n' + hasil.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}
```

- [ ] **Step 4.2: Verifikasi syntax Setup.js**

Run: `node --check /Users/webadmin/Documents/Automations/absent-worker/Setup.js`
Expected: no output.

- [ ] **Step 4.3: Tambah menu item di Triggers.js**

Di [`Triggers.js:168-185`](../../../Triggers.js#L168-L185) (menu admin), cari:
```js
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
```

Ganti dengan (tambah separator + item migrasi di paling bawah):
```js
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
    .addSeparator()
    .addItem('🔄 Migrasi Sheet Tambah Ist 3',  'migrateSheetTambahIst3')
    .addToUi();
```

- [ ] **Step 4.4: Verifikasi syntax Triggers.js**

Run: `node --check /Users/webadmin/Documents/Automations/absent-worker/Triggers.js`
Expected: no output.

- [ ] **Step 4.5: Commit Task 4**

```bash
cd /Users/webadmin/Documents/Automations/absent-worker
git add Setup.js Triggers.js
git commit -m "$(cat <<'EOF'
feat(migrate): fungsi migrateSheetTambahIst3 + menu admin

Idempotent migration helper untuk sheet existing yang masih 21 kolom:
- insertColumnsBefore(11, 2) — Google Sheets auto-shift formula existing
- Set header K3/L3 dengan format Ist 3 (hijau toska, wrap, dst.)
- Set column width K/L = 90
- Re-merge baris 1 (title) dan baris 2 (legenda) ke lebar 23 kolom
- Re-pasang formula baris hari ini (upgrade ke pola flat baru)
- Re-run proteksiBarisBaru() supaya range E:M (was E:K) benar
- Idempotency check via header K3 — kalau sudah Ist Ketiga, skip

Menu admin Triggers.js dapat item '🔄 Migrasi Sheet Tambah Ist 3'
di seksi setup terbawah. Manual trigger; tidak dipanggil otomatis.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: Manual verification & migration (handoff)

This task is performed by HRD/developer after deploy, not the implementing agent.

- [ ] **Step 5.1: Deploy code baru**

Run: `./deploy.sh`
Expected: `clasp push` sukses ke semua TARGETS.

- [ ] **Step 5.2: Migrasi TESTING WORKER 2 (uji pertama)**

Buka spreadsheet TESTING WORKER 2:
1. Menu 🔧 Admin → 🔄 Migrasi Sheet Tambah Ist 3
2. Expected alert: `✓ TESTING WORKER 2_May_2026: migrated (kolom Ist 3 ditambah)`
3. Verify visual: kolom K3 = "Ist. Ketiga Mulai", L3 = "Ist. Ketiga Selesai", M3 = "Pulang", N3 = "Jam Efektif"
4. Baris hari ini (4-5): kolom K/L kosong, kolom M masih ada nilai pulang lama, formula di N updated.

- [ ] **Step 5.3: Update _Settings**

Di TESTING WORKER 2:
1. Menu 🔧 Admin → ⚙️ Buat/Reset Sheet Settings (konfirmasi reset)
2. Sheet _Settings baru muncul dengan row IST3_MULAI dan IST3_SELESAI
3. Isi IST3_MULAI = `14:30`, IST3_SELESAI = `14:45` (atau bebas)

- [ ] **Step 5.4: Test appendHariIni dengan ist3 auto-fill**

1. Hapus baris hari ini di sheet absensi (kalau ada)
2. Apps Script Editor → run `appendHariIni()` manual
3. Verify: kolom K terisi `14:30`, kolom L terisi `14:45`, kolom M terisi nilai pulang dari _Settings
4. Verify formula N (Jam Efektif): kalau worker isi masuk + pulang, hasil sudah subtract durasi ist3 (15 menit)

- [ ] **Step 5.5: Migrate spreadsheet production lainnya**

Untuk setiap spreadsheet di TARGETS deploy.sh (kecuali TESTING WORKER 2 yang sudah diuji):
1. Buka spreadsheet
2. Menu 🔧 Admin → 🔄 Migrasi Sheet Tambah Ist 3
3. Verify log "✓ migrated"

Catatan: kalau ingin atur default ist3 per divisi (mis. WORKER butuh ist3 tapi DEVELOPMENT tidak), edit `_Settings` masing-masing spreadsheet — kosongkan IST3_MULAI/SELESAI di divisi yang tidak butuh.

---

## Self-Review Notes (addressed inline)

- **Spec coverage**:
  - Skema kolom 23 — Task 1 (constants), Task 2 (rendering), Task 3 (row data).
  - Formula flat refactor + ist3 — Task 3 Step 3.4.
  - _Settings IST3 rows + AUTO_ABSENSI — Task 1 + Task 2.
  - Migration function idempotent — Task 4 Step 4.1.
  - Menu entry — Task 4 Step 4.3.
  - Manual verification — Task 5.
- **No placeholders**: Semua code block konkret. Tidak ada "TODO".
- **Type consistency**: 
  - Property name `ist3Mulai`/`ist3Selesai` konsisten di Config.js, _Settings autoFieldMap, Append.js newRows.
  - Setting key `IST3_MULAI`/`IST3_SELESAI` konsisten di buatSheetSettings rows dan autoFieldMap.
  - Function name `migrateSheetTambahIst3` konsisten di definition (Setup.js) dan menu entry (Triggers.js).
- **Risk noted**: color band literal di Append.js sengaja di-refactor ke pakai constant supaya tahan shift mendatang — bukan scope creep, ini menutup risiko regresi di Task 3 sendiri.
