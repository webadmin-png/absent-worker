# _Settings Row Jam Editable Worker+Asisten — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ijinkan worker dan asisten (dari Master_Data) edit row 2–9 (jam settings) di sheet `_Settings`, sementara row 11–15 tetap admin only. Plus helper untuk refresh editor list saat Master_Data berubah.

**Architecture:** Tambah lapis kedua proteksi range `A2:C9` di `buatSheetSettings()` dengan editor = owner + admin + worker + asisten (dibaca dari Master_Data). Range protection lebih spesifik override sheet-level admin protection di area-nya. Helper `perbaruiAksesSettings()` baru untuk refresh editor list tanpa reset sheet, menu admin item baru. DRY: extract `_setProteksiSettingsJam(sheet)` dipakai oleh keduanya.

**Tech Stack:** Google Apps Script (clasp deploy). Verifikasi via `node --check` syntax + manual run di Apps Script editor + browser test sebagai user worker/asisten.

**Spec reference:** [`docs/superpowers/specs/2026-05-14-settings-worker-edit-jam-design.md`](../specs/2026-05-14-settings-worker-edit-jam-design.md)

---

## File Structure

- Modify: `Setup.js` — modifikasi `buatSheetSettings()`, tambah helper `_setProteksiSettingsJam(sheet)`, tambah fungsi `perbaruiAksesSettings()`.
- Modify: `Triggers.js` — tambah menu item `🔑 Perbarui Akses Settings`.

No new files. No file deletions.

---

## Pre-flight

- [ ] **Step 0a: Clean working tree**

Run: `git status`
Expected: working tree clean atau hanya berisi commit spec/plan ist3/asisten yang sudah committed.

---

## Task 1: Helper `_setProteksiSettingsJam` + integrasi `buatSheetSettings`

**Files:**
- Modify: `Setup.js:215-222` (blok sheet protection)
- Modify: `Setup.js` (tambah helper function sebelum `buatSheetSettings`)

- [ ] **Step 1.1: Tambah helper `_setProteksiSettingsJam` di Setup.js**

Sisipkan blok berikut di Setup.js **sebelum** fungsi `buatSheetSettings()` (sekitar [Setup.js:126](Setup.js#L126), di antara fungsi `buatSheetBulanBaru` dan `buatSheetSettings`):

```js
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
```

- [ ] **Step 1.2: Integrasi helper ke `buatSheetSettings()`**

Modifikasi blok proteksi di `buatSheetSettings()`. Replace block [Setup.js:212-229](Setup.js#L212-L229):

```js
  // ── Proteksi: KEY dan Keterangan terkunci, VALUE bisa diedit admin ──
  const owner = Session.getEffectiveUser();

  const sheetProt = sheet.protect();
  sheetProt.setDescription('_Settings — hanya admin yang bisa edit');
  sheetProt.setWarningOnly(false);
  sheetProt.removeEditors(sheetProt.getEditors());
  sheetProt.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { sheetProt.addEditor(adminEmail); } catch(e) {}
  }

  ui.alert(
    '✅ Sheet _Settings berhasil dibuat!\n\n' +
    'Edit kolom VALUE untuk mengubah setting.\n' +
    'Perubahan berlaku otomatis di append/trigger berikutnya.\n\n' +
    'Sheet ini hanya bisa diedit oleh admin.'
  );
}
```

dengan:

```js
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
```

- [ ] **Step 1.3: Verifikasi Setup.js syntax**

Run: `node --check Setup.js && echo OK`
Expected: `OK`.

- [ ] **Step 1.4: Commit Task 1**

```bash
git add Setup.js
git commit -m "$(cat <<'EOF'
feat(setup): _Settings row jam editable worker+asisten

Tambah range protection A2:C9 di _Settings yang override sheet-level
admin protection. Editor = owner + admin + email worker (Master_Data
kolom C) + email asisten (kolom E) untuk semua staf aktif.

Hasil:
- Row 1 (header), row 10 (separator), row 11-15 (operational): admin only
- Row 2-9 (jam: MASUK, IST1/2/3, PULANG): admin + worker + asisten

Helper _setProteksiSettingsJam(sheet) di-extract untuk dipakai juga
oleh perbaruiAksesSettings() di commit berikutnya.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Fungsi `perbaruiAksesSettings` + menu item

**Files:**
- Modify: `Setup.js` (tambah fungsi `perbaruiAksesSettings` setelah `buatSheetSettings`)
- Modify: `Triggers.js` (tambah menu item)

- [ ] **Step 2.1: Tambah fungsi `perbaruiAksesSettings` di Setup.js**

Sisipkan blok berikut di Setup.js **setelah** closing `}` fungsi `buatSheetSettings()` (sekitar [Setup.js:230](Setup.js#L230), tepat sebelum `// ── setupProteksiMaster ──`):

```js

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
```

- [ ] **Step 2.2: Tambah menu item di Triggers.js**

Cari blok menu admin di [Triggers.js:170-185](Triggers.js#L170-L185). Replace baris yang berisi `'🔄 Migrasi Sheet Tambah Ist 3'`:

```js
    .addItem('🔄 Migrasi Sheet Tambah Ist 3',  'migrateSheetTambahIst3')
    .addItem('⚙️ Buat/Reset Sheet Settings',  'buatSheetSettings')
```

dengan:

```js
    .addItem('🔄 Migrasi Sheet Tambah Ist 3',  'migrateSheetTambahIst3')
    .addItem('🔑 Perbarui Akses Settings',     'perbaruiAksesSettings')
    .addItem('⚙️ Buat/Reset Sheet Settings',  'buatSheetSettings')
```

- [ ] **Step 2.3: Verifikasi syntax**

Run: `node --check Setup.js && node --check Triggers.js && echo OK`
Expected: `OK`.

- [ ] **Step 2.4: Commit Task 2**

```bash
git add Setup.js Triggers.js
git commit -m "$(cat <<'EOF'
feat(setup): perbaruiAksesSettings + menu untuk refresh editor _Settings

Fungsi perbaruiAksesSettings() menggunakan helper _setProteksiSettingsJam
untuk refresh editor list A2:C9 ketika Master_Data berubah, tanpa
harus reset seluruh sheet _Settings.

Mirror pattern dengan perbaruiProteksiAdmin() yang sudah ada.

Menu admin baru: 🔑 Perbarui Akses Settings (di antara
"Migrasi Sheet Tambah Ist 3" dan "Buat/Reset Sheet Settings").

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Manual verification (handoff ke user)

Bagian ini untuk HRD yang menjalankan deploy dan uji aktual.

- [ ] **Step 3.1: Deploy via clasp**

```bash
./deploy.sh
```

Expected: clean push ke TARGETS.

- [ ] **Step 3.2: Reset _Settings di TESTING WORKER 2 (smoke test)**

1. Buka spreadsheet TESTING WORKER 2 di browser
2. Reload tab supaya menu `onOpen` rebuild dengan code baru
3. Menu `🔧 Admin` → `⚙️ Buat/Reset Sheet Settings`
4. Konfirm "Reset dan buat ulang?" → YES
5. Verifikasi UI alert: `✅ Sheet _Settings berhasil dibuat!` dengan baris
   `• Row 2–9 (jam): admin + N worker + N asisten`
6. Buka sheet `_Settings`:
   - Klik row 5 (mis. IST2_MULAI). Edit VALUE. Klik kanan → Protect range → cek "Restrictions" → harus list owner + admin + worker + asisten
   - Klik row 12 (mis. JAM_REMINDER). Klik kanan → Protect range → harus list owner + admin saja (sheet-level protection)

- [ ] **Step 3.3: Uji sebagai worker (akun non-admin)**

1. Login sebagai email worker (mis. `prada.dipa@gmail.com` kalau itu worker, bukan asisten)
2. Buka spreadsheet TESTING WORKER 2 di browser
3. Buka sheet `_Settings`
4. Klik sel B5 (VALUE untuk IST2_MULAI). Ketik `14:30`. Tekan Enter.
   Expected: nilai tersimpan, tidak ada permission warning.
5. Klik sel B12 (VALUE untuk JAM_REMINDER). Ketik `18`. Tekan Enter.
   Expected: Google Sheets show "You are trying to edit a protected cell..." dialog. Cancel.

- [ ] **Step 3.4: Uji sebagai asisten**

Ulangi step 3.3 dengan akun asisten (mis. `prada.dipa@gmail.com` kalau dia asisten). Expected sama: bisa edit B5, blocked di B12.

- [ ] **Step 3.5: Uji `perbaruiAksesSettings()`**

1. Di Master_Data, tambah baris worker baru atau ubah email asisten salah satu worker
2. Sebagai admin, menu `🔧 Admin` → `🔑 Perbarui Akses Settings`
3. Expected UI alert dengan count worker/asisten terbaru
4. Login sebagai email baru yang ditambah, coba edit row 5 → harus bisa

- [ ] **Step 3.6: Rollout ke production**

Setelah TESTING WORKER 2 lulus, ulangi step 3.2 untuk 10 spreadsheet production di TARGETS.

---

## Self-Review Notes (already addressed inline)

- **Spec coverage**:
  - Skema 2-lapis proteksi → Task 1 Step 1.1 (helper) + Step 1.2 (integrasi).
  - Editor dari Master_Data kolom C + E → Task 1 Step 1.1 (loop `masterData`).
  - Helper extract DRY → Task 1 Step 1.1 (`_setProteksiSettingsJam`).
  - `perbaruiAksesSettings()` → Task 2 Step 2.1.
  - Menu item → Task 2 Step 2.2.
  - Edge case email invalid → try/catch di addEditor di Task 1.
  - Manual verification per skenario → Task 3.

- **No placeholders**: Setiap step punya kode lengkap atau perintah konkret.

- **Type consistency**: `_setProteksiSettingsJam(sheet)` consistent — return `{ countWorker, countAsisten }` di Task 1.1, dipakai di Task 1.2 (`buatSheetSettings`) dan Task 2.1 (`perbaruiAksesSettings`).

- **Order matters**: Task 1 dan Task 2 saling independen di sisi kode (Task 2 pakai helper dari Task 1), tapi commit-nya self-contained dan bisa di-revert per task.
