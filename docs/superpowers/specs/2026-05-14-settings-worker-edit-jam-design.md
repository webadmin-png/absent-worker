# _Settings Row Jam Editable oleh Worker + Asisten — Design Spec

**Tanggal**: 2026-05-14
**Status**: Approved, siap implementasi
**File terdampak**: `Setup.js`, `Triggers.js`

## Tujuan

Mengizinkan akun worker dan asisten (yang terdaftar di `Master_Data`) untuk mengedit sendiri jam-jam absen default mereka (MASUK, istirahat 1/2/3, PULANG) di sheet `_Settings` row 2–9. Row 11–15 (operational: DIVISI, JAM_REMINDER, SELISIH_MENIT_LOCK, PLAN_JAM, ADMIN_EMAILS) tetap admin only.

## Konteks

Sheet `_Settings` saat ini sepenuhnya protected admin-only via `sheet.protect()` di [Setup.js:211-218](../../../Setup.js#L211-L218). Worker yang ingin mengubah jam masuk default (mis. dari `08:00` ke `08:30`) harus minta tolong admin. Hal ini friction; worker seharusnya bisa atur sendiri tanpa kompromi keamanan setting lain.

Setelah Task 2 fitur Ist 3, struktur `_Settings`:

```
Row 1   : KEY | VALUE | Keterangan (header)
Row 2   : MASUK            ← jam, akan editable worker+asisten
Row 3   : IST1_MULAI       ← jam
Row 4   : IST1_SELESAI     ← jam
Row 5   : IST2_MULAI       ← jam
Row 6   : IST2_SELESAI     ← jam
Row 7   : IST3_MULAI       ← jam
Row 8   : IST3_SELESAI     ← jam
Row 9   : PULANG           ← jam
Row 10  : (separator kosong)
Row 11  : DIVISI            ← tetap admin only
Row 12  : JAM_REMINDER      ← tetap admin only
Row 13  : SELISIH_MENIT_LOCK ← tetap admin only
Row 14  : PLAN_JAM          ← tetap admin only
Row 15  : ADMIN_EMAILS      ← tetap admin only
```

## Keputusan Desain

### Skema proteksi: sheet-level + range-level override
Pakai dua lapis proteksi:

1. **Sheet-level protection** — owner + admin (existing, tidak berubah). Default-deny untuk seluruh sheet.
2. **Range protection `A2:C9`** — owner + admin + tiap email worker (Master_Data kolom C) + tiap email asisten (kolom E). Override sheet protection di area tersebut.

Google Sheets selalu prioritaskan range-level protection yang lebih spesifik daripada sheet-level. Hasil:
- Cells **di dalam `A2:C9`**: dipakai range protection → worker/asisten/admin bisa edit
- Cells **di luar `A2:C9`** (termasuk row 1 header, row 10 separator, row 11-15 operational): dipakai sheet protection → admin only

### Editor list dari Master_Data
Worker + asisten dari `Master_Data` dibaca saat `buatSheetSettings()` jalan. Filter sama dengan tempat-tempat lain di codebase: `r[0] !== '' && r[3] === 'TRUE'` (divisi terisi + status aktif).

### Refresh saat Master_Data berubah
Karena email editor di-snapshot saat protection dibuat, perubahan Master_Data tidak otomatis terpropagasi. Solusi: helper baru `perbaruiAksesSettings()` yang admin-callable via menu — sama pola dengan `perbaruiProteksiAdmin()` yang sudah ada. Helper ini hanya refresh editor `A2:C9`, **tidak** menyentuh struktur sheet atau row jam.

### Range A2:C9 (3 kolom)
Sheet `_Settings` hanya pakai kolom A (KEY), B (VALUE), C (Keterangan). Range proteksi cover persis 3 kolom × 8 row. Kolom D+ secara teknis tidak ada konten — tapi sheet protection tetap menutupnya untuk konsistensi.

## Perubahan File

### `Setup.js`

1. **Modifikasi `buatSheetSettings()`** — di akhir fungsi (setelah sheet-level protect), tambah blok:
   - Buat range protection `A2:C9`
   - Set description, warning-only false, remove default editors
   - Add owner + admin emails dari `CONFIG.ADMIN_EMAILS`
   - Baca `Master_Data!A4:E200`, filter aktif
   - Loop tiap row aktif: addEditor untuk worker email (kolom C) + asisten email (kolom E, jika ada)
   - Wrap tiap addEditor di try/catch agar email invalid tidak gagalkan deploy

2. **Fungsi baru `perbaruiAksesSettings()`** — di-append setelah `buatSheetSettings()`:
   - `_requireAdmin()` + `_loadSettings()`
   - Ambil sheet `_Settings`, throw kalau tidak ada
   - Cari proteksi range A2:C9 existing → kalau ada, `.remove()`
   - Buat ulang protection persis seperti di `buatSheetSettings()` (DRY: extract helper jika ada nilai)
   - UI alert dengan ringkasan editor yang ditambah

   Decision DRY: helper internal `_setProteksiSettingsJam(sheet)` yang dipanggil oleh kedua fungsi. Lebih bersih.

### `Triggers.js`

Tambah menu item di `onOpen()` setelah baris `🔄 Migrasi Sheet Tambah Ist 3`:
```js
.addItem('🔑 Perbarui Akses Settings',     'perbaruiAksesSettings')
```

## Edge Cases

| Skenario | Behavior |
|---|---|
| Worker tanpa asisten (kolom E kosong) | Skip addEditor untuk asisten, worker tetap ditambah |
| Email worker/asisten invalid | `addEditor` throw → di-catch, log warning, lanjut email berikutnya |
| Master_Data berubah (worker baru/email diubah) | `_Settings` editor masih lama. HRD run `perbaruiAksesSettings()` |
| User non-admin coba edit row 11 (DIVISI) | Sheet protection block → Google Sheets show permission warning |
| Worker email coba edit row 5 (IST2_MULAI) | Range A2:C9 protection allow karena email ada di editor list |
| User dengan edit access tapi bukan worker/asisten/admin coba edit row 5 | Range protection block — email tidak ada di editor list |
| Sheet `_Settings` belum ada saat `perbaruiAksesSettings()` dipanggil | Throw error: "Run 'Buat/Reset Sheet Settings' dulu" |

## Workflow HRD

```
1. (sekali) ./deploy.sh atau clasp push code baru
2. Untuk SETIAP spreadsheet:
   a. Buka spreadsheet
   b. Menu 🔧 Admin → ⚙️ Buat/Reset Sheet Settings → konfirm Reset
      → _Settings dibuat ulang dengan proteksi 2-lapis
   c. (atau, kalau _Settings sudah ada dan tidak mau reset)
      Menu 🔧 Admin → 🔑 Perbarui Akses Settings
      → editor row A2:C9 di-refresh dari Master_Data terkini
3. Worker login, buka _Settings, edit row 5 (IST2_MULAI) — harus bisa
4. Worker coba edit row 11 (DIVISI) — harus dapat permission warning
```

## Out of Scope

- Auto-refresh saat Master_Data diedit (butuh trigger onEdit yang baca Master_Data — overkill).
- Per-worker editor (mis. worker A hanya boleh edit MASUK = jam dia sendiri). Saat ini semua worker share akses ke seluruh row 2-9.
- UI form untuk worker edit jam-nya — pakai Google Sheets langsung sudah cukup.
- Audit log siapa edit apa di _Settings.

## Estimasi

- `Setup.js`: ~50 baris (modifikasi buatSheetSettings + helper `_setProteksiSettingsJam` + fungsi `perbaruiAksesSettings`)
- `Triggers.js`: 1 baris (menu)
- File disentuh: 2
- Risiko regresi: low — perubahan hanya tambah protection, tidak mengubah data atau formula
