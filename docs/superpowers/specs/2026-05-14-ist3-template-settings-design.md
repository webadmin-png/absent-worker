# Istirahat 3 di Template & Settings — Design Spec

**Tanggal**: 2026-05-14
**Status**: Approved per section, siap implementasi
**File terdampak**: `Config.js`, `Setup.js`, `Append.js`. Plus migration helper baru di `Setup.js`.

## Tujuan

Menambahkan kolom Istirahat 3 (`Mulai` + `Selesai`) ke template sheet absensi dan `_Settings`, sehingga staf bisa mencatat istirahat ketiga dan jam tersebut otomatis dipotong dari Jam Efektif. Kolom-kolom Ist 3 disisipkan dalam urutan logis (setelah Ist 2, sebelum Pulang).

## Konteks

Sheet absensi saat ini menyimpan 2 istirahat (`Ist 1` di G/H, `Ist 2` di I/J). Skema 21 kolom (A–U) dengan `Pulang` di K dan kolom formula (L–O) di tengah. Beberapa divisi memerlukan tiga istirahat (mis. shift panjang dengan break makan siang + dua coffee break).

## Keputusan Desain

### Posisi kolom
Sisipkan **Ist 3 Mulai** di K dan **Ist 3 Selesai** di L. Semua kolom yang sebelumnya K–U geser +2 ke M–W. Pilihan ini dipilih dibanding "append di belakang" karena UX (kolom istirahat terurut) lebih penting daripada effort implementation; perubahan-nya banyak tapi mekanis.

### Formula
Refactor formula `Jam Efektif` (L → N setelah shift) dari nested-IF ke pola flat (yang sudah dipakai OT 1 / OT 2). Tambah subtract `ist3_duration` ke L, OT 1, OT 2. Hasil matematis identik dengan formula nested, struktur lebih bersih, ekstensi ke ist4/ist5 di masa depan jadi trivial.

### Stamp menu
**Tidak** ditambahkan menu Stamp Ist 3. Staf harus ketik manual jam ist3 di sel — sama treatment dengan kolom `KETERANGAN`, `PLAN`. Keputusan ini dikonfirmasi user.

### Migrasi
Sheet bulan ini yang sudah dibuat dengan 21 kolom **tidak** kompatibel dengan kode baru tanpa migrasi. Disediakan fungsi `migrateSheetTambahIst3()` (manual-trigger via menu admin) yang idempotent dan insert 2 kolom di posisi 11 untuk tiap sheet divisi. Google Sheets auto-shift referensi formula existing saat kolom di-insert.

## Skema Kolom Baru

| Kolom | Sebelum | Setelah | Constant |
|---|---|---|---|
| A | Tanggal | Tanggal | `COL_TANGGAL=1` |
| B | Hari | Hari | `COL_HARI=2` |
| C | Nama | Nama | `COL_NAMA=3` |
| D | Email | Email | `COL_EMAIL=4` |
| E | Status | Status | `COL_STATUS=5` |
| F | Masuk | Masuk | `COL_MASUK=6` |
| G | Ist 1 Mulai | Ist 1 Mulai | `COL_IST1_M=7` |
| H | Ist 1 Selesai | Ist 1 Selesai | `COL_IST1_S=8` |
| I | Ist 2 Mulai | Ist 2 Mulai | `COL_IST2_M=9` |
| J | Ist 2 Selesai | Ist 2 Selesai | `COL_IST2_S=10` |
| K | ~~Pulang~~ | **Ist 3 Mulai** | `COL_IST3_M=11` ⭐ baru |
| L | ~~Jam Efektif~~ | **Ist 3 Selesai** | `COL_IST3_S=12` ⭐ baru |
| M | ~~Regular Hours~~ | Pulang | `COL_PULANG=13` (was 11) |
| N | ~~OT 1~~ | Jam Efektif | `COL_EFEKTIF=14` (was 12) |
| O | ~~OT 2~~ | Regular Hours | `COL_REGULAR_JAM=15` (was 13) |
| P | ~~NOTE~~ | OT 1 | `COL_OT1=16` (was 14) |
| Q | ~~SUNDAY~~ | OT 2 | `COL_OT2=17` (was 15) |
| R | ~~KETERANGAN~~ | NOTE | `COL_NOTE=18` (was 16) |
| S | ~~PLAN~~ | SUNDAY/RED DAY | `COL_SUNDAY=19` (was 17) |
| T | ~~CATATAN TELAT~~ | KETERANGAN | `COL_KETERANGAN=20` (was 18) |
| U | ~~CATATAN PULANG AWAL~~ | PLAN | `COL_PLAN=21` (was 19) |
| V | — | CATATAN TELAT | `COL_TELAT=22` (was 20) |
| W | — | CATATAN PULANG AWAL | `COL_PULANG_AWAL=23` (was 21) |

`TOTAL_COL=23`. `COL_EDIT_END = COL_PULANG_AWAL = 23` (otomatis).

## Perubahan File

### `Config.js`

1. Update 11 `COL_*` constant values + tambah `COL_IST3_M=11`, `COL_IST3_S=12`. `TOTAL_COL=23`.
2. Update comment block "Mapping kolom sheet" dengan layout baru.
3. `AUTO_ABSENSI` per divisi — tambah field `ist3Mulai: ''` dan `ist3Selesai: ''` (default kosong, opsional).
4. `_loadSettings()` autoFieldMap — tambah entry `IST3_MULAI: 'ist3Mulai'` dan `IST3_SELESAI: 'ist3Selesai'`.
5. `_loadSettings()` fallback default object — tambah `ist3Mulai: '', ist3Selesai: ''`.

### `Append.js`

1. `newRows` di `appendHariIni()` — sisipkan 2 cell `toSerial(auto.ist3Mulai)` dan `toSerial(auto.ist3Selesai)` di posisi 11 dan 12 (sebelum pulang).
2. setNumberFormat HH:mm array — tambah `COL_IST3_M, COL_IST3_S`.
3. Color band E:M `setBackground` literal `(insertAt, 5, n, 7)` → `(insertAt, 5, n, 9)` (cover ist3).
4. `_pasangFormulaBaris()`:
   - Local vars: tambah `m=\`M${r}\`` (pulang), `n=\`N${r}\`` (jam efektif self-ref).
   - Formula L (Jam Efektif) — refactor ke pola flat, subtract ist3:
     ```
     =IF(E="Red Day",0,
        IF(status!="Hadir",0,
           IF(masuk="" OR pulang="",0,
              pulang-masuk
                -IF(AND(g!="",h!=""),h-g,0)
                -IF(AND(i!="",j!=""),j-i,0)
                -IF(AND(k!="",l!=""),l-k,0))))
     ```
     Catatan: formula lama tidak include Red Day check di L — dia hanya di M (Regular Hours). Tetap.
   - Formula M (Regular Hours) — referensi `${l}` (jam efektif) → `${n}`.
   - Formula N (OT 1) — pulang `${k}` → `${m}`, tambah term `-IF(AND(${k}<>"",${l}<>""),${l}-${k},0)`.
   - Formula O (OT 2) — sama: pulang `${k}` → `${m}`, tambah subtract ist3.

### `Setup.js`

1. `buatSheetBulanBaru()`:
   - Header array — sisipkan 2 entry "Ist. Ketiga Mulai/Selesai" di posisi 10-11 (setelah Ist Kedua Selesai).
   - Column widths array — sisipkan `90,90` di posisi 10.
   - Legenda baris 2 — update span dan posisi setiap segment.
   - Header protection range `'A1:S3'` → `'A1:W3'`.
2. `buatSheetSettings()`:
   - rows array — sisipkan `['IST3_MULAI', ..., ...]` dan `['IST3_SELESAI', ..., ...]` sebelum row `PULANG`.
   - `setNumberFormat('HH:mm')` range — `(2, 2, 5, 1)` → `(2, 2, 7, 1)` (7 row jam, bukan 5).
3. `setupValidasiBaris()`:
   - Loop validation jam — array `['G','H','I','J','K']` → `['G','H','I','J','K','L','M']` (tambah ist3 mulai/selesai + pulang yang sekarang di M).
4. `proteksiBarisBaru()` — TIDAK perlu kode berubah. Range pakai `COL_PULANG - COL_STATUS + 1` yang otomatis adjust dari 7 (E:K) ke 9 (E:M) karena `COL_PULANG` shift 11→13.
5. **Fungsi baru** `migrateSheetTambahIst3()`:
   - Iterate semua sheet yang nama-nya match `<DIVISI>_<MMM>_<YYYY>` atau `<DIVISI>` (sesuai pola di `getSheetAktifDivisi`).
   - Cek idempotency: kalau header sel `K3` sudah berisi "Ist. Ketiga" (atau lastColumn >= 23), skip dengan log.
   - `sheet.insertColumnsBefore(11, 2)` — insert 2 kolom kosong. Google Sheets auto-shift formula yang sudah ada.
   - Set header sel `K3` dan `L3` dengan teks + format (hijau toska, bold, wrap, dst.) sama dengan ist1/ist2.
   - Set column width K dan L = 90.
   - Update merge baris 1 (judul) — jika range existing A1:S1 atau A1:U1, extend ke A1:W1.
   - Update legenda baris 2 — re-merge segments dengan span dan posisi baru.
   - Re-run `_pasangFormulaBaris(sheet, startToday, numToday)` untuk baris hari ini supaya formula upgrade ke versi flat baru yang subtract ist3.
   - Re-run `proteksiBarisBaru(sheet, divisi, startToday, numToday)` untuk fix range proteksi yang shift.
   - Log per sheet dengan emoji ✓ atau ⏭.
   - Tampilkan UI alert ringkasan di akhir.
   - Tambahkan menu item `🔄 Migrasi Sheet Tambah Ist 3` di `onOpen` (Triggers.js) — admin only.

### `Triggers.js`

1. Menu admin di `onOpen` — tambah item baru:
   ```js
   .addItem('🔄 Migrasi Sheet Tambah Ist 3', 'migrateSheetTambahIst3')
   ```
   Sebaiknya di seksi "Setup" (setelah `Setup Awal`, sebelum `Setup Trigger`), atau di bagian terpisah dengan separator.
2. `onEditInstalled` — tidak perlu berubah. Logic-nya pakai `COL_EDIT_START`, `COL_EFEKTIF`, `COL_NOTE`, `COL_SUNDAY`, `COL_EMAIL` yang semua adalah constant — otomatis ikut shift.

### `Lock.js`, `Stamp.js`, `Rekap.js`

Tidak ada perubahan eksplisit. Semua referensi ke kolom pakai `COL_*` constant yang auto-update. Verifikasi spot-check:
- `Lock.js`: pakai `COL_TANGGAL`, `COL_NAMA`, `COL_STATUS`, `COL_MASUK`, `COL_PULANG`, `COL_EMAIL` — semua constant ✓
- `Stamp.js`: pakai `COL_MASUK`, `COL_IST1_M`, `COL_IST1_S`, `COL_IST2_M`, `COL_IST2_S`, `COL_PULANG`, `COL_STATUS`, `COL_NAMA` — semua constant ✓
- `Rekap.js`: belum di-spot-check secara detail, tapi diasumsikan pakai constants. Implementation plan akan verifikasi.

## Migration Workflow (untuk HRD)

```
1. ./deploy.sh                              # deploy code baru ke semua target
2. Per spreadsheet target:
   a. Buka spreadsheet
   b. Menu 🔧 Admin → 🔄 Migrasi Sheet Tambah Ist 3
      → log: "✓ <sheet>: migrated" atau "⏭ <sheet>: sudah migrated, skip"
   c. (Opsional) menu 🔧 Admin → ⚙️ Buat/Reset Sheet Settings
      atau edit manual: tambah 2 row IST3_MULAI/IST3_SELESAI di _Settings
3. Tunggu trigger 06:00 esok, atau test manual:
   - Edit nilai IST3_MULAI/IST3_SELESAI di _Settings
   - Hapus baris hari ini (kalau perlu)
   - Run appendHariIni() manual
   - Cek kolom K, L terisi dengan jam ist3
   - Cek formula jam efektif subtract ist3 dengan benar
```

## Edge Cases

| Skenario | Behavior |
|---|---|
| Worker isi Ist 3 Mulai tapi tidak Selesai | Formula L flat: `IF(AND(k!="",l!=""), l-k, 0)` → 0 (tidak subtract). Tidak error. |
| Worker isi Ist 3 Selesai > 24:00 (mustahil di praktek) | Sheets handle sebagai negative selisih, formula tetap jalan tapi hasilnya aneh. Validasi format HH:MM mencegah ini. |
| Sheet bulan ini sudah dimigrasi, lalu migrasi dijalankan ulang | Idempotent: log "⏭ skip" dan return. Tidak error, tidak duplicate. |
| Sheet 21-col belum dimigrasi tapi appendHariIni jalan | NEW rows ditulis dengan struktur 23-col → schema mismatch dengan header existing. **Harus migrate dulu sebelum trigger jalan.** |
| `_Settings` belum punya row IST3_*  | `_loadSettings` skip baris yang key-nya bukan known, default fallback `''` → ist3 tidak auto-fill (kosong). Aman. |
| Divisi WORKER set `IST3_MULAI=14:30`, `IST3_SELESAI=14:45` di _Settings | appendHariIni esok pagi tulis `14:30`/`14:45` di kolom K/L, formula jam efektif kurangi 15 menit. |
| Auto-fill ist3 tapi worker tidak ist3 hari itu | Worker harus hapus manual nilai di K/L kalau memang tidak istirahat. Sama pattern dengan ist1/ist2. |

## Out of Scope

- Stamp menu untuk Ist 3 (sengaja, sesuai keputusan user).
- Generic "ist-N" framework untuk N istirahat dinamis (premature abstraction).
- Migrasi sheet bulan-bulan **lalu** yang sudah lock — tidak perlu, baris lama sudah read-only.
- Update template Web App (kalau ada Web App rendering kolom) — tidak ada Web App di repo saat ini.
- Validasi total durasi istirahat (mis. max 2 jam/hari) — tidak diminta.

## Estimasi Effort

- File disentuh: 4 (Config.js, Append.js, Setup.js, Triggers.js).
- Kode berubah/baru: ~120 baris (termasuk migration function ~60 baris).
- Risiko regresi: **medium** karena banyak constant geser, walaupun mekanis. Mitigasi: migration function idempotent, formula refactor diuji manual di TESTING WORKER 2 sebelum rollout ke spreadsheet lain.
- Effort migrasi production (per spreadsheet): 1-2 menit (1 klik menu admin + verifikasi).
