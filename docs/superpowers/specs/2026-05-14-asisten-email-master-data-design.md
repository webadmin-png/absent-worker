# Email Asisten di Master_Data — Design Spec

**Tanggal**: 2026-05-14
**Status**: Approved, siap implementasi
**File terdampak**: `Append.js`, `Setup.js`, `Config.js` (komentar saja)

## Tujuan

Memungkinkan dua email mengisi absensi seorang worker: email worker itu sendiri **dan** email asisten yang ditunjuk. Mapping bersifat 1-ke-1 (satu worker punya nol atau satu asisten).

## Konteks

Saat ini Master_Data hanya menyimpan satu email per worker di kolom C. Email itu jadi sumber kebenaran untuk:
- Identifikasi user di menu Stamp via `getInfoUser()` ([Utils.js:40-61](../../../Utils.js#L40-L61))
- Editor proteksi range E:K per baris di `proteksiBarisBaru()` ([Setup.js:432-469](../../../Setup.js#L432-L469))
- Display di kolom D sheet absensi harian ([Append.js:104](../../../Append.js#L104))

Beberapa worker tidak punya akses Google Account konsisten (mis. pekerja produksi yang shift atau tidak melek IT). HRD ingin menunjuk asisten — bisa rekan kerja atau leader — yang berwenang mengisi absen atas nama worker tersebut.

## Keputusan Desain

### Mapping: 1-ke-1
Setiap baris Master_Data punya **satu kolom asisten opsional**. Asisten X hanya bisa edit baris worker A, bukan worker lain. Tidak menggunakan `ADMIN_EMAILS` karena asisten harus terikat ke worker spesifik (principle of least privilege).

### Stamp behavior: tidak berubah
Asisten **tidak menggunakan menu Stamp**. Asisten edit jam absen worker secara manual di sel (ketik langsung). `getInfoUser()` dan semua fungsi di `Stamp.js` tidak berubah.

Konsekuensi: jika asisten klik menu Stamp, dia akan dapat error "Email tidak terdaftar di Master_Data" — diterima sebagai behavior yang benar.

### Display: hanya di Master_Data
Sheet absensi harian **tidak menampilkan email asisten**. `TOTAL_COL` tetap 21, semua `COL_*` constant tetap, tidak ada perubahan formula atau header sheet absensi.

Sumber kebenaran tentang siapa yang bisa edit baris X: panel Google Sheets Protection + Master_Data. HRD yang ingin tahu daftar asisten lihat di Master_Data kolom E.

## Skema Master_Data Baru

```
A         B      C              D       E (baru)
Divisi    Nama   Email Worker   Aktif   Email Asisten
WORKER    Andi   andi@x.com     TRUE    asisten1@x.com
WORKER    Budi   budi@x.com     TRUE    (kosong — tidak punya asisten)
WORKER    Cindy  cindy@x.com    TRUE    asisten2@x.com
```

Kolom E **opsional** per baris. String kosong berarti worker tidak punya asisten — behavior persis seperti sekarang.

## Perubahan Kode

### Append.js

1. Line 31 — perluas range read:
   ```js
   // sebelum
   const masterData = master.getRange('A4:D200').getValues()
   // sesudah
   const masterData = master.getRange('A4:E200').getValues()
   ```

2. Line 47-50 — tambah `emailAsisten` ke object staf:
   ```js
   stafPerDivisi[div].push({
     nama         : String(k[1]).trim(),
     email        : String(k[2]).trim(),
     emailAsisten : String(k[4] || '').trim(),
   });
   ```

3. `newRows` tidak berubah — sheet absensi tetap 21 kolom.

### Setup.js — `proteksiBarisBaru()`

1. Line 407 — perluas range read:
   ```js
   // sebelum
   const masterData = master.getRange('A4:D200').getValues()
     .filter(r => r[0] !== '' && String(r[0]).trim().toUpperCase() === ...
   // sesudah
   const masterData = master.getRange('A4:E200').getValues()
     .filter(r => r[0] !== '' && String(r[0]).trim().toUpperCase() === ...
   ```

2. Setelah blok `prot.addEditor(email)` di sekitar line 463, sisipkan:
   ```js
   const asisten = String(k[4] || '').trim();
   if (asisten) {
     try {
       prot.addEditor(asisten);
       Logger.log('✓ Proteksi asisten: ' + nama + ' ← ' + asisten);
     } catch(e) {
       Logger.log('⚠ Gagal tambah asisten ' + asisten + ': ' + e.message);
     }
   }
   ```

`perbaruiProteksiAdmin()` tidak berubah — itu fungsi untuk admin global, bukan asisten per baris.

### Config.js

Hanya tambah komentar di area dokumentasi (sekitar line 158-178) yang menjelaskan skema Master_Data baru. Tidak ada `COL_*` constant baru karena Master_Data tidak punya constant (dibaca langsung via `A4:E200`).

### Master_Data sheet header

HRD update manual sekali: tambah label "Email Asisten" di sel E3 (atau di mana header berada). Tidak ada kode untuk ini — header Master_Data memang dikelola manual.

## Migrasi

### Baris hari ini & masa depan
`appendHariIni()` jalan pagi 06:00 → otomatis tarik kolom E saat membaca Master_Data → `proteksiBarisBaru()` set editor asisten. Tidak perlu intervensi.

### Baris yang sudah di-append sebelum perubahan
Dua skenario:

1. **Sudah di-lock oleh `lockBarisWebSudahPulang`**: tidak perlu apa-apa. Lock memang membatasi ke owner + admin saja, asisten memang tidak boleh edit. Diterima.

2. **Belum di-lock (mis. hari ini, worker belum pulang)**: kalau HRD baru tambah asisten siang ini, baris hari ini tidak otomatis pick up. `proteksiBarisBaru()` butuh parameter `(sheet, divisi, startRow, numRows)` jadi tidak bisa dipanggil langsung dari menu.

Untuk implementasi pertama: **tidak buat helper migrasi otomatis**. Cukup deploy lalu HRD update asisten di Master_Data sebelum 06:00 esok hari — baris hari berikutnya otomatis pick up. Kalau ada kasus mendesak hari ini, HRD bisa tambahkan email asisten manual via panel Protection Google Sheets (klik kanan baris → Protect range → tambah editor). Helper fungsi migrasi bisa ditambahkan kemudian jika kasus mendesak sering terjadi.

## Behavior Edge Cases

| Skenario | Behavior |
|---|---|
| Worker tanpa asisten (kolom E kosong) | Guard `if (asisten)` skip → behavior identik sekarang |
| Email asisten salah/invalid | `addEditor` lempar exception → di-catch, log warning, worker tetap dapat akses |
| Asisten email sama dengan email admin | `addEditor` di-call dua kali → Google Sheets idempotent, no-op |
| Asisten kebetulan juga worker (punya baris sendiri) | `getInfoUser` cari exact-match di kolom C → dapat baris dia sendiri saat stamp. Saat edit worker lain → punya editor access dari proteksi |
| Asisten email muncul di banyak baris Master_Data (1 asisten bantu banyak worker) | Secara teknis tidak masalah — Google Sheets izinkan email yang sama jadi editor banyak range. Bertentangan dengan keputusan 1-ke-1 secara konseptual, tapi tidak memblok teknis |
| HRD ubah asisten setelah baris sudah di-append hari ini | Tidak otomatis update. HRD harus rerun `proteksiBarisBaru()` |
| Asisten mencoba klik menu Stamp | Dapat error "Email tidak terdaftar" — by design |
| Baris di-lock setelah jam pulang | Asisten kehilangan akses (sama seperti worker) — by design |

## Estimasi Effort

- ~15-20 baris kode diubah, 3 file disentuh.
- Tidak ada perubahan di `Stamp.js`, `Lock.js`, `Rekap.js`, `Triggers.js`, `Utils.js`.
- Tidak ada perubahan di formula L:O, range proteksi, atau `TOTAL_COL`.
- Risiko regresi: rendah.

## Tidak Termasuk dalam Spec Ini (Out of Scope)

- Fungsi migrasi otomatis untuk baris yang sudah ada.
- Tampilan email asisten di sheet absensi (Pendekatan A/B yang ditolak).
- Banyak asisten per worker.
- Notifikasi ke asisten saat worker belum isi pulang.
- UI/form untuk HRD menambah asisten (HRD edit Master_Data langsung).

Jika requirement bertambah, refactor ke pendekatan kolom-di-sheet tetap mungkin dilakukan — basis kode tidak akan punya teknis debt karena perubahan saat ini sangat lokal.
