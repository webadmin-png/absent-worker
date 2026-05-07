# SOP Sistem Absensi — PT InFashion

> Versi sistem: Sheet-based (staf isi langsung di spreadsheet via menu)
> Setiap divisi memiliki spreadsheet terpisah.

---

## DAFTAR ISI

1. [Cara Kerja Sistem](#1-cara-kerja-sistem)
2. [Setup Pertama Kali](#2-setup-pertama-kali)
3. [Setiap Hari — Staf](#3-setiap-hari--staf)
4. [Setiap Hari — Admin/HR](#4-setiap-hari--adminhr)
5. [Tiap Bulan Baru](#5-tiap-bulan-baru)
6. [Tambah Staf Baru](#6-tambah-staf-baru)
7. [Tambah Admin Baru](#7-tambah-admin-baru)
8. [Tambah Divisi Baru](#8-tambah-divisi-baru)
9. [Ubah Settings Operasional](#9-ubah-settings-operasional)
10. [Melihat Rekap](#10-melihat-rekap)
11. [Jika Ada Masalah](#11-jika-ada-masalah)

---

## 1. Cara Kerja Sistem

```
Tiap divisi punya 1 spreadsheet Google Sheets yang sudah terpasang script otomatis.

Setiap hari pukul 06:00  →  baris staf ditambahkan otomatis
Setiap tgl 1 pukul 05:00 →  sheet bulan baru dibuat otomatis
Setiap pukul 17:00       →  reminder ke staf yang belum isi pulang
Setiap 1 jam             →  baris yang sudah pulang dikunci otomatis
```

**Struktur sheet per spreadsheet divisi:**

| Sheet | Isi |
|---|---|
| `Master_Data` | Daftar staf aktif divisi ini |
| `DIVISI_Mmm_yyyy` | Absensi bulan berjalan (contoh: `DEVELOPMENT_May_2026`) |
| `_Settings` | Konfigurasi operasional (jam, reminder, admin) |

**Hak akses:**

| Peran | Yang bisa dilakukan |
|---|---|
| Staf | Stamp masuk/pulang/istirahat via menu 📋 Absensi Saya |
| Admin/HR | Semua menu 🔧 Admin + edit semua kolom semua baris |
| Owner | Sama seperti Admin + bisa ubah kode dan deploy |

---

## 2. Setup Pertama Kali

> Dilakukan oleh **Owner** (developer). Cukup sekali per spreadsheet divisi baru.

1. Buka spreadsheet divisi
2. Klik **🔧 Admin → ⚙️ Setup Awal (pertama kali)**
3. Tunggu notifikasi ✅
4. Klik **🔧 Admin → ⚙️ Buat/Reset Sheet Settings** → isi jam auto-absensi dan pengaturan lain
5. Share spreadsheet ke semua staf divisi tersebut dengan akses **Editor**

---

## 3. Setiap Hari — Staf

> Staf membuka Google Spreadsheet divisinya, lalu menggunakan menu **📋 Absensi Saya**.

### Urutan penggunaan menu dalam sehari

```
Pagi   →  📍 Ke baris saya hari ini  (opsional, untuk navigasi cepat)
           ✅ Stamp MASUK

Siang  →  ☕ Stamp ISTIRAHAT 1 MULAI
           ▶  Stamp ISTIRAHAT 1 SELESAI

Sore   →  🏁 Stamp PULANG
```

> ⚠️ Data dikunci otomatis 30 menit setelah jam pulang diisi. Setelah terkunci, staf tidak bisa mengubah data — hubungi Admin/HR jika ada koreksi.

### Jika tidak hadir

1. Buka spreadsheet → klik **📍 Ke baris saya hari ini**
2. Isi kolom **Status** dengan: `Sakit`, `Izin`, atau `Alpha`
3. Isi kolom **Keterangan Tidak Hadir** dengan alasan singkat

### Jika terlambat masuk

1. Stamp MASUK seperti biasa
2. Isi kolom **CATATAN TELAT** dengan alasan

### Jika pulang lebih awal

1. Stamp PULANG seperti biasa
2. Isi kolom **CATATAN PULANG AWAL** dengan alasan

---

## 4. Setiap Hari — Admin/HR

> Dilakukan melalui **menu 🔧 Admin** di spreadsheet.

### Rutinitas harian

| Waktu | Tugas | Keterangan |
|---|---|---|
| Pagi | Cek baris sudah ter-append | Otomatis jam 06:00. Jika belum ada → klik **➕ Append Baris Hari Ini** |
| Kapan saja | Isi kolom NOTE (P) | Klik sel kolom P pada baris staf → pilih dari dropdown |
| Kapan saja | Isi kolom SUNDAY/RED DAY (Q) | Klik sel kolom Q → pilih SWAP / DOUBLE / HALF DAY SUNDAY |
| Sore | Cek yang belum isi pulang | **⚠️ Cek Belum Isi Pulang** (otomatis jam 17:00) |

### Opsi kolom NOTE (P) — admin only

| Pilihan | Keterangan |
|---|---|
| HALF DAY | Masuk setengah hari |
| RED DAY | Hari merah / libur nasional |
| RED DAY DOUBLE | Libur nasional, gaji double |
| VACATION PAID | Cuti berbayar |
| SICK PAID | Sakit berbayar |
| SICK UNPAID | Sakit tidak berbayar |
| FLEX DAY | Hari fleksibel |
| MATERNITY LEAVE | Cuti melahirkan |
| SWAP RED DAY | Tukar hari libur |
| DAY OFF UNPAID | Libur tanpa gaji |

---

## 5. Tiap Bulan Baru

> Sheet bulan baru dibuat **otomatis setiap tanggal 1 pukul 05:00**.
> Jika tidak terbuat otomatis, lakukan manual:

1. Buka spreadsheet divisi
2. Klik **🔧 Admin → 📅 Buat Sheet Bulan Baru**
3. Tunggu notifikasi ✅

---

## 6. Tambah Staf Baru

> Dilakukan oleh **Admin/HR** langsung di spreadsheet.

1. Buka spreadsheet divisi yang bersangkutan
2. Buka sheet **Master_Data**
3. Tambahkan baris baru:

   | Kolom A | Kolom B | Kolom C | Kolom D |
   |---|---|---|---|
   | Nama divisi (contoh: `DEVELOPMENT`) | Nama lengkap staf | Email Google staf | `TRUE` |

4. Share spreadsheet ke email staf tersebut dengan akses **Editor**
5. Staf akan muncul di sheet absensi mulai esok hari (atau jalankan **➕ Append Baris Hari Ini**)

> ⚠️ **Nama staf harus unik dalam satu divisi.** Jika ada dua staf bernama sama, tambahkan inisial — contoh: `Andi S.` dan `Andi W.`

### Menonaktifkan staf

- Buka **Master_Data** → ubah kolom D staf tersebut dari `TRUE` menjadi `FALSE`
- Staf tidak akan muncul lagi di append hari berikutnya

---

## 7. Tambah Admin Baru

> Tidak perlu sentuh kode. Cukup ubah _Settings di spreadsheet.

**Langkah:**

1. Buka spreadsheet divisi yang ingin ditambah admin-nya
2. Buka sheet **_Settings**
3. Temukan baris **ADMIN_EMAILS** di kolom KEY
4. Klik sel di kolom **VALUE** → tambahkan email baru pisah koma:
   ```
   email-lama@perusahaan.com, email-baru@perusahaan.com
   ```
5. Tekan Enter
6. Klik **🔧 Admin → 🔑 Perbarui Akses Admin**
7. Tunggu notifikasi ✅ — admin baru sudah bisa menggunakan semua fitur admin

> Ulangi langkah 1–6 untuk setiap spreadsheet divisi yang ingin diberikan akses.

---

## 8. Tambah Divisi Baru

> Dilakukan oleh **Admin/HR** — **tanpa menyentuh kode sama sekali**.
> Cukup duplicate spreadsheet yang sudah ada, lalu ubah nama divisi di _Settings.
> Estimasi waktu: 10 menit.

### Ringkasan langkah

```
A. Duplicate spreadsheet yang sudah ada
B. Ubah nama divisi di _Settings
C. Hapus sheet absensi bulan lama (dari spreadsheet asal)
D. Bersihkan dan isi ulang Master_Data
E. Jalankan Setup Awal
F. Atur jam dan admin di _Settings
G. Share ke staf divisi baru
```

---

### A. Duplicate Spreadsheet

1. Buka spreadsheet divisi yang sudah berjalan (contoh: `Absensi DEVELOPMENT`)
2. Klik menu **File → Make a copy**
3. Beri nama baru, contoh: `Absensi FINANCE 2026`
4. Pilih folder penyimpanan → klik **Make a copy**
5. Spreadsheet baru terbuka di tab baru — script otomatis ikut ter-copy

---

### B. Ubah Nama Divisi di _Settings

1. Di spreadsheet baru, buka sheet **_Settings**
2. Temukan baris **DIVISI** di kolom KEY
3. Klik sel di kolom **VALUE** → ubah isinya ke nama divisi baru:
   ```
   FINANCE
   ```
   > Wajib huruf kapital semua. Tidak boleh ada spasi.
4. Tekan Enter

---

### C. Hapus Sheet Absensi Lama

Sheet absensi dari spreadsheet asal ikut ter-copy dan perlu dihapus.

1. Klik kanan pada tab sheet `DEVELOPMENT_May_2026` (atau nama apapun dari divisi lama)
2. Pilih **Delete**
3. Ulangi untuk semua sheet absensi lama yang ada
4. Biarkan sheet **Master_Data** dan **_Settings** — akan dipakai ulang

---

### D. Bersihkan dan Isi Ulang Master_Data

1. Buka sheet **Master_Data**
2. Hapus semua baris data lama (baris 4 ke bawah — jangan hapus baris header 1–3)
3. Isi dengan data staf divisi baru mulai baris 4:

   | Kolom A | Kolom B | Kolom C | Kolom D |
   |---|---|---|---|
   | `FINANCE` | Nama Lengkap Staf | email@perusahaan.com | `TRUE` |
   | `FINANCE` | Nama Staf 2 | email2@perusahaan.com | `TRUE` |

   > Kolom A harus **persis sama** dengan nilai DIVISI di _Settings (huruf kapital).

---

### E. Jalankan Setup Awal

1. Klik **🔧 Admin → ⚙️ Setup Awal (pertama kali)**
2. Tunggu notifikasi ✅ — sistem akan:
   - Membuat sheet absensi bulan ini dengan nama `FINANCE_May_2026`
   - Mendaftarkan semua trigger otomatis untuk spreadsheet ini
   - Mengunci sheet Master_Data dengan akses admin yang benar

---

### F. Atur Jam dan Admin di _Settings

1. Buka sheet **_Settings**
2. Sesuaikan nilai-nilai berikut:

   | KEY | Ubah ke | Keterangan |
   |---|---|---|
   | `MASUK` | Jam masuk default, atau kosongkan | Kosong = staf isi sendiri |
   | `IST1_MULAI` | Jam istirahat mulai | Kosong = tidak ada auto-fill |
   | `IST1_SELESAI` | Jam istirahat selesai | |
   | `PULANG` | Kosongkan (umumnya) | Staf yang stamp pulang sendiri |
   | `JAM_REMINDER` | Jam reminder (contoh: `17`) | |
   | `ADMIN_EMAILS` | Email admin divisi ini | Pisah koma |

3. Setelah ADMIN_EMAILS diisi, klik **🔧 Admin → 🔑 Perbarui Akses Admin**

---

### G. Share ke Staf

1. Klik tombol **Share** (pojok kanan atas)
2. Tambahkan email semua staf divisi baru
3. Set akses ke **Editor** → klik Send
4. Beritahu staf bahwa spreadsheet sudah bisa digunakan

---

### Checklist Tambah Divisi Baru (Tanpa Kode)

```
[ ] Spreadsheet di-duplicate dari yang sudah ada
[ ] _Settings → DIVISI diubah ke nama divisi baru (huruf kapital)
[ ] Sheet absensi lama dihapus
[ ] Master_Data dibersihkan dan diisi staf baru
[ ] Setup Awal dijalankan → sheet bulan ini terbuat
[ ] _Settings → jam dan ADMIN_EMAILS disesuaikan
[ ] Perbarui Akses Admin dijalankan
[ ] Spreadsheet di-share ke staf divisi baru
[ ] Staf bisa menggunakan menu 📋 Absensi Saya
```

---

### Catatan: Kapan perlu minta bantuan Owner?

Duplicate-dan-ubah-DIVISI sudah cukup untuk semua kasus normal.
Minta bantuan Owner (deploy kode) **hanya jika**:
- Ingin mengubah format kolom atau struktur sheet
- Ingin menambah fitur baru ke sistem
- Terjadi error yang tidak bisa diselesaikan lewat menu Admin

---

## 9. Ubah Settings Operasional

> Dilakukan oleh **Admin/HR** langsung dari spreadsheet — tanpa perlu sentuh kode.

1. Buka spreadsheet divisi
2. Buka sheet **_Settings**
3. Edit nilai di kolom **VALUE** sesuai kebutuhan

| KEY | Keterangan | Format |
|---|---|---|
| `MASUK` | Jam masuk auto-fill | `HH:MM` atau kosong |
| `IST1_MULAI` | Jam istirahat 1 mulai | `HH:MM` atau kosong |
| `IST1_SELESAI` | Jam istirahat 1 selesai | `HH:MM` atau kosong |
| `IST2_MULAI` | Jam istirahat 2 mulai | `HH:MM` atau kosong |
| `IST2_SELESAI` | Jam istirahat 2 selesai | `HH:MM` atau kosong |
| `PULANG` | Jam pulang auto-fill | `HH:MM` atau kosong |
| `JAM_REMINDER` | Jam reminder staf belum isi pulang | Angka `0`–`23` |
| `SELISIH_MENIT_LOCK` | Menit setelah pulang hingga baris terkunci | Angka positif |
| `PLAN_JAM` | Pilihan shift di menu staf | `HH:MM - HH:MM, HH:MM - HH:MM` |
| `ADMIN_EMAILS` | Daftar email admin | `email1@..., email2@...` |

> Perubahan berlaku mulai append berikutnya — tidak perlu deploy ulang.

---

## 10. Melihat Rekap

### Rekap satu bulan

Buka sheet `DIVISI_Mmm_yyyy` (contoh: `DEVELOPMENT_May_2026`) — data sudah tersusun per staf per hari.

### Template rekap otomatis (rumus Excel)

1. Klik **🔧 Admin → 📋 Generate Template Rekap**
2. Masukkan nama sheet sumber, periode, dan divisi
3. Sheet rekap dibuat otomatis dengan rumus SUMIFS — **update sendiri jika data sumber berubah**

Output rekap:

| Kolom | Keterangan |
|---|---|
| TTL Regular HRS | Total jam kerja normal |
| TTL O/T 1st HRS | Total lembur tier 1 (maks 1 jam/hari) |
| TTL O/T After 1st HRS | Total lembur tier 2 |
| Bonus TTL O/T | Lembur melebihi batas 78 jam |
| Sunday/Red Day TTL HRS | Jam kerja hari Minggu/merah |
| TTL Day Reg Meal Allowance | Jumlah hari dapat uang makan |
| TTL OT Meal | Jumlah hari lembur dapat uang makan OT |
| Deduction Skill Bonus | Hari yang mengurangi skill bonus |

### Rekap rentang tanggal (lintas bulan)

1. Klik **🔧 Admin → 📊 Rekap Rentang Tanggal**
2. Masukkan tanggal mulai dan selesai
3. Masukkan nama sheet hasil
4. Sheet rekap dikalkulasi otomatis dari semua sheet yang masuk dalam rentang

### Data mentah rentang tanggal

1. Klik **🔧 Admin → 📅 Data Rentang Tanggal**
2. Masukkan rentang tanggal dan nama sheet output
3. Sheet berisi rumus QUERY — data otomatis update jika sumber berubah

---

## 11. Jika Ada Masalah

| Masalah | Solusi |
|---|---|
| Baris hari ini tidak muncul | **🔧 Admin → ➕ Append Baris Hari Ini** |
| Menu 🔧 Admin tidak muncul | Klik **📋 Absensi Saya → 🔄 Refresh Menu** |
| Admin baru tidak bisa akses | Tambah email di `_Settings → ADMIN_EMAILS` → klik **🔧 Admin → 🔑 Perbarui Akses Admin** |
| Admin tidak bisa edit Master_Data atau _Settings | Klik **🔧 Admin → 🔑 Perbarui Akses Admin** |
| Sheet bulan baru tidak terbuat otomatis | **🔧 Admin → 📅 Buat Sheet Bulan Baru** |
| Data sudah terkunci, perlu dikoreksi | Minta Owner buka proteksi baris tersebut via Data → Protected sheets and ranges |
| Jam auto-fill salah | Ubah nilai di sheet `_Settings` → berlaku mulai append esok hari |
| Staf tidak muncul di sheet absensi | Cek **Master_Data**: kolom A = nama divisi persis, kolom D = `TRUE` |
| Error saat klik menu admin | Cek apakah email terdaftar di `_Settings → ADMIN_EMAILS`; coba Refresh Menu |
| Trigger tidak jalan otomatis | **🔧 Admin → ⏰ Setup Trigger** untuk mendaftarkan ulang |

---

*Dokumen ini sesuai dengan sistem versi sheet-based (tanpa Web App).*
*Untuk perubahan konfigurasi lanjutan (jam kerja normal, zona waktu, nama instansi), hubungi Owner untuk edit `Config.js` dan deploy ulang.*
