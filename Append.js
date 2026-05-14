// ═══════════════════════════════════════════════════════════════════════
// APPEND.JS — Penambahan baris harian + tampilan sheet
//
// Berisi:
//   appendHariIni()    — inti fungsi harian: tambah baris staf ke sheet divisi
//   highlightHariIni() — warnai baris hari ini (kuning) dan lewat (abu)
//   groupByToday()     — collapse baris hari lama, buka hari ini
// ═══════════════════════════════════════════════════════════════════════

// ── appendHariIni — Core function harian ──────────────────────────────
// Dipanggil otomatis pukul 06:00 via trigger, atau manual oleh HRD.
// Untuk setiap divisi:
//   1. Ambil daftar staf aktif dari Master_Data
//   2. Cek duplikat (skip jika hari ini sudah ada)
//   3. Append baris kosong per staf dengan formula L–O
//   4. Pasang validasi dan proteksi baris baru
function appendHariIni() {
  _loadSettings();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  if (!master) {
    Logger.log('❌ Master_Data tidak ditemukan.');
    return;
  }

  const today    = getToday();
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const namaHari = Utilities.formatDate(today, CONFIG.TIMEZONE, 'EEEE');

  // Ambil semua staf aktif
  const masterData = master.getRange('A4:D200').getValues()
    .filter(r =>
      r[0] !== '' && r[1] !== '' &&
      String(r[3]).trim().toUpperCase() === 'TRUE'
    );

  if (masterData.length === 0) {
    Logger.log('Tidak ada staf aktif di Master_Data.');
    return;
  }

  // Kelompokkan per divisi (case-insensitive)
  const stafPerDivisi = {};
  for (const k of masterData) {
    const div = String(k[0]).trim().toUpperCase();
    if (!stafPerDivisi[div]) stafPerDivisi[div] = [];
    stafPerDivisi[div].push({
      nama : String(k[1]).trim(),
      email: String(k[2]).trim(),
    });
  }

  const hasil = [];

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) {
      hasil.push('⚠ ' + divisi + ': sheet tidak ditemukan');
      continue;
    }

    const staf = stafPerDivisi[divisi.toUpperCase()] || [];
    if (staf.length === 0) {
      hasil.push('⚠ ' + divisi + ': tidak ada staf aktif');
      continue;
    }

    // Hapus proteksi hari lalu sebelum append — hemat kuota protect Google Sheets
    const hapus = _bersihkanProteksiLama(sheet, today);
    if (hapus > 0) Logger.log('🗑 ' + divisi + ': hapus ' + hapus + ' proteksi hari lalu');

    // Cegah duplikat — skip jika hari ini sudah ada di sheet
    const existingData = sheet.getLastRow() > 3
      ? sheet.getRange(4, 1, sheet.getLastRow() - 3, 1).getValues()
      : [];

    const sudahAda = existingData.some(r => {
      const tgl = r[0];
      return tgl instanceof Date && isSameDate(tgl, today);
    });

    if (sudahAda) {
      hasil.push('⚠ ' + divisi + ': hari ini sudah ada — skip');
      continue;
    }

    // Cek apakah divisi ini menggunakan auto-absensi (jam diisi otomatis)
    const auto = (CONFIG.AUTO_ABSENSI || {})[divisi.toUpperCase()] || null;

    // Konversi "HH:mm" → time serial (fraction hari) agar formula L bisa menghitung
    // Contoh: "07:00" → 7/24 = 0.29166...
    function toSerial(hhmm) {
      if (!hhmm) return '';
      const m = String(hhmm).match(/^(\d{1,2}):(\d{2})$/);
      if (!m) return '';
      return (parseInt(m[1]) * 60 + parseInt(m[2])) / 1440;
    }

    // Siapkan baris baru
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

    const insertAt = sheet.getLastRow() + 1;
    sheet.getRange(insertAt, 1, newRows.length, TOTAL_COL).setValues(newRows);

    // Format tanggal kolom A
    sheet.getRange(insertAt, 1, newRows.length, 1).setNumberFormat('DD/MM/YYYY');

    // Format HH:mm untuk kolom jam yang diisi otomatis
    if (auto) {
      [COL_MASUK, COL_IST1_M, COL_IST1_S, COL_IST2_M, COL_IST2_S, COL_PULANG]
        .forEach(col => sheet.getRange(insertAt, col, newRows.length, 1)
          .setNumberFormat('HH:mm'));
    }

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

    // Border
    sheet.getRange(insertAt, 1, newRows.length, TOTAL_COL)
      .setBorder(true,true,true,true,true,true,
        '#B0D9C8', SpreadsheetApp.BorderStyle.SOLID);

    // Pasang formula per baris
    _pasangFormulaBaris(sheet, insertAt, newRows.length);

    // Format kolom jam sebagai [h]:mm
    sheet.getRange(insertAt, COL_EFEKTIF,     newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_REGULAR_JAM, newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT1,         newRows.length, 1).setNumberFormat('[h]:mm');
    sheet.getRange(insertAt, COL_OT2,         newRows.length, 1).setNumberFormat('[h]:mm');

    setupValidasiBaris(sheet, insertAt, newRows.length);
    proteksiBarisBaru(sheet, divisi, insertAt, newRows.length);

    hasil.push('✓ ' + divisi + ' (' + sheet.getName() + '): ' +
               staf.length + ' staf — ' + todayStr);
    Logger.log('Append selesai: ' + divisi + ' — ' + todayStr);
  }

  groupByToday();
  highlightHariIni();

  Logger.log('appendHariIni selesai: ' + todayStr + '\n' + hasil.join('\n'));

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Append hari ini selesai!\n' +
      todayStr + ' (' + namaHari + ')\n\n' +
      hasil.join('\n')
    );
  } catch(e) {
    // Dipanggil dari trigger — tidak ada UI
  }
}

// ── highlightHariIni — Warnai baris berdasarkan tanggal ───────────────
// Kuning = hari ini,  Abu = sudah lewat
function highlightHariIni() {
  const today = getToday();

  for (const divisi of CONFIG.DIVISI) {
    const sheet = getSheetAktifDivisi(divisi);
    if (!sheet) continue;

    const data    = sheet.getDataRange().getValues();
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    for (let r = 4; r <= lastRow; r++) {
      const val = data[r - 1][0];
      if (!val || !(val instanceof Date)) continue;

      if (isSameDate(val, today)) {
        sheet.getRange(r, 1,  1, 8).setBackground('#FFF9C4');
        sheet.getRange(r, 5,  1, 7).setBackground('#FFF9C4');
        sheet.getRange(r, 12, 1, 4).setBackground('#FFF9C4');
        sheet.getRange(r, 16, 1, 6).setBackground('#FFF9C4');
      } else if (isPast(val, today)) {
        sheet.getRange(r, 1,  1, 8).setBackground('#F1EFE8');
        sheet.getRange(r, 5,  1, 7).setBackground('#F8F8F8');
        sheet.getRange(r, 12, 1, 4).setBackground('#EEEDFE');
        sheet.getRange(r, 16, 1, 6).setBackground('#F8F8F8');
      }
    }
  }
  Logger.log('Highlight selesai.');
}

// ── groupByToday — Collapse baris lama, buka baris hari ini ──────────
// TargetSheet opsional — jika null, proses semua sheet divisi aktif
function groupByToday(targetSheet) {
  const today  = getToday();
  const sheets = targetSheet
    ? [targetSheet]
    : CONFIG.DIVISI.map(d => getSheetAktifDivisi(d)).filter(s => s);

  for (const sheet of sheets) {
    const data    = sheet.getDataRange().getValues();
    const lastCol = sheet.getLastColumn();

    // Reset grouping lama
    try { sheet.expandAllRowGroups(); } catch(e) {}
    try {
      const rng = sheet.getDataRange();
      for (let i = 0; i < 3; i++) {
        try { rng.shiftRowGroupDepth(-1); } catch(e) { break; }
      }
    } catch(e) {}

    // Temukan baris pertama hari ini
    let firstRowToday = -1;
    for (let i = 3; i < data.length; i++) {
      const val = data[i][0];
      if (!val || !(val instanceof Date)) continue;
      if (isSameDate(val, today)) { firstRowToday = i + 1; break; }
    }

    const groupEnd = firstRowToday === -1
      ? sheet.getLastRow()
      : firstRowToday - 1;

    if (groupEnd >= 4) {
      const groupStart  = 4;
      const groupLength = groupEnd - groupStart + 1;
      sheet.getRange(groupStart, 1, groupLength, lastCol).shiftRowGroupDepth(1);
    }

    try { sheet.collapseAllRowGroups(); } catch(e) {}

    Logger.log(
      sheet.getName() + ': group baris 4–' + groupEnd +
      (firstRowToday > -1 ? ', hari ini mulai baris ' + firstRowToday : ' (semua)')
    );
  }
}

// ── _pasangFormulaBaris — Private: set formula L, M, N, O ─────────────
// Dipanggil oleh appendHariIni() dan generateFullMonth()
function _pasangFormulaBaris(sheet, startRow, numRows) {
  for (let r = startRow; r < startRow + numRows; r++) {
    const a=`A${r}`, b=`B${r}`, e=`E${r}`, f=`F${r}`,
          g=`G${r}`, h=`H${r}`, i=`I${r}`, j=`J${r}`,
          k=`K${r}`, l=`L${r}`;

    // L: Jam Efektif (fraction hari)
    sheet.getRange(r, COL_EFEKTIF).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF(AND(${g}<>"",${h}<>""),` +
        `IF(AND(${i}<>"",${j}<>""),` +
          `${k}-${f}-(${h}-${g})-(${j}-${i}),` +
          `${k}-${f}-(${h}-${g})),` +
      `IF(AND(${i}<>"",${j}<>""),` +
          `${k}-${f}-(${j}-${i}),` +
          `${k}-${f}))))`
    );

    // M: Regular Hours — Red Day langsung dapat 7 jam (hari libur dibayar penuh)
    sheet.getRange(r, COL_REGULAR_JAM).setFormula(
      `=IF(${e}="Red Day",${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
      `IF(${b}="Saturday",` +
        `IF(${l}>=${CONFIG.DAYS_HOUR.SATURDAY}/24,` +
          `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l}),` +
      `IF(${l}>=${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,` +
        `${CONFIG.DAYS_HOUR.REGULAR_DAYS}/24,${l})))`
    );

    // N: OT 1 (maks 1 jam di atas regular)
    sheet.getRange(r, COL_OT1).setFormula(
      `=IF(${e}<>"Hadir",0,IF(OR(${f}="",${k}=""),0,IF((${k}-${f}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))<=IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24,0,MIN(1/24,(${k}-${f}-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))-IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})/24))))`
    );

    // O: OT 2 (di atas OT 1)
    sheet.getRange(r, COL_OT2).setFormula(
      `=IF(${e}<>"Hadir",0,` +
      `IF(OR(${f}="",${k}=""),0,` +
      `IF((${k}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0))` +
        `<=(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24,` +
      `0,` +
      `${k}-${f}` +
        `-IF(AND(${g}<>"",${h}<>""),${h}-${g},0)` +
        `-IF(AND(${i}<>"",${j}<>""),${j}-${i},0)` +
        `-(IF(WEEKDAY(${a},2)=6,${CONFIG.DAYS_HOUR.SATURDAY},${CONFIG.DAYS_HOUR.REGULAR_DAYS})+1)/24)))`
    );
  }
}

// ── _bersihkanProteksiLama — Konsolidasi proteksi hari lalu ──────────
// Dipanggil di awal appendHariIni() sebelum menambah baris baru.
// Strategi:
//   1. Temukan batas baris hari lalu (baris 4 s/d firstTodayRow-1)
//   2. Hapus semua range protection individual di area hari lalu
//   3. Buat SATU proteksi konsolidasi yang mengcover seluruh area hari lalu
//      → Baris hari lalu tetap terkunci (owner + admin only), hemat kuota protect
// Header (baris 1–3) dan proteksi hari ini tidak disentuh.
// Return: jumlah proteksi individual yang dihapus (untuk logging).
function _bersihkanProteksiLama(sheet, today) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return 0;

  // Temukan baris pertama yang merupakan hari ini
  const dates = sheet.getRange(4, COL_TANGGAL, lastRow - 3, 1).getValues();
  let firstTodayRow = lastRow + 1; // default: belum ada baris hari ini
  for (let i = 0; i < dates.length; i++) {
    if (dates[i][0] instanceof Date && isSameDate(dates[i][0], today)) {
      firstTodayRow = i + 4; // 1-indexed
      break;
    }
  }

  const lastPastRow = firstTodayRow - 1;
  if (lastPastRow < 4) return 0; // tidak ada baris hari lalu sama sekali

  // Hapus semua range protection yang seluruhnya berada di area hari lalu
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let removed = 0;
  for (const prot of protections) {
    const range  = prot.getRange();
    const pStart = range.getRow();
    const pEnd   = range.getLastRow();

    if (pEnd <= 3) continue;                    // lewati header
    if (pStart >= 4 && pEnd <= lastPastRow) {   // seluruhnya di area hari lalu
      prot.remove();
      removed++;
    }
  }

  // Ganti dengan SATU proteksi konsolidasi (owner + admin only)
  // → baris hari lalu tetap read-only meski grup dibuka
  const owner     = Session.getEffectiveUser();
  const pastRange = sheet.getRange(4, 1, lastPastRow - 3, TOTAL_COL);
  const newProt   = pastRange.protect();
  newProt.setDescription('Hari lalu — terkunci (konsolidasi)');
  newProt.setWarningOnly(false);
  newProt.removeEditors(newProt.getEditors());
  newProt.addEditor(owner);
  for (const adminEmail of CONFIG.ADMIN_EMAILS) {
    try { newProt.addEditor(adminEmail); } catch(e) {}
  }

  return removed;
}
