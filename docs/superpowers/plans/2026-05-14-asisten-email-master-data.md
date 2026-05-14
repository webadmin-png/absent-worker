# Email Asisten di Master_Data — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Allow worker rows in the attendance sheet to be edited by both the worker's email AND an optional assistant email defined in Master_Data column E.

**Architecture:** Master_Data gains an optional column E ("Email Asisten"). `proteksiBarisBaru()` in `Setup.js` reads column E in addition to A–D and adds the assistant as a row-range editor. No other code changes are required — the attendance sheet schema, formulas, and Stamp/Lock logic stay the same.

**Tech Stack:** Google Apps Script (clasp deploy). No local test runner — verification is via Apps Script Logger output and manual checks in Google Sheets.

**Spec reference:** [`docs/superpowers/specs/2026-05-14-asisten-email-master-data-design.md`](../specs/2026-05-14-asisten-email-master-data-design.md)

**Plan deviation from spec:** Spec proposed editing `Append.js` to add an `emailAsisten` field to the staf object. Plan-time self-review found this would be dead code — the field is never read (sheet schema unchanged, and `Setup.js` reads Master_Data independently). YAGNI: dropped from plan. Spec remains the architectural source of truth; if future requirements need the field, add then.

---

## File Structure

- Modify: `Setup.js` — extend Master_Data read range in `proteksiBarisBaru()` from `A4:D200` to `A4:E200`; add asisten `prot.addEditor` block if present.
- Modify: `Config.js` — add a new comment block documenting the Master_Data schema (no executable code change).

No new files. No file deletions. `Append.js`, `Stamp.js`, `Lock.js`, `Utils.js`, `Rekap.js`, `Triggers.js` are not touched.

---

## Pre-flight

- [ ] **Step 0a: Confirm Master_Data structure**

Open one of the active Google Spreadsheets and confirm:
- Row 3 contains headers.
- Columns A–D are: Divisi, Nama, Email, Aktif.
- Column E is currently unused (or already named "Email Asisten" with empty cells — also fine).

If column E already contains unrelated data, stop and surface to user.

- [ ] **Step 0b: Confirm clean working tree for files this plan touches**

Run: `git status -- Setup.js Config.js`
Expected: clean (or only unrelated edits the user is in the middle of — note them and proceed only on the lines this plan changes; do not include unrelated lines in the commits).

---

## Task 1: Add asisten as editor in `proteksiBarisBaru()`

**Files:**
- Modify: `Setup.js:407` (range string)
- Modify: `Setup.js:462-468` (insert asisten editor block after worker editor block)

- [ ] **Step 1.1: Extend the master read range**

In `Setup.js`, change line 407 from:

```js
  const masterData = master.getRange('A4:D200').getValues()
```

to:

```js
  const masterData = master.getRange('A4:E200').getValues()
```

The `.filter()` chain on the following lines still uses `r[0]` and `r[3]` — those column references remain correct.

- [ ] **Step 1.2: Add asisten editor block after worker editor**

Find this existing block (around lines 462-468):

```js
    try {
      prot.addEditor(email);
      berhasil++;
      Logger.log('✓ Proteksi: ' + nama + ' baris ' + baris[0] + '–' + baris[baris.length - 1]);
    } catch(e) {
      Logger.log('⚠ Gagal tambah editor ' + email + ': ' + e.message);
    }
```

Immediately after this `try/catch` (still inside the `for (const k of masterData)` loop, before its closing `}`), insert:

```js

    // Email asisten (kolom E Master_Data) — opsional, skip kalau kosong
    const asisten = String(k[4] || '').trim();
    if (asisten) {
      try {
        prot.addEditor(asisten);
        Logger.log('✓ Proteksi asisten: ' + nama + ' ← ' + asisten);
      } catch(err) {
        Logger.log('⚠ Gagal tambah asisten ' + asisten + ': ' + err.message);
      }
    }
```

The leading blank line is intentional — visually separates worker block from asisten block.

- [ ] **Step 1.3: Read back the modified region to verify**

Read `Setup.js` lines 405-485. Confirm:
- The `for (const k of masterData)` loop boundary is intact.
- The new block sits between the worker `addEditor` try/catch and the loop's closing `}`.
- Variables `k`, `nama`, `prot` are in scope inside the new block.
- Step 5 (`Proteksi L:O`) and Step 6 (`Proteksi P:Q`) below the loop remain unchanged.

- [ ] **Step 1.4: Commit Task 1**

```bash
git add Setup.js
git commit -m "$(cat <<'EOF'
feat(setup): tambah email asisten sebagai editor proteksi baris

proteksiBarisBaru() sekarang membaca kolom E Master_Data dan
menambahkan email asisten sebagai editor range E:K worker.
Kolom E opsional — guard if(asisten) skip kalau kosong.
addEditor di-wrap try/catch agar email invalid tidak gagalkan deploy.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Document Master_Data schema in `Config.js`

**Files:**
- Modify: `Config.js` — insert comment block before line 158 (the existing `// ── Mapping kolom sheet (1-indexed) ───` header).

- [ ] **Step 2.1: Insert the schema comment**

In `Config.js`, find this existing line:

```js
// ── Mapping kolom sheet (1-indexed) ───────────────────────────────────
```

Immediately before it, insert:

```js
// ── Skema Master_Data ─────────────────────────────────────────────────
// Sheet "Master_Data" — daftar staf yang dibaca oleh appendHariIni() dan
// proteksiBarisBaru(). Header di baris 3, data mulai baris 4.
//
// A=1  Divisi          → nama divisi (HURUF KAPITAL, cocok dengan CONFIG.DIVISI)
// B=2  Nama            → nama lengkap staf
// C=3  Email           → email Google Account staf (jadi editor barisnya)
// D=4  Aktif           → "TRUE" atau "FALSE" — hanya TRUE yang di-append
// E=5  Email Asisten   → opsional. Email tambahan yang boleh edit baris
//                        worker ini (kolom E:K). Kosongkan kalau tidak ada.
//                        Asisten TIDAK pakai menu Stamp — edit manual via sel.

```

(Note the trailing blank line — keeps separation from the next comment header.)

- [ ] **Step 2.2: Verify Config.js still parses**

Read `Config.js` lines 155-210 to confirm:
- The new comment block precedes `// ── Mapping kolom sheet`.
- The `const TOTAL_COL = 21;` and all `const COL_*` declarations below are untouched.

- [ ] **Step 2.3: Commit Task 2**

```bash
git add Config.js
git commit -m "$(cat <<'EOF'
docs(config): dokumentasikan skema Master_Data termasuk kolom E asisten

Tambah comment block yang menjelaskan kolom A–E Master_Data —
sebelumnya skema implicit, hanya di dalam kode.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: Manual verification in Google Sheets

This task is run by the developer or HRD after code is deployed. Surface it to the user as a checklist; the implementing agent should NOT auto-deploy.

- [ ] **Step 3.1: Deploy via clasp**

Run: `./deploy.sh` (deploys to all configured TARGETS)
Expected: clean `clasp push` for each target without errors.

- [ ] **Step 3.2: Add asisten header in Master_Data**

For each active spreadsheet:
- Sheet `Master_Data`, cell E3: type `Email Asisten` (or preferred label).
- Optional: match formatting of A3–D3 (font weight, background) so it looks like part of the header.

- [ ] **Step 3.3: Fill at least one asisten email**

Pick one active worker row. In column E, enter a Google Account email that should also be able to edit that worker's daily absence rows.

For initial sanity check, pick an email you control so you can verify access.

- [ ] **Step 3.4: Trigger `appendHariIni` for a fresh test**

The cleanest test: wait for tomorrow's 06:00 trigger, OR run `appendHariIni()` manually in the Apps Script editor on a fresh test spreadsheet.

Expected Logger output for any worker with an asisten in column E:
```
✓ Proteksi: <nama worker> baris <N>–<N>
✓ Proteksi asisten: <nama worker> ← <email asisten>
```

For workers WITHOUT asisten (column E empty), only the first line appears — no error.

- [ ] **Step 3.5: Verify asisten access in the sheet**

Sign in as the asisten user (or share the sheet with that account). Confirm:
- Asisten can edit cells E–K (Status through Pulang) for the assigned worker's row of the day.
- Asisten CANNOT edit other workers' rows — should get a Google Sheets permission warning.
- Worker themselves can still edit their own row (no regression).
- Asisten clicking the menu `Stamp Masuk` etc. gets a "Email tidak terdaftar" error — this is by design.

- [ ] **Step 3.6: Edge case — worker without asisten**

Pick a worker whose column E is empty. After `appendHariIni()` runs, confirm:
- No Logger error.
- Worker can still edit their own row.
- No additional editor appears in the protection panel for that row.

- [ ] **Step 3.7: Edge case — invalid email**

Optional but recommended. Put a non-Google email (e.g. `nonexistent_xyz@invalid.zzz`) in someone's column E. Run `appendHariIni()`. Expected:
- Logger logs `⚠ Gagal tambah asisten nonexistent_xyz@invalid.zzz: …`
- The worker's protection still succeeds — no cascade failure.

---

## Self-Review Notes (already addressed inline)

- **Spec coverage**:
  - Master_Data column E added — Task 2 (docs) + Step 3.2 (HRD adds header).
  - `proteksiBarisBaru` reads column E and adds asisten editor — Task 1.
  - `addEditor` failure caught with logging — Task 1, Step 1.2.
  - Edge cases (empty asisten, invalid email, asisten using Stamp menu) — verified in Step 3.5–3.7.
  - No migration helper — explicitly out of scope per spec.
- **Spec deviation justified above**: Append.js change dropped (would be dead code).
- **No placeholders**: All code is concrete.
- **Type consistency**: `k[4]` is consistently the asisten email reference (Setup.js is the only file that touches it).
