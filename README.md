# Absent Worker

A Google Apps Script automation system for employee attendance and time tracking, built on Google Sheets. Designed for non-technical users with a familiar spreadsheet interface, real-time collaboration, and automated daily operations.

## Overview

Each division (e.g., DEVELOPMENT, WORKER) runs its own Google Sheets spreadsheet with the same codebase deployed via [clasp](https://github.com/google/clasp). Staff stamp their attendance through a custom Google Sheets menu; all calculations, reminders, and row protection are handled automatically.

## Features

- **Automated daily rows** — Rows for today's staff are appended each morning at 06:00
- **Staff time stamping** — Check in/out and break times via a custom menu
- **Formula-based calculations** — Effective hours, regular hours, and overtime calculated automatically
- **Row locking** — Rows are locked 30 minutes after clock-out to prevent edits
- **Clock-out reminders** — Email reminder sent at 17:00 to staff who haven't clocked out
- **Monthly sheet creation** — New attendance sheet auto-created on the 1st of each month
- **Role-based protection** — Staff can only edit their own rows; admins have full access
- **Payroll reporting** — Admin tools to generate SUMIFS-based payroll summaries

## Project Structure

```
absent-worker/
├── Config.js           # Global config, column mapping, auto-fill settings
├── Setup.js            # Sheet initialization, validation, and protection setup
├── Triggers.js         # Time-based triggers, menus, onEdit handler
├── Utils.js            # Date/time helpers (parseHHMM, decimalToHHMM, etc.)
├── Append.js           # Daily row appending, highlighting, grouping
├── Stamp.js            # Staff menu actions (clock in/out, breaks)
├── Lock.js             # Row locking and clock-out reminder logic
├── Rekap.js            # Reporting and payroll calculations
├── appsscript.json     # Apps Script manifest (OAuth scopes, V8 runtime)
├── .clasp.json         # clasp config (current deployment script ID)
├── deploy.sh           # Multi-spreadsheet deployment script
└── SOP.md              # Full operational guide (Indonesian)
```

## Spreadsheet Structure

Each division spreadsheet contains three sheet types:

| Sheet | Purpose | Who Can Edit |
|---|---|---|
| `Master_Data` | Staff registry (active/inactive) | Admin only |
| `DIVISI_Mmm_yyyy` | Daily attendance (e.g., `DEVELOPMENT_May_2026`) | Staff (own rows), Admin (all) |
| `_Settings` | Per-spreadsheet configuration | Admin only |

### Attendance Sheet Columns

| Column | Field | Access |
|---|---|---|
| A–D | Date, Day, Name, Email | Locked (system) |
| E–K | Status, Clock In, Break 1 Start/End, Break 2 Start/End, Clock Out | Staff (own row) |
| L–O | Effective Hours, Regular Hours, OT 1st Hour, OT After 1st | Formula (locked) |
| P–Q | Note, Sunday/Red Day | Admin only |
| R–U | Absence Notes, Plan, Late Note, Early Leave Note | Staff (own row) |

## Automated Triggers

| Time | Action |
|---|---|
| 1st of month, 05:00 | Create new month sheet |
| Daily 06:00 | Append today's rows for all active staff |
| Daily at `JAM_REMINDER` (default 17:00) | Email reminder for missing clock-outs |
| Every hour | Lock rows 30+ minutes after clock-out |

## Setup

### Prerequisites

- [Node.js](https://nodejs.org/) and `clasp` installed globally:
  ```bash
  npm install -g @google/clasp
  clasp login
  ```
- Google account with access to the target spreadsheets
- Editor access granted to the Apps Script projects

### First-Time Setup for a New Division

1. Create a new Google Spreadsheet for the division
2. Note the spreadsheet's bound Apps Script script ID (from **Extensions → Apps Script → Project Settings**)
3. Add the division to `deploy.sh`:
   ```bash
   TARGETS=(
     "DEVELOPMENT:your-script-id-here"
     "WORKER:another-script-id-here"
   )
   ```
4. Deploy the code:
   ```bash
   chmod +x deploy.sh
   ./deploy.sh
   ```
5. In the spreadsheet, run **🔧 Admin → ⚙️ Setup Awal (pertama kali)** to:
   - Initialize `Master_Data` protection
   - Register all time-based triggers
   - Create the current month sheet and append today's rows
6. Run **🔧 Admin → ⚙️ Buat/Reset Sheet Settings** to create the `_Settings` sheet
7. Fill in `_Settings` with the appropriate values (see [Configuration](#configuration))
8. Share the spreadsheet with all staff as Editor

### Deploying Updates

After editing any `.js` file, push to all divisions at once:

```bash
./deploy.sh
```

The script temporarily patches `.clasp.json` and `Config.js` for each division, pushes, then restores the originals.

## Configuration

### `_Settings` Sheet (per spreadsheet, highest priority)

| Key | Description | Default |
|---|---|---|
| `DIVISI` | Division name (uppercase) | From `Config.js` |
| `MASUK` | Auto-fill clock-in time (HH:MM) | `''` |
| `IST1_MULAI` | Auto-fill break 1 start | `''` |
| `IST1_SELESAI` | Auto-fill break 1 end | `''` |
| `IST2_MULAI` | Auto-fill break 2 start (optional) | `''` |
| `IST2_SELESAI` | Auto-fill break 2 end (optional) | `''` |
| `PULANG` | Auto-fill clock-out time | `''` |
| `JAM_REMINDER` | Hour for clock-out reminder (0–23) | `17` |
| `SELISIH_MENIT_LOCK` | Minutes after clock-out before row locks | `30` |
| `PLAN_JAM` | Comma-separated shift options | From `Config.js` |
| `ADMIN_EMAILS` | Comma-separated admin email list | From `Config.js` |

### `Config.js` (code-level defaults)

Static defaults applied if not overridden by `_Settings`:

```javascript
CONFIG = {
  NAMA_INSTANSI: 'PT InFashion',
  TIMEZONE: 'Asia/Makassar',          // UTC+8 (WITA)
  JAM_REMINDER: 17,
  SELISIH_MENIT_LOCK: 30,
  ADMIN_EMAILS: ['webadmin@wooden-ships.com'],
  DAYS_HOUR: { REGULAR_DAYS: 7, SATURDAY: 5 },
  AUTO_ABSENSI: {
    'DEVELOPMENT': { masuk: '07:00', ist1Mulai: '12:00', ist1Selesai: '13:00', ... },
    'WORKER':      { masuk: '08:00', ist1Mulai: '12:00', ist1Selesai: '13:00', ... }
  }
}
```

## Usage

### Staff Menu: `📋 Absensi Saya`

| Item | Action |
|---|---|
| 📍 Ke baris saya hari ini | Navigate to today's row |
| ✅ Stamp MASUK | Clock in |
| ☕ Stamp ISTIRAHAT 1 MULAI | Start break 1 |
| ▶ Stamp ISTIRAHAT 1 SELESAI | End break 1 |
| ☕ Stamp ISTIRAHAT 2 MULAI | Start break 2 (optional) |
| ▶ Stamp ISTIRAHAT 2 SELESAI | End break 2 (optional) |
| 🏁 Stamp PULANG | Clock out |
| 📊 Rekap absensi saya | Personal monthly summary |
| 🔄 Refresh Menu | Reload menu |

### Admin Menu: `🔧 Admin`

| Item | Action |
|---|---|
| 📅 Buat Sheet Bulan Baru | Create next month's sheet |
| ➕ Append Baris Hari Ini | Manually add today's rows |
| 📋 Generate Template Rekap | Create payroll summary template |
| 📅 Data Rentang Tanggal | Raw data extract for date range |
| 📊 Rekap Rentang Tanggal | Summary report for date range |
| ⚠️ Cek Belum Isi Pulang | List staff missing clock-out |
| 🔒 Lock Baris Sudah Pulang | Manually lock completed rows |
| ⚙️ Setup Awal (pertama kali) | One-time initialization |
| ⏰ Setup Trigger | Register automated triggers |
| 🛡️ Setup Proteksi Master | Lock Master_Data sheet |
| 🔑 Perbarui Akses Admin | Grant admin access across all sheets |
| ⚙️ Buat/Reset Sheet Settings | Create/reset `_Settings` sheet |

## Documentation

See [SOP.md](SOP.md) for the complete operational guide (Indonesian), covering:
- Staff and admin daily routines
- Adding new staff or divisions
- Troubleshooting common issues
- Payroll summary generation
