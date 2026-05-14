#!/bin/bash
# ═══════════════════════════════════════════════════════════════════════
# deploy.sh — Push codebase ke semua spreadsheet divisi sekaligus
#
# Cara pakai:
#   chmod +x deploy.sh    (sekali saja)
#   ./deploy.sh
#
# Cara dapat scriptId:
#   Buka spreadsheet divisi → Extensions → Apps Script
#   → Project Settings (⚙️) → IDs → Script ID
# ═══════════════════════════════════════════════════════════════════════

set -euo pipefail

# ── Daftar divisi: "NAMA:scriptId" ──────────────────────────────────────
# Tambah baris baru untuk setiap divisi baru
TARGETS=(
    "IF - PARWATI:1WBaTgmLvIIMCUjaqaZL4Ak2_cIBYt5x3J1SiCKVsXtjkX1bMLVChSOM0"
    "IF - ASTIN:1W8ovD-3HrsscR9Pos8ppH7_NfmP3K_94xVrxxhQUlvAe4YLa9BIe1QSs"
    "DEV:1BAIO7zfmr1bovKNCRz2KOj_WLz6r2YiolGR1hdV2bnu93_wNqW1-8tN4"
    "IF - YARN WAREHOUSE:1ozJ_04_2avBafxM40YBr5u1laCRjlXCMDyWX36kn6FITfDehWLZsEAmt"
    "IF - SELMA:1QxZhANB2wBooD-wOtHjNh5fS6ecbmzP5d-R0NB70eHeMuBdu4TbWa1qD"
    "IF - ST5:13gQY2wmk9iuTR2ViQvK5u651_yGB1aHCRcv6hL8tFYbpnIgOMpBSg6ND"
    "IF - SRI:10WwhxCoa4Q8WLFzmzQdViwk5xxKwEvXVAGtuRA-DmE9skO0Dv6JigSrE"
    "IF - RUPIASIH:1XxpBAvBVyY-sE_MMGirGlzS0jZaGPMnqux2vSsdcyoH2_nw3YctAgDIu"
    "IF - RINCE:1_XDDJG1VwkuL4bY0FJ58ao66kzCy2Ok1xZ0emcnj3CRUeCEwKg6E3yhw"
    "IF - KADEK VERA:1KuLIBUUabgOPj3BbThe_dPUXjyZ_5WfY88C5WSDxvGB8RfJZXNqzeVGd"
    "TESTING WORKER:1vMbBBnUEw7ZlAoGkaHSSCp2Xp7qBBpOix8YlH9Bh7dpRNC3jUX4Irbgr"
)

# ── Validasi ─────────────────────────────────────────────────────────────
if ! command -v clasp &> /dev/null; then
  echo "❌ clasp tidak ditemukan. Install dulu: npm install -g @google/clasp"
  exit 1
fi

if [ ! -f ".clasp.json" ]; then
  echo "❌ File .clasp.json tidak ditemukan. Jalankan dari folder project."
  exit 1
fi

# ── Backup & restore otomatis jika script crash ───────────────────────────
CLASP_BACKUP=$(cat .clasp.json)
CONFIG_BACKUP=$(cat Config.js)

restore_all() {
  echo "$CLASP_BACKUP" > .clasp.json
  echo "$CONFIG_BACKUP" > Config.js
}
trap restore_all EXIT

# ── Push ke tiap divisi ───────────────────────────────────────────────────
SUCCESS=0
FAILED=()

echo ""
echo "🚀 Push ke ${#TARGETS[@]} divisi..."
echo "════════════════════════════════════════"

for entry in "${TARGETS[@]}"; do
  NAMA="${entry%%:*}"
  SCRIPT_ID="${entry##*:}"

  # Peringatan jika scriptId belum diganti
  if [[ "$SCRIPT_ID" == GANTI_* ]]; then
    echo "⚠  [$NAMA] Dilewati — scriptId belum diisi"
    FAILED+=("$NAMA (scriptId kosong)")
    echo ""
    continue
  fi

  echo "→  [$NAMA] scriptId: ${SCRIPT_ID:0:28}…"

  # Tulis .clasp.json sementara dengan scriptId divisi ini
  sed "s|\"scriptId\": *\"[^\"]*\"|\"scriptId\": \"$SCRIPT_ID\"|" \
    <<< "$CLASP_BACKUP" > .clasp.json

  # Tulis Config.js sementara dengan DIVISI hanya untuk divisi ini
  sed "s|DIVISI *:.*\[.*\]|DIVISI        : ['$NAMA']|" \
    <<< "$CONFIG_BACKUP" > Config.js

  if clasp push --force 2>&1; then
    echo "   ✅ [$NAMA] berhasil"
    SUCCESS=$((SUCCESS + 1))
  else
    echo "   ❌ [$NAMA] GAGAL"
    FAILED+=("$NAMA")
  fi
  echo ""
done

# ── Ringkasan ─────────────────────────────────────────────────────────────
echo "════════════════════════════════════════"
echo "✅ Berhasil : $SUCCESS divisi"

if [ ${#FAILED[@]} -gt 0 ]; then
  echo "❌ Gagal    : ${FAILED[*]}"
  echo ""
  echo "Cek scriptId di bagian TARGETS dan pastikan clasp sudah login."
  exit 1
fi

echo "🎉 Semua divisi terupdate!"
