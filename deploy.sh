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
  "DEVELOPMENT:1nH3AmfYpDLj0QoskX45YKZ72fs0q1TinpoWfhvN4OT5rc-vvg2eAX5Wk"
  "WORKER:1GSnVsll5y8tx8iExoxt1JmJ5cV5loy4JfsuklmzjV-F4qxv2zmy7iySC"
  "LINE 1:14YgLiZHIXIyf8W4GVHh_P2udL04bY1qglOr-v0eBGpLVJ335Lrf4ke1J"
  # "FINANCE:GANTI_DENGAN_SCRIPT_ID_FINANCE"
  # "MARKETING:GANTI_DENGAN_SCRIPT_ID_MARKETING"
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
