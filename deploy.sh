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
  "HR:1fusNRSgk7H5t-s8gd_oHL5dT1wt5SS65ZPBR83EuL6e30pj6xExr8dQf"
  "WEB:1cSbrUTMAnOaJqrG1Ev5qjdgMnhV1JXWopI5qA4WBVvpIwh8eFRqprv_A"
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
