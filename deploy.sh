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
  "IF - STAFF - HR:1fusNRSgk7H5t-s8gd_oHL5dT1wt5SS65ZPBR83EuL6e30pj6xExr8dQf"
  "IF - STAFF - WEB:1cSbrUTMAnOaJqrG1Ev5qjdgMnhV1JXWopI5qA4WBVvpIwh8eFRqprv_A"
  "IF - STAFF - PPIC:1R1C65Rp76CL5By3KaiY-eK_lzfd_sXWlsHn9t6ZTtjS1dHd50JzkZk11"
  "IF - STAFF - R&D:1mcPn439jC5Fv9MlnhRlJKc5X0XG5QcwZZMaQlD2B7AcHop55TEO_K9Ra"
  "IF - STAMPER:1p6RQYf47xy3axaIOw3GPZ0weX2LieQQPa1aR8bT7424GFoQHg4INX-zO"
  "IF - ST5:1KjdVt9XZrr5B-YhJ73xHkxfsPdrbxmV1VI7bZFu07H2WHO5snE29wj7z"
  "IF - SATPAM:17pqHm0r5uo0ADAhXDbLQs4YrVRFkDlgzHozfrMEPzIuCBQlhR3yuSXh6"
  "IF - MECHANIC:1-1db2_bQz86se8f2Y0IDMSYxvnuGvEYrUpZ_kHAEP0ntsv12tdu9t8rT"
  "IF - CS:1EKASbK1OsM1Rs3HaK9Cz0aqb57k8PAjcbh8nMrHwBj9lqTBS2_nLgYG-"
  "DEV:1lfIZETbtj2BL1v12vWFSpMroyabTEeqbEQQnqKiW7Uae5FYVW9MRoSP1"
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
  # NOTE: escape karakter spesial sed di replacement: & = whole match,
  # | = delimiter, \ = escape. Tanpa ini, nama divisi seperti "R&D" rusak.
  ESCAPED_NAMA=$(printf '%s' "$NAMA" | sed 's/[\\&|]/\\&/g')
  sed "s|DIVISI *:.*\[.*\]|DIVISI        : ['$ESCAPED_NAMA']|" \
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
