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
  "IF - STAFF - HR:13rbBtFNyiQSHUYtRY19jnH5EZcJ3iOclyvSJ-7Is6KFNLX0VOsADa3S2"
  "IF - STAFF - WEB:1w_k-vX_Rlx9pN6VfIIVe5Ua553HLjyFjolovs1oRwA4M1AwGqYt9lLDQ"
  "IF - STAFF - PPIC:1a9kCvIAlKNxd5SV4HEbMVGAQFwPM0rYXzHL64zdS8O_YolOgLvEl5ACo"
  "IF - STAFF - R&D:1P5uhQ612M-XV9LBogZe9ohpzf1eiz-_2g-KVinRirBDPNi3ash6-ZA5p"
  "IF - STAMPER:1bFfYmBvZSazWKQQrQt12Mtmcu2UCtJigImm0ADh60fehfrBd_KkET5kW"
  "IF - ST5:1s_s45nl_XUVvu-BUs_HHc_XH1Tj2gKPWGFJvFljzD6TvzVZGYnvUWSKr"
  "IF - SATPAM:10wYmcIGPcdmA-7QE4IRiFE32Iagw4Dq1Q0Bg9XxJQR3kyr6JrUnj63_e"
  "IF - MECHANIC:1QipA_Nw5OwUUrK23aaiwnzj1Pv_4u-Ga4URJmwtjvyg9q7QRkSOdGVoZ"
  "IF - CS:16eGjkM4ciAgvpTsvI5JH5XOKddtrShMvyKg4Uvgum-9Bg5lqD7fOHjx3"
  "DEV:1FVp18iBxrTmIjS9MoC_f-DUGPn9odkYLmkkxu8Z9GhEcvOUQN5vb7xjH"
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
