#!/usr/bin/env bash
# Banka veri hattı: tek komut. Aşağıdaki üç satırı bir kez düzenleyin.
BURASI="$(cd "$(dirname "$0")" && pwd)"
if [ -f "$BURASI/../tbb_cek.py" ]; then DEPO="${DEPO:-$(cd "$BURASI/.." && pwd)}"; else DEPO="${DEPO:-$HOME/tbb-banka-finansal-tablolar}"; fi
KOK="${KOK:-$HOME/BankaVeri}"
export EVDS_API_KEY="${EVDS_API_KEY:-}"
set -e
cd "$(dirname "$0")"
PY="$(command -v python3 || command -v python)"
"$PY" -m pip install -q playwright pymupdf openpyxl requests
"$PY" -m playwright install chromium >/dev/null
"$PY" veri_hatti.py --kok "$KOK" --depo "$DEPO" "$@"
echo "Özet: $KOK/OZET.md   Site: $KOK/7_site/site/index.html"
