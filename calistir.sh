#!/usr/bin/env bash
# macOS / Linux: kurulum, veri çekme ve rapor.  Çalıştırma:  bash calistir.sh
set -e
cd "$(dirname "$0")"
PY=python3; command -v python3 >/dev/null || PY=python
$PY --version
echo "[1/4] Paketler kuruluyor..."
$PY -m pip install --disable-pip-version-check -q playwright pandas openpyxl
echo "[2/4] Chromium indiriliyor..."
$PY -m playwright install chromium
echo "[3/4] TBB'den veri çekiliyor (5-10 dakika)..."
$PY tbb_cek.py "$@"
echo "[4/4] Rapor oluşturuluyor..."
$PY rapor_olustur.py --excel tbb_veri.xlsx --uzun tbb_uzun.csv
echo "Bitti: Turk_Bankacilik_Sistemi.html"
