#!/usr/bin/env bash
# macOS / Linux: her pazartesi 09:17'de veri_hatti.py'yi çalıştıran cron satırını ekler.
# Kullanım: bash haftalik_gorev_kur.sh /yol/tbb-banka-finansal-tablolar /yol/BankaVeri [EVDS_API_ANAHTARI]
set -e
DEPO="${1:?Depo klasörünü verin}"
KOK="${2:?Ana veri klasörünü verin}"
BURASI="$(cd "$(dirname "$0")" && pwd)"
PY="$(command -v python3 || command -v python)"
"$PY" -m pip install -q playwright pymupdf openpyxl requests
"$PY" -m playwright install chromium
ANAHTAR="${3:-$EVDS_API_KEY}"
SATIR="17 9 * * 1 cd \"$BURASI\" && EVDS_API_KEY=\"$ANAHTAR\" \"$PY\" veri_hatti.py --kok \"$KOK\" --depo \"$DEPO\" --yayinla >> \"$KOK/gunluk/cron.log\" 2>&1"
mkdir -p "$KOK/gunluk"
( crontab -l 2>/dev/null | grep -v -e otomatik_guncelle.py -e veri_hatti.py; echo "$SATIR" ) | crontab -
echo "Kuruldu: her pazartesi 09:17.  İlk çalıştırma (elle): $PY $BURASI/veri_hatti.py --kok $KOK --depo $DEPO"
