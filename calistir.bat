@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
echo ============================================================
echo  TBB finansal tablolar: kurulum, veri cekme ve rapor
echo ============================================================
set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY (
  echo.
  echo [HATA] Python bulunamadi. python.org'dan Python kurun ve
  echo "Add python.exe to PATH" kutusunu isaretleyin.
  echo PYTHON_YOK> durum.txt
  pause
  exit /b 1
)
%PY% --version
echo.
echo [1/4] Paketler kuruluyor (playwright, pandas, openpyxl)...
%PY% -m pip install --disable-pip-version-check -q playwright pandas openpyxl
if errorlevel 1 ( echo PIP_HATA> durum.txt & echo [HATA] pip kurulumu basarisiz. & pause & exit /b 1 )
echo [2/4] Chromium indiriliyor (ilk seferde birkac dakika)...
%PY% -m playwright install chromium
if errorlevel 1 ( echo CHROMIUM_HATA> durum.txt & echo [HATA] Chromium kurulamadi. & pause & exit /b 1 )
echo [3/4] TBB'den veri cekiliyor (5-10 dakika)...
%PY% tbb_cek.py %*
if errorlevel 1 ( echo CEKME_HATA> durum.txt & echo [HATA] Veri cekilemedi. & pause & exit /b 1 )
echo [4/4] Rapor olusturuluyor...
%PY% rapor_olustur.py --excel tbb_veri.xlsx --uzun tbb_uzun.csv
if errorlevel 1 ( echo RAPOR_HATA> durum.txt & echo [HATA] Rapor olusturulamadi. & pause & exit /b 1 )
echo TAMAM> durum.txt
echo.
echo Bitti. Rapor aciliyor: Turk_Bankacilik_Sistemi.html
start "" "Turk_Bankacilik_Sistemi.html"
pause
