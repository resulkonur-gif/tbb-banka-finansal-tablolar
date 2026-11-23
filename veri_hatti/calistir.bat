@echo off
chcp 65001 >nul
setlocal
rem ===== Banka veri hattı: tek tık çalıştırma =====
rem Klasörleri değiştirmek isterseniz yalnız bu iki satırı düzenleyin (boş bırakırsanız Belgeler altında kurulur).
set KOK=
set DEPO=
rem EVDS anahtarı (isteğe bağlı): evds2.tcmb.gov.tr > Profil > API Anahtarı
if "%EVDS_API_KEY%"=="" set EVDS_API_KEY=
rem ===============================================
if "%KOK%"=="" set "KOK=%USERPROFILE%\Documents\BankaVeri"
rem Bu klasör deponun içindeyse (tbb-banka-finansal-tablolar\veri_hatti) depo bir üst klasördür
if "%DEPO%"=="" if exist "%~dp0..\tbb_cek.py" for %%I in ("%~dp0..") do set "DEPO=%%~fI"
if "%DEPO%"=="" set "DEPO=%USERPROFILE%\Documents\tbb-banka-finansal-tablolar"
cd /d "%~dp0"

where python >nul 2>nul || (echo [HATA] Python bulunamadı. python.org'dan kurun ve "Add python.exe to PATH" kutusunu işaretleyin. & pause & exit /b 1)

if not exist "%DEPO%\tbb_cek.py" (
  echo Depo bulunamadı: %DEPO%
  where git >nul 2>nul || (echo [HATA] Git yok. git-scm.com'dan kurun ya da DEPO satırına deponun yerini yazın. & pause & exit /b 1)
  echo Depo indiriliyor...
  git clone https://github.com/resulkonur-gif/tbb-banka-finansal-tablolar "%DEPO%" || (echo [HATA] Depo indirilemedi. & pause & exit /b 1)
)
if not exist "%KOK%" mkdir "%KOK%" || (echo [HATA] Ana klasör oluşturulamadı: %KOK% & pause & exit /b 1)

echo Ana klasör : %KOK%
echo Depo       : %DEPO%
echo.
python -m pip install --disable-pip-version-check -q playwright pymupdf openpyxl requests
python -m playwright install chromium >nul
python veri_hatti.py --kok "%KOK%" --depo "%DEPO%" %*
set SONUC=%ERRORLEVEL%
echo.
if exist "%KOK%\7_site\site\index.html" (
  echo Site : %KOK%\7_site\site\index.html
  start "" "%KOK%\7_site\site\index.html"
) else (
  echo [UYARI] Site oluşmadı. Ayrıntı: %KOK%\gunluk klasöründeki son kayıt.
)
if exist "%KOK%\OZET.md" echo Özet : %KOK%\OZET.md
if not "%SONUC%"=="0" echo [HATA] veri_hatti.py hata koduyla bitti: %SONUC%
pause
