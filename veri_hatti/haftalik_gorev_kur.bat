@echo off
rem Her pazartesi 09:17'de veri_hatti.py'yi calistiran Windows zamanlanmis gorevi kurar.
rem Kullanim: haftalik_gorev_kur.bat "C:\yol\tbb-banka-finansal-tablolar" "D:\BankaVeri" [EVDS_API_ANAHTARI]
rem Bu gorev yalniz sizin bilgisayarinizda, sizin Windows hesabinizla calisir; Claude'a baglanmaz.
setlocal
if "%~2"=="" (
  echo Kullanim:  haftalik_gorev_kur.bat "C:\yol\tbb-banka-finansal-tablolar" "D:\BankaVeri" [EVDS_ANAHTARI]
  exit /b 1
)
set DEPO=%~1
set KOK=%~2
if not "%~3"=="" ( setx EVDS_API_KEY "%~3" >nul & echo EVDS anahtari kaydedildi. )
set BURASI=%~dp0
for /f "delims=" %%p in ('where python') do (set PY=%%p& goto :bulundu)
:bulundu
if "%PY%"=="" ( echo Python bulunamadi. & exit /b 1 )
"%PY%" -m pip install --disable-pip-version-check -q playwright pymupdf openpyxl requests
"%PY%" -m playwright install chromium
schtasks /Create /F /SC WEEKLY /D MON /ST 09:17 /TN "Banka veri hatti haftalik" ^
  /TR "\"%PY%\" \"%BURASI%veri_hatti.py\" --kok \"%KOK%\" --depo \"%DEPO%\" --yayinla"
echo.
echo Gorev kuruldu: her pazartesi 09:17. Ilk calistirmayi simdi elle yapin (2025-12'den beri her seyi indirir):
echo   "%PY%" "%BURASI%veri_hatti.py" --kok "%KOK%" --depo "%DEPO%"
endlocal
