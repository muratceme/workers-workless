@echo off
rem Gerisi Bende - Windows baslaticisi: bu dosyaya cift tiklayin.
chcp 65001 >nul
cd /d "%~dp0"
title Gerisi Bende

set "PY="
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
  if not errorlevel 1 set "PY=py -3"
)
if not defined PY (
  python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
  if not errorlevel 1 set "PY=python"
)
if defined PY goto calistir

echo.
echo Bu paketi calistirmak icin Python 3.10 veya ustu gerekiyor ve bilgisayarinizda bulunamadi.
where winget >nul 2>nul
if not %errorlevel%==0 goto elle
echo.
choice /c EH /m "Python'u Microsoft'un uygulama yukleyicisi (winget) ile simdi kurayim mi? (E = evet, H = hayir)"
if errorlevel 2 goto elle
winget install -e --id Python.Python.3.12 --scope user
echo.
echo Kurulum tamamlandiysa bu pencereyi kapatip BASLAT.bat dosyasina yeniden cift tiklayin.
pause
exit /b

:elle
echo.
echo Python'u https://www.python.org/downloads/ adresinden indirip kurun.
echo Kurulumun ilk ekraninda "Add python.exe to PATH" kutusunu isaretlemeyi unutmayin.
start "" "https://www.python.org/downloads/"
pause
exit /b

:calistir
%PY% "%~dp0_baslat.py"
