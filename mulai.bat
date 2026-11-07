@echo off
cd /d %~dp0
if not exist node_modules (
  echo Memasang kebutuhan, tunggu sebentar...
  call npm install
)
echo.
echo Aplikasi berjalan di http://localhost:3000
echo Biarkan jendela ini tetap terbuka. Tekan Ctrl+C untuk berhenti.
echo.
node server.js
pause
