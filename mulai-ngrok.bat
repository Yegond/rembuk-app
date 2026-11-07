@echo off
echo Pastikan jendela "mulai.bat" sudah terbuka dan menampilkan "berjalan".
echo Membuka link publik lewat ngrok...
echo Salin alamat "Forwarding" (https://....ngrok-free.app) dan bagikan.
echo.
ngrok http 127.0.0.1:3000
pause
