@echo off
title Traffic Counting ITDP - Web Console
color 0b
echo =====================================================================
echo       SISTEM WEB TRAFFIC COUNTING ITDP CLEAR JABODETABEK
echo =====================================================================
echo.
echo [1/2] Memeriksa instalasi Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python belum terinstal atau tidak ada di PATH.
    pause
    exit /b
)

echo [2/2] Membuka server aplikasi di port 5000...
echo.
echo Akses aplikasi melalui browser di:
echo http://127.0.0.1:5000
echo.
echo Tekan Ctrl+C di terminal ini jika ingin menghentikan server.
echo =====================================================================
echo.

start "" http://127.0.0.1:5000
python app.py

pause
