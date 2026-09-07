@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Ambil data HDKS - ROH / INT / RT

if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment .venv tidak ditemukan.
    echo Jalankan dulu:  python -m venv .venv ^&^& .venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo   Ambil data ROH / INT / RT dari HDKS
echo   Hasil ditulis ke sheet ROH, INT, RT di shofwans_tools_v3.xlsx
echo ============================================================
echo.

".venv\Scripts\python.exe" main.py --tanya
set KODE=%ERRORLEVEL%

echo.
if %KODE% neq 0 (
    echo *** Selesai dengan error. Baca pesan di atas. ***
) else (
    echo *** Selesai. ***
)
echo.
pause
exit /b %KODE%
