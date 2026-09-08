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
echo   Hasil ditulis ke raw\ROH.xlsx, raw\INT.xlsx, raw\RT.xlsx
echo ============================================================
echo.

".venv\Scripts\python.exe" main.py --tanya
set KODE=%ERRORLEVEL%

echo.
if %KODE% neq 0 (
    echo *** Selesai dengan error. Baca pesan di atas. ***
    echo.
    rem Error ditahan supaya pesannya sempat dibaca.
    pause
) else (
    echo *** Selesai. ***
    echo.
    echo   Langkah berikutnya: buka shofwans_tools_v4.xlsx,
    echo   lalu Data ^> Refresh All  [Ctrl+Alt+F5]  untuk menarik data baru.
)
exit /b %KODE%
