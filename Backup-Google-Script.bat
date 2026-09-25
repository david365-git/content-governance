@echo off
title Google Apps Script Backup

cd /d "F:\!!!clasp-downloads"

echo.
echo ========================================
echo   Google Apps Script Backup
echo ========================================
echo.

echo Downloading latest Apps Script files...
echo.

call clasp pull

if errorlevel 1 (
    echo.
    echo ERROR: clasp pull failed.
    echo.
    pause
    exit /b 1
)

echo.
echo Download complete.
echo.

if not exist "Backups" mkdir "Backups"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set TIMESTAMP=%%i

echo Creating backup ZIP...

powershell -NoProfile -Command ^
"$files = Get-ChildItem -Force | Where-Object { $_.Name -ne 'Backups' -and $_.Name -ne 'Backup-Google-Script.bat' }; Compress-Archive -Path $files.FullName -DestinationPath 'Backups\AppsScript_%TIMESTAMP%.zip' -Force"

echo.
echo ========================================
echo   BACKUP COMPLETE
echo ========================================
echo.
echo Latest files are in:
echo F:\!!!clasp-downloads
echo.
echo Backup saved as:
echo F:\!!!clasp-downloads\Backups\AppsScript_%TIMESTAMP%.zip
echo.

pause