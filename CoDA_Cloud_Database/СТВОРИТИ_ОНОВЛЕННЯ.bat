@echo off
chcp 65001 >nul
set /p CODA_VERSION=Введіть унікальну версію з Налаштування!B3 (наприклад 2026.09.01.1):
set /p CODA_DATE=Введіть дату з Налаштування!B2 у форматі РРРР-ММ-ДД:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Build-CoDAUpdate.ps1" -Mode Update -Version "%CODA_VERSION%" -DatabaseDate "%CODA_DATE%"
echo.
pause
