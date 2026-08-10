@echo off
chcp 65001 >nul
set /p CODA_VERSION=Введіть версію пакета:
set /p CODA_URL=Вставте публічне посилання GitHub Release:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Build-CoDAUpdate.ps1" -Mode SetUrl -Version "%CODA_VERSION%" -Url "%CODA_URL%"
echo.
pause
