@echo off
cd /d "%~dp0"
if not exist node_modules\electron\dist\electron.exe (
  echo Electron is not installed. Run: pnpm install
  pause
  exit /b 1
)
node_modules\electron\dist\electron.exe .
