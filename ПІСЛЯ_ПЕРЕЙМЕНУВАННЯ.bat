@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Відновлення залежностей CoDA після перейменування папки...
where pnpm >nul 2>nul
if errorlevel 1 (
  echo.
  echo pnpm не знайдено. Встановіть Node.js LTS, потім виконайте:
  echo corepack enable
  echo pnpm install --force
  pause
  exit /b 1
)
call pnpm install --force
if errorlevel 1 (
  echo.
  echo Не вдалося відновити залежності.
  pause
  exit /b 1
)
echo.
echo Залежності відновлено. Запуск CoDA...
call pnpm start
