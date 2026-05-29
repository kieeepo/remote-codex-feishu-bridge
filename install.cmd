@echo off
cd /d "%~dp0"
echo Installing Remote Codex Feishu Bridge dependencies...
npm.cmd install
if not exist ".env.local" (
  copy ".env.example" ".env.local" >nul
  echo Created .env.local from .env.example
)
echo.
echo Install finished.
echo Next: edit .env.local, then double-click start-feishu.cmd
pause
