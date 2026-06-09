@echo off
cd /d "%~dp0front"
echo.
echo  SYQ Dev Server - http://localhost:5274
echo  Stop: Ctrl+C
echo.
npm run dev
pause
