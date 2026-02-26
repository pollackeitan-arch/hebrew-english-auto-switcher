@echo off
echo Stopping Auto Switcher...
taskkill /IM auto_switcher_v3.1.64.exe /F >nul 2>&1
REM Remove registry startup entry
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "HebrewEnglishSwitcher" /f >nul 2>&1
REM Remove old VBS startup method if present
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul
echo Done! Auto Switcher has been stopped and removed from startup.
pause
