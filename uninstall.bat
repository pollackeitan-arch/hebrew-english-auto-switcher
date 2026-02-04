@echo off
echo Stopping Auto Switcher...
taskkill /IM auto_switcher_v3.1.63.exe /F >nul 2>&1
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul
echo Done! Auto Switcher has been stopped and removed from startup.
pause
