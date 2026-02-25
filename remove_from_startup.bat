@echo off
echo Removing Auto Switcher from Windows startup...
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\AutoSwitcher.vbs" 2>nul
echo.
echo Done! Auto Switcher removed from startup.
pause
