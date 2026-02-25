@echo off
echo Adding Auto Switcher to Windows startup...

set "SCRIPT_DIR=%~dp0"
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

REM Check if running from exe or script
if exist "%SCRIPT_DIR%AutoSwitcher.exe" (
    set "TARGET=%SCRIPT_DIR%AutoSwitcher.exe"
) else (
    set "TARGET=%SCRIPT_DIR%auto_switcher_v3.1.64.py"
)

(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run """%TARGET%""", 0, False
) > "%STARTUP_FOLDER%\AutoSwitcher.vbs"

echo.
echo SUCCESS! Auto Switcher will start automatically on login.
echo To remove, run remove_from_startup.bat
echo.
pause
