@echo off
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%auto_switcher_v3.1.64.exe"
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run """%EXE_PATH%""", 0, False
) > "%STARTUP_FOLDER%\HebrewEnglishSwitcher.vbs"

echo Added v3.1.64 to startup!
echo Starting Auto Switcher...
start "" "%EXE_PATH%"
echo Done! Auto Switcher is running and will start automatically on login.
pause
