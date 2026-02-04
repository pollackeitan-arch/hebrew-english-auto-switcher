@echo off
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%auto_switcher_v3.1.63.exe"
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run """%EXE_PATH%""", 0, False
) > "%STARTUP_FOLDER%\HebrewEnglishSwitcher.vbs"

echo Added v3.1.63 to startup!
pause
