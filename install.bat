@echo off
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%auto_switcher_v3.1.64.exe"

REM Remove old VBS startup method if present
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul

REM Add to registry Run key (reliable on all Windows versions)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "HebrewEnglishSwitcher" /t REG_SZ /d "\"%EXE_PATH%\"" /f >nul

REM Mark as approved in StartupApproved (required by Windows 11)
powershell -Command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' -Name 'HebrewEnglishSwitcher' -Value ([byte[]](2,0,0,0,0,0,0,0,0,0,0,0))" >nul 2>&1

echo Added v3.1.64 to startup!
echo Starting Auto Switcher...
start "" "%EXE_PATH%"
echo Done! Auto Switcher is running and will start automatically on login.
pause
