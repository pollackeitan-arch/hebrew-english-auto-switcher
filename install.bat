@echo off
set "SCRIPT_DIR=%~dp0"
set "EXE_NAME=auto_switcher_v3.1.65.exe"
set "EXE_SRC=%SCRIPT_DIR%%EXE_NAME%"
set "INSTALL_DIR=%LOCALAPPDATA%\HebrewSwitcher"
set "EXE_PATH=%INSTALL_DIR%\%EXE_NAME%"

if not exist "%EXE_SRC%" (
    echo ERROR: %EXE_NAME% not found next to install.bat
    pause
    exit /b 1
)

REM Stop any running version (old or new) so its exe can be replaced
powershell -Command "Get-Process auto_switcher_v* -ErrorAction SilentlyContinue | Stop-Process -Force" >nul 2>&1
ping -n 3 127.0.0.1 >nul

REM Remove old startup methods if present
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "HebrewEnglishSwitcher" /f >nul 2>&1

REM Install to LocalAppData (outside OneDrive, always available at logon)
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
del "%INSTALL_DIR%\auto_switcher_v*.exe" >nul 2>&1
copy /Y "%EXE_SRC%" "%INSTALL_DIR%\" >nul
copy /Y "%SCRIPT_DIR%hebrew_words.txt" "%INSTALL_DIR%\" >nul
copy /Y "%SCRIPT_DIR%align_left.png" "%INSTALL_DIR%\" >nul
copy /Y "%SCRIPT_DIR%align_right.png" "%INSTALL_DIR%\" >nul

REM User data: only copy if missing, so an upgrade keeps your settings and words
if not exist "%INSTALL_DIR%\config.ini" copy "%SCRIPT_DIR%config.ini" "%INSTALL_DIR%\" >nul
if not exist "%INSTALL_DIR%\ignore_words.txt" copy "%SCRIPT_DIR%ignore_words.txt" "%INSTALL_DIR%\" >nul
if not exist "%INSTALL_DIR%\learned_words.txt" if exist "%SCRIPT_DIR%learned_words.txt" copy "%SCRIPT_DIR%learned_words.txt" "%INSTALL_DIR%\" >nul

REM Use Task Scheduler for startup (most reliable on Windows 11)
powershell -Command "Unregister-ScheduledTask -TaskName 'HebrewEnglishSwitcher' -Confirm:$false -ErrorAction SilentlyContinue; $a = New-ScheduledTaskAction -Execute '%EXE_PATH%'; $t = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME; $s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero); $p = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited; Register-ScheduledTask -TaskName 'HebrewEnglishSwitcher' -Action $a -Trigger $t -Settings $s -Principal $p -Description 'Hebrew-English Auto Keyboard Switcher' | Out-Null"

echo Added v3.1.65 to startup!
echo Starting Auto Switcher...
start "" "%EXE_PATH%"
echo Done! Auto Switcher is running and will start automatically on login.
pause
