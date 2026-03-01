@echo off
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%auto_switcher_v3.1.64.exe"

REM Remove old startup methods if present
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "HebrewEnglishSwitcher" /f >nul 2>&1

REM Use Task Scheduler for startup (most reliable on Windows 11)
powershell -Command "Unregister-ScheduledTask -TaskName 'HebrewEnglishSwitcher' -Confirm:$false -ErrorAction SilentlyContinue; $a = New-ScheduledTaskAction -Execute '%EXE_PATH%'; $t = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME; $s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero); $p = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited; Register-ScheduledTask -TaskName 'HebrewEnglishSwitcher' -Action $a -Trigger $t -Settings $s -Principal $p -Description 'Hebrew-English Auto Keyboard Switcher' | Out-Null"

echo Added v3.1.64 to startup!
echo Starting Auto Switcher...
start "" "%EXE_PATH%"
echo Done! Auto Switcher is running and will start automatically on login.
pause
