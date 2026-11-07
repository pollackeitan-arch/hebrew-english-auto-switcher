@echo off
echo Stopping Auto Switcher...
powershell -Command "Get-Process auto_switcher_v* -ErrorAction SilentlyContinue | Stop-Process -Force" >nul 2>&1
REM Remove Task Scheduler entry
powershell -Command "Unregister-ScheduledTask -TaskName 'HebrewEnglishSwitcher' -Confirm:$false -ErrorAction SilentlyContinue"
REM Remove old registry/VBS startup methods if present
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "HebrewEnglishSwitcher" /f >nul 2>&1
powershell -Command "Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' -Name 'HebrewEnglishSwitcher' -ErrorAction SilentlyContinue" >nul 2>&1
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\HebrewEnglishSwitcher.vbs" 2>nul
echo Done! Auto Switcher has been stopped and removed from startup.
pause
