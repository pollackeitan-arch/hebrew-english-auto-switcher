@echo off
echo ================================================
echo   Building Auto Switcher
echo   Created by Eitan Pollack
echo ================================================
echo.

REM Install PyInstaller if needed
pip install pyinstaller >nul 2>&1

REM Clean previous builds
echo Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"

echo.
echo Building executable...
pyinstaller --onefile --noconsole --name AutoSwitcher auto_switcher_v3.1.64.py

if not exist "dist\AutoSwitcher.exe" (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ================================================
echo   BUILD COMPLETE!
echo ================================================
echo.
echo   Executable: dist\AutoSwitcher.exe
echo.
echo   To create a release:
echo   1. Copy dist\AutoSwitcher.exe
echo   2. Add data files (config.ini, *.txt, *.png)
echo   3. Zip and upload to GitHub Releases
echo.
pause
