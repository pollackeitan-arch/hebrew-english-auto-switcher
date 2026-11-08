@echo off
echo ================================================
echo   Building Auto Switcher v3.1.66 EXE
echo ================================================
echo.

python -m pip install pyinstaller >nul 2>&1

echo Building executable...
python -m PyInstaller --noconfirm --onefile --noconsole --name auto_switcher_v3.1.66 auto_switcher_v3.1.66.py

echo.
echo Creating distribution folder...
if not exist "dist\auto_switcher_v3.1.66" mkdir "dist\auto_switcher_v3.1.66"

copy "dist\auto_switcher_v3.1.66.exe" "dist\auto_switcher_v3.1.66\" >nul
copy "config.ini" "dist\auto_switcher_v3.1.66\" >nul
copy "ignore_words.txt" "dist\auto_switcher_v3.1.66\" >nul
copy "hebrew_words.txt" "dist\auto_switcher_v3.1.66\" >nul
copy "align_left.png" "dist\auto_switcher_v3.1.66\" >nul
copy "align_right.png" "dist\auto_switcher_v3.1.66\" >nul
copy "README.txt" "dist\auto_switcher_v3.1.66\" >nul
copy "demo.html" "dist\auto_switcher_v3.1.66\" >nul

copy "install.bat" "dist\auto_switcher_v3.1.66\" >nul
copy "uninstall.bat" "dist\auto_switcher_v3.1.66\" >nul

echo.
echo Creating ZIP file...
powershell Compress-Archive -Path "dist\auto_switcher_v3.1.66" -DestinationPath "auto_switcher_v3.1.66.zip" -Force

echo.
echo ================================================
echo   BUILD COMPLETE!
echo ================================================
echo.
echo Distribution folder: dist\auto_switcher_v3.1.66
echo ZIP file: auto_switcher_v3.1.66.zip
echo.
pause
