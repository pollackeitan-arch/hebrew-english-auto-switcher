@echo off
echo ================================================
echo   Building Auto Switcher v3.1.65 EXE
echo ================================================
echo.

python -m pip install pyinstaller >nul 2>&1

echo Building executable...
python -m PyInstaller --noconfirm --onefile --noconsole --name auto_switcher_v3.1.65 auto_switcher_v3.1.65.py

echo.
echo Creating distribution folder...
if not exist "dist\auto_switcher_v3.1.65" mkdir "dist\auto_switcher_v3.1.65"

copy "dist\auto_switcher_v3.1.65.exe" "dist\auto_switcher_v3.1.65\" >nul
copy "config.ini" "dist\auto_switcher_v3.1.65\" >nul
copy "ignore_words.txt" "dist\auto_switcher_v3.1.65\" >nul
copy "hebrew_words.txt" "dist\auto_switcher_v3.1.65\" >nul
copy "align_left.png" "dist\auto_switcher_v3.1.65\" >nul
copy "align_right.png" "dist\auto_switcher_v3.1.65\" >nul
copy "README.txt" "dist\auto_switcher_v3.1.65\" >nul
copy "demo.html" "dist\auto_switcher_v3.1.65\" >nul

copy "install.bat" "dist\auto_switcher_v3.1.65\" >nul
copy "uninstall.bat" "dist\auto_switcher_v3.1.65\" >nul

echo.
echo Creating ZIP file...
powershell Compress-Archive -Path "dist\auto_switcher_v3.1.65" -DestinationPath "auto_switcher_v3.1.65.zip" -Force

echo.
echo ================================================
echo   BUILD COMPLETE!
echo ================================================
echo.
echo Distribution folder: dist\auto_switcher_v3.1.65
echo ZIP file: auto_switcher_v3.1.65.zip
echo.
pause
