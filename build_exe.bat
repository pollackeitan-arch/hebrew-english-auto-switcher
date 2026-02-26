@echo off
echo ================================================
echo   Building Auto Switcher v3.1.64 EXE
echo ================================================
echo.

pip install pyinstaller >nul 2>&1

echo Building executable...
pyinstaller --onefile --noconsole --name auto_switcher_v3.1.64 auto_switcher_v3.1.64.py

echo.
echo Creating distribution folder...
if not exist "dist\auto_switcher_v3.1.64" mkdir "dist\auto_switcher_v3.1.64"

copy "dist\auto_switcher_v3.1.64.exe" "dist\auto_switcher_v3.1.64\" >nul
copy "config.ini" "dist\auto_switcher_v3.1.64\" >nul
copy "ignore_words.txt" "dist\auto_switcher_v3.1.64\" >nul
copy "hebrew_words.txt" "dist\auto_switcher_v3.1.64\" >nul
copy "align_left.png" "dist\auto_switcher_v3.1.64\" >nul
copy "align_right.png" "dist\auto_switcher_v3.1.64\" >nul
copy "README.txt" "dist\auto_switcher_v3.1.64\" >nul
copy "demo.html" "dist\auto_switcher_v3.1.64\" >nul

copy "install.bat" "dist\auto_switcher_v3.1.64\" >nul
copy "uninstall.bat" "dist\auto_switcher_v3.1.64\" >nul

echo.
echo Creating ZIP file...
powershell Compress-Archive -Path "dist\auto_switcher_v3.1.64" -DestinationPath "auto_switcher_v3.1.64.zip" -Force

echo.
echo ================================================
echo   BUILD COMPLETE!
echo ================================================
echo.
echo Distribution folder: dist\auto_switcher_v3.1.64
echo ZIP file: auto_switcher_v3.1.64.zip
echo.
pause
