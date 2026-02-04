@echo off
echo ================================================
echo   Building Auto Switcher v3.1.63 EXE
echo ================================================
echo.

pip install pyinstaller >nul 2>&1

echo Building executable...
pyinstaller --onefile --noconsole --name auto_switcher_v3.1.63 auto_switcher_v3.1.63.py

echo.
echo Creating distribution folder...
if not exist "dist\auto_switcher_v3.1.63" mkdir "dist\auto_switcher_v3.1.63"

copy "dist\auto_switcher_v3.1.63.exe" "dist\auto_switcher_v3.1.63\" >nul
copy "config.ini" "dist\auto_switcher_v3.1.63\" >nul
copy "ignore_words.txt" "dist\auto_switcher_v3.1.63\" >nul
copy "hebrew_words.txt" "dist\auto_switcher_v3.1.63\" >nul
copy "align_left.png" "dist\auto_switcher_v3.1.63\" >nul
copy "align_right.png" "dist\auto_switcher_v3.1.63\" >nul
copy "README.txt" "dist\auto_switcher_v3.1.63\" >nul
copy "demo.html" "dist\auto_switcher_v3.1.63\" >nul

(
echo @echo off
echo set "SCRIPT_DIR=%%~dp0"
echo set "EXE_PATH=%%SCRIPT_DIR%%auto_switcher_v3.1.63.exe"
echo set "STARTUP_FOLDER=%%APPDATA%%\Microsoft\Windows\Start Menu\Programs\Startup"
echo.
echo ^(
echo echo Set WshShell = CreateObject^^^("WScript.Shell"^^^)
echo echo WshShell.Run """%%EXE_PATH%%""", 0, False
echo ^) ^> "%%STARTUP_FOLDER%%\HebrewEnglishSwitcher.vbs"
echo.
echo echo Added v3.1.63 to startup!
echo pause
) > "dist\auto_switcher_v3.1.63\add_to_startup.bat"

copy "remove_from_startup.bat" "dist\auto_switcher_v3.1.63\" >nul

echo.
echo Creating ZIP file...
powershell Compress-Archive -Path "dist\auto_switcher_v3.1.63" -DestinationPath "auto_switcher_v3.1.63.zip" -Force

echo.
echo ================================================
echo   BUILD COMPLETE!
echo ================================================
echo.
echo Distribution folder: dist\auto_switcher_v3.1.63
echo ZIP file: auto_switcher_v3.1.63.zip
echo.
pause
