@echo off
pip install pynput pywin32 keyboard pyenchant pyautogui opencv-python >nul 2>&1
python "%~dp0auto_switcher_v3.1.63.py"
pause
