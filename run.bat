@echo off
echo Installing dependencies...
pip install pynput pywin32 keyboard pyenchant pyautogui opencv-python
echo.
echo Starting Auto Switcher...
python auto_switcher_v3.1.64.py
pause
