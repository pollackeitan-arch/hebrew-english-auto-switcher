# Hebrew-English Auto Keyboard Switcher

## What this project does
Windows app that auto-detects words typed in the wrong keyboard layout (Hebrew vs English), fixes them in real-time, and auto-aligns text direction (RTL/LTR). Runs silently in the background.

## Project structure
- `auto_switcher_v3.1.64.py` - Main source code (1340 lines, single file)
- `build_exe.bat` - Builds exe via PyInstaller, creates dist folder + zip
- `install.bat` - For end users: registers auto-start + launches exe
- `uninstall.bat` - Stops app + removes from startup
- `run.bat` / `run_debug.bat` - Run Python script directly (dev use)
- `hebrew_words.txt` - Hebrew dictionary (470K words)
- `ignore_words.txt` - Words to never auto-fix
- `learned_words.txt` - User-taught words via Ctrl+`
- `config.ini` - Settings (block delay)
- `align_left.png` / `align_right.png` - Outlook button images for alignment

## How to build
Run `build_exe.bat` — it uses PyInstaller to create a standalone exe, assembles a dist folder with all needed files, and creates a zip for distribution.

## How to test changes
1. Edit `auto_switcher_v3.1.64.py`
2. Run `run.bat` (or `run_debug.bat` for verbose output)
3. Test the behavior
4. If good → commit, push, and optionally rebuild exe + create new release

## GitHub
- Repo: https://github.com/pollackeitan-arch/hebrew-english-auto-switcher
- Releases contain the distributable zip (exe + data files + install.bat)
- The exe, zip, build/, dist/, .spec files are gitignored

## Key dependencies
pynput, pywin32, keyboard, pyenchant, pyautogui, opencv-python, pyperclip

## Important notes
- Windows only (uses win32 APIs, keyboard hooks, WM_INPUTLANGCHANGEREQUEST)
- The exe runs via a VBS script in Windows Startup folder
- When version changes, update: filename, build_exe.bat, run.bat, run_debug.bat, install.bat, README.txt
- Don't modify the running exe — kill it first, rebuild, then restart
