# Hebrew-English Auto Keyboard Switcher

## What this project does
Windows app that auto-detects words typed in the wrong keyboard layout (Hebrew vs English), fixes them in real-time, and auto-aligns text direction (RTL/LTR). Runs silently in the background.

## Project structure
- `auto_switcher_v3.1.65.py` - Main source code (~1650 lines, single file)
- `build_exe.bat` - Builds exe via PyInstaller, creates dist folder + zip
- `install.bat` - End-user installer (copied into the dist folder by build_exe.bat): stops any running version, copies exe + data files to `%LOCALAPPDATA%\HebrewSwitcher` (never overwrites existing learned/ignore/config), registers auto-start, launches it
- `uninstall.bat` - Stops app (any version) + removes from startup
- `run.bat` / `run_debug.bat` - Run Python script directly (dev use)
- `hebrew_words.txt` - Hebrew dictionary (470K words)
- `ignore_words.txt` - Words to never auto-fix
- `learned_words.txt` - User-taught words via Ctrl+`
- `config.ini` - Settings (block delay)
- `align_left.png` / `align_right.png` - Outlook button images for alignment

## How to build
Run `build_exe.bat` — it uses PyInstaller to create a standalone exe, assembles a dist folder with all needed files, and creates a zip for distribution.

## How to test changes
1. Edit `auto_switcher_v3.1.65.py`
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
- Auto-start uses a Task Scheduler task named `HebrewEnglishSwitcher` (AtLogOn). The old VBS/Registry Run methods are removed by install/uninstall.
- When version changes, update: filename, build_exe.bat, run.bat, run_debug.bat, install.bat, README.txt
- Don't modify the running exe — kill it first, rebuild, then restart
- The installed copy lives in `%LOCALAPPDATA%\HebrewSwitcher` — rebuilding here does NOT update it; run `dist\auto_switcher_v3.1.65\install.bat`
- Data files (learned_words.txt, ignore_words.txt, switcher_log.txt) used at runtime are the ones in `%LOCALAPPDATA%\HebrewSwitcher`, not this folder

## Current state (Nov 2026)
- Machine: Lenovo PC, project at `C:\Claude projects\hebrew-english-auto-switcher-20261004T191411Z-1-001\hebrew-english-auto-switcher` — it IS a git checkout of the GitHub repo
- Python 3.12 is installed with all dependencies + PyInstaller, so run.bat / build_exe.bat work
- v3.1.65 (Nov 7, 2026): reads the real system keyboard language at startup and on wake/unlock/logon instead of assuming English
