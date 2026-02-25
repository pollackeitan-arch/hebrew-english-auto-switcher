# Hebrew-English Auto Keyboard Switcher

🔄 **Automatically fixes words typed in the wrong keyboard language**

Ever started typing in Hebrew and realized your keyboard was on English? Or vice versa? This tool detects the mistake and fixes it instantly.

![Demo](demo.gif)

## ✨ Features

- **Auto-correction** - Detects wrong language and fixes the word when you press Space
- **Auto-alignment** - Sets text direction (RTL/LTR) based on the first word
- **Smart detection** - Uses dictionaries with 470,000+ Hebrew words and English spell-checking
- **Learn new words** - Press `Ctrl+`` ` on unfixed words to teach the program
- **Works everywhere** - Outlook, Word, Notepad, browsers, and more
- **Runs silently** - No window, just works in the background

## 🚀 Quick Start

### Option 1: Download Ready-to-Use (Recommended)
1. Go to [Releases](../../releases)
2. Download `AutoSwitcher.zip`
3. Extract and run `AutoSwitcher.exe`

### Option 2: Run from Source
```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/hebrew-english-auto-switcher.git
cd hebrew-english-auto-switcher

# Install dependencies
pip install pynput pywin32 keyboard pyenchant pyautogui opencv-python

# Run
python auto_switcher.py
```

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt+Shift` | Toggle language (auto-detected) |
| `Ctrl+`` ` | Undo last fix / Force-fix unfixed word |
| `Ctrl+Alt+E` | Force English mode |
| `Ctrl+Alt+H` | Force Hebrew mode |
| `Ctrl+Alt+A` | About dialog |
| `Ctrl+Alt+Q` | Quit |

## 📖 How It Works

1. **You type**: `akuo` (meant to type שלום but keyboard was on English)
2. **You press**: Space
3. **Tool detects**: "akuo" isn't English, but maps to valid Hebrew "שלום"
4. **Tool fixes**: Deletes "akuo", switches keyboard, types "שלום"

All in milliseconds - you barely notice it happened!

## 🔧 Configuration

### Files
| File | Purpose |
|------|---------|
| `config.ini` | Settings (timing delays) |
| `ignore_words.txt` | Words to never auto-correct |
| `learned_words.txt` | Words you taught via Ctrl+` |
| `hebrew_words.txt` | Hebrew dictionary |

### Auto-Start with Windows
Run `add_to_startup.bat` to launch automatically on login.

## 🛠️ Building from Source

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --onefile --noconsole --name AutoSwitcher auto_switcher.py
```

## 📋 Requirements

- Windows 10/11
- Python 3.8+ (if running from source)

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests

## 📄 License

MIT License - see [LICENSE](LICENSE) file.

**Created by Eitan Pollack**

---

⭐ **If you find this useful, please star the repository!**
