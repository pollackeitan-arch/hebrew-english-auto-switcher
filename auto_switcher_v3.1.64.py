"""
Auto Switcher v3.1.64
==================
Hebrew-English keyboard switcher with auto-alignment.

Features:
- Track language changes via Alt+Shift
- Use key.vk (physical keys) - always reliable
- Manual override: Ctrl+Alt+E (English), Ctrl+Alt+H (Hebrew)
- Skip numbers, common short words, ignore list
- Full Hebrew dictionary (470,000 words)
- Resets to English on wake from sleep
- AUTO ALIGNMENT: First word Hebrew → Align Right, English → Align Left

Press Ctrl+Alt+Q to quit
"""

import threading
import time
import ctypes
import sys
import os
import configparser

try:
    from pynput import keyboard as pynput_keyboard
    from pynput.keyboard import Key, Controller as KeyboardController, KeyCode
    from pynput import mouse as pynput_mouse
    import win32api
    import win32gui
    import win32con
    import keyboard
    import pyautogui
    import pyperclip
except ImportError as e:
    print("Missing packages. Run: pip install pynput pywin32 keyboard pyenchant pyautogui opencv-python pyperclip")
    sys.exit(1)

# Try to import enchant for English dictionary
try:
    import enchant
    ENCHANT_AVAILABLE = True
except ImportError:
    ENCHANT_AVAILABLE = False

# Configure pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


def get_script_dir():
    """Get the directory where the script/exe is located"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def load_config():
    """Load config from config.ini"""
    config_path = os.path.join(get_script_dir(), 'config.ini')
    
    config = configparser.ConfigParser()
    
    if os.path.exists(config_path):
        config.read(config_path)
    else:
        config['Settings'] = {'block_delay_ms': '2000'}
        with open(config_path, 'w') as f:
            f.write("[Settings]\n")
            f.write("# Time in milliseconds to keep keyboard blocked after fix\n")
            f.write("block_delay_ms=2000\n")
    
    try:
        block_delay_ms = config.getint('Settings', 'block_delay_ms')
    except:
        block_delay_ms = 2000
    
    return {'block_delay_ms': block_delay_ms}


def load_ignored_words():
    """Load ignored words from ignore_words.txt"""
    words_path = os.path.join(get_script_dir(), 'ignore_words.txt')
    ignored = set()
    
    if os.path.exists(words_path):
        try:
            with open(words_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        ignored.add(line)
                        ignored.add(line.lower())
            print(f"  Loaded {len(ignored)//2} ignored words")
        except Exception as e:
            print(f"  Error loading ignore_words.txt: {e}")
    
    return ignored


def load_learned_words():
    """Load learned words from learned_words.txt
    Format: keys,target_language (e.g., akuo,hebrew)
    Returns dict: {keys: target_language}
    """
    words_path = os.path.join(get_script_dir(), 'learned_words.txt')
    learned = {}
    
    if os.path.exists(words_path):
        try:
            with open(words_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and ',' in line:
                        parts = line.split(',', 1)
                        if len(parts) == 2:
                            keys, target_lang = parts
                            learned[keys.lower()] = target_lang
            print(f"  Loaded {len(learned)} learned words")
        except Exception as e:
            print(f"  Error loading learned_words.txt: {e}")
    
    return learned


def load_hebrew_dictionary():
    """Load Hebrew dictionary from hebrew_words.txt"""
    dict_path = os.path.join(get_script_dir(), 'hebrew_words.txt')
    words = set()
    
    if os.path.exists(dict_path):
        try:
            with open(dict_path, 'r', encoding='utf-8') as f:
                for line in f:
                    word = line.strip()
                    if word:
                        words.add(word)
            print(f"  Loaded Hebrew dictionary: {len(words):,} words")
        except Exception as e:
            print(f"  Error loading hebrew_words.txt: {e}")
    else:
        print(f"  Warning: hebrew_words.txt not found, using built-in list")
    
    return words


CONFIG = load_config()
IGNORED_WORDS = load_ignored_words()
LEARNED_WORDS = load_learned_words()
HEBREW_DICTIONARY = load_hebrew_dictionary()


# Common short English words (2-3 letters) and contractions
COMMON_SHORT_ENGLISH = {
    'hi', 'no', 'ok', 'go', 'yes', 'the', 'and', 'but', 'for', 'not', 'you', 'all', 'can',
    'had', 'her', 'was', 'one', 'our', 'out', 'are', 'has', 'his', 'how', 'its', 'may',
    'new', 'now', 'old', 'see', 'two', 'way', 'who', 'boy', 'did', 'get', 'him', 'let',
    'put', 'say', 'she', 'too', 'use', 'me', 'my', 'we', 'he', 'it', 'is', 'in', 'on',
    'so', 'to', 'up', 'us', 'an', 'as', 'at', 'be', 'by', 'do', 'if', 'or',
    # Contractions
    "i'm", "i'd", "i'll", "i've", "don't", "can't", "won't", "didn't", "doesn't",
    "isn't", "aren't", "wasn't", "weren't", "haven't", "hasn't", "couldn't",
    "wouldn't", "shouldn't", "let's", "that's", "what's", "there's", "here's",
    "who's", "it's", "he's", "she's", "we're", "they're", "you're", "you've",
    "you'd", "you'll", "we've", "we'd", "we'll", "they've", "they'd", "they'll",
}

# Common short Hebrew words
SHORT_HEBREW_WORDS = {
    'אם', 'אל', 'אז', 'או', 'אף', 'בו', 'בה', 'בי', 'בך', 'גם', 'גב', 'דם', 'דר', 'הם', 'הן',
    'הו', 'וו', 'זה', 'זו', 'זא', 'חם', 'חי', 'טו', 'יד', 'כה', 'כי', 'לא', 'לב', 'לה', 'לו',
    'לי', 'לך', 'מה', 'מי', 'מן', 'נא', 'נו', 'סו', 'עד', 'על', 'עם', 'פה', 'פי', 'צו', 'צט',
    'קם', 'רב', 'רע', 'שם', 'תו', 'תן', 'אב', 'אח', 'בן', 'בת', 'גן', 'דג', 'דף', 'הר', 'וד',
    'זב', 'חג', 'חן', 'טל', 'יא', 'כד', 'כף', 'לח', 'מד', 'מט', 'נד', 'נח', 'סף', 'עז', 'פח',
    'צב', 'קו', 'רם', 'שד', 'שק', 'תל', 'גמר', 'זאת', 'חבר', 'טוב', 'יום', 'כאן', 'למה', 'מאד',
    'נכח', 'סוף', 'עוד', 'פעם', 'צאת', 'קצת', 'רגע', 'שוב', 'תוך', 'היי', 'כן',
}

# Common Hebrew words (expanded list)
HEBREW_WORDS = {
    'לא', 'את', 'אני', 'זה', 'אתה', 'מה', 'הוא', 'לי', 'על', 'כן', 'לך', 'של', 'יש', 'בסדר',
    'אבל', 'כל', 'שלי', 'טוב', 'עם', 'היא', 'אם', 'רוצה', 'שלך', 'היה', 'אנחנו', 'הם', 'אותך',
    'יודע', 'אז', 'רק', 'אותו', 'יכול', 'אותי', 'יותר', 'הזה', 'אל', 'כאן', 'או', 'למה', 'שאני',
    'כך', 'אחד', 'עכשיו', 'משהו', 'להיות', 'היי', 'תודה', 'כמו', 'אין', 'איך', 'זאת', 'נכון',
    'שלום', 'פה', 'הזאת', 'שם', 'בבקשה', 'כבר', 'לעשות', 'עוד', 'מי', 'שלו', 'תראה', 'לו',
    'ממש', 'צריך', 'ואני', 'שהוא', 'הייתי', 'קצת', 'אמר', 'אנשים', 'אחת', 'ידעתי', 'אוהב',
    'בא', 'לנו', 'לפני', 'ככה', 'שאתה', 'אפשר', 'מאוד', 'הנה', 'אמרתי', 'אותה', 'בו', 'זמן',
    'הכל', 'חושב', 'בית', 'שום', 'טובה', 'ובא', 'אתם', 'לדבר', 'בואו', 'בטח', 'אולי', 'כמה',
    'דבר', 'שזה', 'היום', 'חייב', 'הרבה', 'הוה', 'אמא', 'לראות', 'פעם', 'כזה', 'בדיוק', 'יפה',
    'הולך', 'הלילה', 'יהיה', 'קרה', 'מישהו', 'ביותר', 'רגע', 'להם', 'קורה', 'אלוהים', 'ילד',
    'שנה', 'עושה', 'מדבר', 'איפה', 'בטוח', 'חיים', 'ילדים', 'אומר', 'שהיא', 'עליו', 'שוב',
    'מקום', 'אבא', 'לקחת', 'חשבתי', 'ראית', 'דרך', 'יודעת', 'שני', 'עלי', 'להגיד', 'מר',
    'שאנחנו', 'אהיה', 'בה', 'הי', 'איתו', 'וזה', 'שלה', 'סליחה', 'בן', 'כלום', 'תן', 'יכולה',
    'עבודה', 'אדם', 'הביתה', 'ואתה', 'ראש', 'עליי', 'בחיים', 'ללכת', 'לעזאזל', 'ישר', 'גדול',
    'אחי', 'לכם', 'שלנו', 'הייתה', 'יודעים', 'אלה', 'חשוב', 'הראש', 'לחזור', 'אצל', 'תהיה',
    'רואה', 'מצטער', 'עושים', 'למצוא', 'ואז', 'מוכן', 'היית', 'לתת', 'כדי', 'חדש', 'תמיד',
    'ראיתי', 'כסף', 'לפה', 'לעולם', 'פשוט', 'ולא', 'בנות', 'בעיה', 'עדיין', 'יום', 'לגבי',
    'כאילו', 'לאן', 'חברים', 'שלהם', 'האמת', 'לכל', 'רע', 'סוף', 'תעשה', 'שהם', 'ממני',
    'מותק', 'מספיק', 'קח', 'לשם', 'לילה', 'הדבר', 'מזה', 'לקרות', 'עצמי', 'העולם', 'לה',
    'ביום', 'בחור', 'אליך', 'נראה', 'ולך', 'ימים', 'קשה', 'לכאן', 'הבית', 'בוא', 'מהר',
    'באמת', 'מבין', 'לך', 'מאמין', 'חדשות', 'נהדר', 'אחרת', 'חכה', 'ואת', 'בי', 'להביא',
    'גברת', 'שעה', 'תגיד', 'מכיר', 'אוכל', 'עזוב', 'קדימה', 'שאת', 'אליו', 'מעולם', 'שלושה',
    'בוקר', 'לומר', 'תוכל', 'רציתי', 'סתם', 'טיפש', 'הדרך', 'מילה', 'שמעתי', 'בך', 'ממך',
    'מגיע', 'מוזר', 'לעזור', 'בשביל', 'חבל', 'לזה', 'האיש', 'אמרת', 'יקר', 'כולם', 'לדעת',
    'בגלל', 'הבן', 'עליך', 'תראי', 'מנסה', 'מרגיש', 'שאם', 'משנה', 'צריכה', 'הזו', 'לבד',
    'בחוץ', 'אדוני', 'לעבוד', 'עצור', 'בעל', 'חי', 'לפחות', 'אהבה', 'רוצים', 'סיפור', 'למעלה',
    'הבחור', 'שנים', 'מקווה', 'בלי', 'בשבילך', 'מכאן', 'בתוך', 'לשמוע', 'מעט', 'חזק', 'שיהיה',
    'מדי', 'עשית', 'אחרי', 'שלא', 'וגם', 'מת', 'בפנים', 'נגיד', 'אישה', 'לבוא', 'בכל', 'להרוג',
    'תקשיב', 'שמע', 'משחק', 'גבר', 'כלב', 'יפה', 'להתראות', 'מיד', 'חכי', 'הילד', 'נו',
    'תשמע', 'שוטר', 'הילדים', 'מעולה', 'חבר', 'סיבה', 'הלו', 'נוכל', 'מחר', 'עצמך', 'בחורה',
    'לעבור', 'הערב', 'מים', 'מה', 'הבא', 'שמח', 'מתי', 'לספר', 'הייתם', 'אשתי', 'דקות',
    'לכה', 'קטן', 'צודק', 'יוצא', 'נעים', 'דברים', 'לצאת', 'נתן', 'בשלב', 'אכפת', 'אחר',
    'תעשי', 'קרוב', 'הו', 'רב', 'מחכה', 'שעות', 'בעלי', 'ידי', 'חוץ', 'ואם', 'בראש', 'אותם',
    'הבחורה', 'הולכים', 'קטנה', 'תפסיק', 'אי', 'אוי', 'כולנו', 'יקרה', 'משם', 'אמור', 'נשמה',
    'רצית', 'אלא', 'לב', 'עולם', 'חייבים', 'הבאה', 'עליה', 'חושבת', 'כרגע', 'להבין', 'למטה',
    'בעיות', 'גרוע', 'מאז', 'מפה', 'שקרה', 'מכיר', 'אמרה', 'מתכוון', 'מהם', 'הכי', 'מוכנים',
    'לעצמך', 'בחזרה', 'הדברים', 'אחרים', 'שכל', 'להיכנס', 'אלי', 'רעיון', 'כעת', 'אגיד',
    'דקה', 'אלינו', 'הצילו', 'בעיר', 'לדאוג', 'ידע', 'לשים', 'בוודאי', 'האקדח', 'תנו', 'עולה',
    'שאוכל', 'היד', 'זוז', 'יד', 'כלבה', 'עובדת', 'ראה', 'נמצאת', 'כשהוא', 'אופן', 'שמה',
    'נוסף', 'הפה', 'דוקטור', 'עזוב', 'תרצה', 'במקרה', 'הטלפון', 'שמך', 'צעיר', 'לשבת',
    'מכיוון', 'הקטנה', 'שניכם', 'לתפוס', 'תלך', 'בהצלחה', 'יין', 'חכו', 'קשור', 'אנא', 'ככל',
    'דין', 'בנו', 'למישהו', 'השנה', 'הלוואי', 'בעצמי', 'לרגע', 'קר', 'ממנה', 'ישנה', 'בעצמך',
    'האל', 'שש', 'תוריד', 'התיק', 'איכפת', 'אחורה', 'יעשה', 'אכן', 'מעמד', 'לירות', 'מבקש',
    'החרא', 'בשעה', 'לבן', 'לבחור', 'עדיף', 'נזוז', 'אילו', 'ואל', 'כזאת', 'יקירתי', 'שקרה',
    'מישהי', 'איתם', 'חם', 'זוזו', 'שים', 'לטפל', 'מול', 'לכו', 'הפנים', 'אליה', 'בסוף',
    'בשקט', 'לכולם', 'מידע', 'אהבתי', 'לברוח', 'נשק', 'תחשוב', 'ספק',
}


class HebrewEnglishSwitcher:
    
    # English key to Hebrew character mapping (standard Israeli keyboard)
    ENGLISH_KEY_TO_HEBREW_CHAR = {
        'q': '/', 'w': "'", 'e': 'ק', 'r': 'ר', 't': 'א',
        'y': 'ט', 'u': 'ו', 'i': 'ן', 'o': 'ם', 'p': 'פ',
        'a': 'ש', 's': 'ד', 'd': 'ג', 'f': 'כ', 'g': 'ע',
        'h': 'י', 'j': 'ח', 'k': 'ל', 'l': 'ך', ';': 'ף',
        'z': 'ז', 'x': 'ס', 'c': 'ב', 'v': 'ה', 'b': 'נ',
        'n': 'מ', 'm': 'צ', ',': 'ת', '.': 'ץ', '/': '.',
    }
    
    HEBREW_CHAR_TO_ENGLISH_KEY = {v: k for k, v in ENGLISH_KEY_TO_HEBREW_CHAR.items()}
    
    # Virtual key codes to English letters (physical key mapping)
    VK_TO_ENGLISH = {
        65: 'a', 66: 'b', 67: 'c', 68: 'd', 69: 'e', 70: 'f', 71: 'g', 72: 'h',
        73: 'i', 74: 'j', 75: 'k', 76: 'l', 77: 'm', 78: 'n', 79: 'o', 80: 'p',
        81: 'q', 82: 'r', 83: 's', 84: 't', 85: 'u', 86: 'v', 87: 'w', 88: 'x',
        89: 'y', 90: 'z',
        48: '0', 49: '1', 50: '2', 51: '3', 52: '4', 53: '5', 54: '6', 55: '7', 56: '8', 57: '9',
        186: ';', 187: '=', 188: ',', 189: '-', 190: '.', 191: '/', 192: '`',
        219: '[', 220: '\\', 221: ']', 222: "'",
    }
    
    LANG_ENGLISH = 0x0409
    LANG_HEBREW = 0x040D
    
    def __init__(self, debug=False):
        self.is_running = True
        self.is_fixing = False
        self.last_switch_time = 0
        self.debug = debug
        self.pynput_controller = KeyboardController()
        
        # Language tracking - start with English (login language)
        self.tracked_language = 'english'
        
        # Current word being typed (physical keys)
        self.current_word_keys = ""
        
        # Track Alt and Shift for language switch detection
        self.alt_pressed = False
        self.shift_pressed = False
        
        self.WM_INPUTLANGCHANGEREQUEST = 0x0050
        
        self.block_input = ctypes.windll.user32.BlockInput
        self.block_input.argtypes = [ctypes.c_bool]
        self.block_input.restype = ctypes.c_bool
        
        # Initialize English dictionary
        self.english_dict = None
        if ENCHANT_AVAILABLE:
            for lang in ['en_US', 'en_GB', 'en']:
                try:
                    self.english_dict = enchant.Dict(lang)
                    if self.debug:
                        print(f"  Loaded English dictionary: {lang}")
                    break
                except:
                    pass
        
        # Load alignment button images (for Outlook)
        self.script_dir = get_script_dir()
        self.align_left_img = os.path.join(self.script_dir, 'align_left.png')
        self.align_right_img = os.path.join(self.script_dir, 'align_right.png')
        
        # Setup log file
        self.log_file = os.path.join(self.script_dir, 'switcher_log.txt')
        self.recent_words = []  # Buffer for last 10 words
        self.file_log("="*50)
        self.file_log(f"STARTUP - v3.1.64")
        self.file_log(f"Initial tracked_language: {self.tracked_language}")
        self.file_log("="*50)
        
        # Check if image files exist
        if not os.path.exists(self.align_left_img):
            print(f"  WARNING: align_left.png not found at: {self.align_left_img}")
        if not os.path.exists(self.align_right_img):
            print(f"  WARNING: align_right.png not found at: {self.align_right_img}")
        
        # Direction tracking - only change direction on first word of line
        self.is_first_word_of_line = True
        self.direction_set_for_line = False
        
        # Undo tracking
        self.last_fix_info = None  # Store info about last fix for undo detection
        self.ctrl_pressed = False
        
        # Force-fix tracking (for Ctrl+` on unfixed words)
        self.last_unfixed_word = None  # Store last word that wasn't auto-fixed
        self.last_unfixed_was_first = False  # Was it the first word of line?
        
        # Context tracking - clear buffer on window/mouse change
        self.last_active_window = win32gui.GetForegroundWindow()
        
        if self.debug:
            print("="*60)
            print("  Hebrew-English Auto Switcher v3.1.64")
            print("="*60)
            print(f"  English: {'pyenchant dictionary' if self.english_dict else 'heuristics'}")
            hebrew_count = len(HEBREW_DICTIONARY) if HEBREW_DICTIONARY else len(HEBREW_WORDS)
            print(f"  Hebrew: {hebrew_count:,} words")
            print(f"  Starting language: {self.tracked_language.upper()}")
            print(f"  Auto-direction: Enabled (first word only)")
            print("  ")
            print("  Hotkeys:")
            print("    Alt+Shift    - Toggle language (auto-detected)")
            print("    Ctrl+`       - Undo fix / Force-fix unfixed word")
            print("    Ctrl+Alt+E   - Force English")
            print("    Ctrl+Alt+H   - Force Hebrew")
            print("    Ctrl+Alt+A   - About")
            print("    Ctrl+Alt+R   - Release stuck keys (Caps Lock / Ctrl / Shift)")
            print("    Ctrl+Alt+Q   - Quit")
            print("="*60 + "\n")
    
    def log(self, msg):
        if self.debug:
            print(msg)
    
    def file_log(self, msg):
        """Write to log file with timestamp (only in debug mode)"""
        if not self.debug:
            return
        from datetime import datetime
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {msg}\n")
            
            # Keep log file small - max 1000 lines
            try:
                with open(self.log_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                if len(lines) > 1000:
                    with open(self.log_file, 'w', encoding='utf-8') as f:
                        f.writelines(lines[-500:])  # Keep last 500 lines
            except:
                pass
        except:
            pass
    
    def log_word(self, keys, screen_word, action):
        """Log a word to recent words buffer and file"""
        self.recent_words.append({
            'keys': keys,
            'screen': screen_word,
            'tracked': self.tracked_language,
            'action': action
        })
        # Keep only last 10 words
        if len(self.recent_words) > 10:
            self.recent_words.pop(0)
        
        self.file_log(f"WORD: keys='{keys}' screen='{screen_word}' tracked={self.tracked_language} action={action}")
    
    def keys_to_hebrew(self, english_keys):
        """Convert English keys to Hebrew characters"""
        return ''.join(self.ENGLISH_KEY_TO_HEBREW_CHAR.get(c, c) for c in english_keys.lower())
    
    def get_screen_word(self, keys):
        """Get what appears on screen based on tracked language"""
        if self.tracked_language == 'hebrew':
            return self.keys_to_hebrew(keys)
        else:
            return keys
    
    def is_valid_english(self, word):
        """Check if word is valid English"""
        if not word:
            return False
        
        word_lower = word.lower()
        
        # Check contractions first (they contain apostrophe)
        if "'" in word_lower:
            return word_lower in COMMON_SHORT_ENGLISH
        
        # Regular words must be alphabetic
        if not word.isalpha():
            return False
        
        if self.english_dict:
            try:
                return self.english_dict.check(word_lower)
            except:
                pass
        
        # Fallback heuristic
        vowels = sum(1 for c in word_lower if c in 'aeiou')
        ratio = vowels / len(word_lower) if word_lower else 0
        return 0.15 <= ratio <= 0.6
    
    def is_valid_hebrew(self, text):
        """Check if text is a valid Hebrew word using the full dictionary"""
        # First check the large dictionary
        if HEBREW_DICTIONARY and text in HEBREW_DICTIONARY:
            return True
        # Fallback to built-in lists if dictionary not loaded
        return text in HEBREW_WORDS or text in SHORT_HEBREW_WORDS
    
    def detect_language(self, keys):
        """
        Detect what language a word is, without correction.
        Returns 'hebrew' or 'english' or None if can't determine.
        """
        if len(keys) < 2:
            return None
        
        # Strip trailing punctuation for detection
        punctuation_chars = ',.!?;:'
        while keys and keys[-1] in punctuation_chars:
            keys = keys[:-1]
        
        if len(keys) < 2:
            return None
        
        hebrew_version = self.keys_to_hebrew(keys)
        english_version = keys.lower()
        
        if self.tracked_language == 'hebrew':
            # User is typing in Hebrew, check if it's valid Hebrew
            if self.is_valid_hebrew(hebrew_version):
                return 'hebrew'
            return None
        else:
            # User is typing in English, check if it's valid English
            if self.is_valid_english(english_version):
                return 'english'
            return None
    
    def analyze_and_fix(self, keys):
        """
        Analyze the typed keys and decide if correction is needed.
        keys: physical keys pressed (always English letters)
        Returns: (corrected_word, target_language) or (None, None)
        """
        if len(keys) < 2:
            return None, None
        
        # Strip trailing punctuation
        punctuation = ''
        punctuation_chars = ',.!?;:'
        while keys and keys[-1] in punctuation_chars:
            punctuation = keys[-1] + punctuation
            keys = keys[:-1]
        
        if len(keys) < 2:
            return None, None
        
        # Skip if contains numbers
        if any(c.isdigit() for c in keys):
            self.log(f"  Skipping (contains numbers): '{keys}'")
            return None, None
        
        # Check learned words first (user force-fixed these before)
        keys_lower = keys.lower()
        if keys_lower in LEARNED_WORDS:
            target_lang = LEARNED_WORDS[keys_lower]
            if target_lang == 'hebrew':
                corrected = self.keys_to_hebrew(keys)
            else:
                corrected = keys_lower
            self.log(f"  -> Learned word! Fixing to {target_lang}: '{corrected}'")
            self.log_word(keys, self.get_screen_word(keys), f"FIX (learned) to {target_lang}: {corrected}")
            return corrected + punctuation, target_lang
        
        # Get what's on screen based on tracked language
        screen_word = self.get_screen_word(keys)
        hebrew_version = self.keys_to_hebrew(keys)
        english_version = keys.lower()
        
        self.log(f"  Keys: '{keys}' | Screen: '{screen_word}' | Tracked: {self.tracked_language.upper()}" + (f" | Punct: '{punctuation}'" if punctuation else ""))
        
        # Skip if in ignored words
        if screen_word in IGNORED_WORDS or screen_word.lower() in IGNORED_WORDS:
            self.log(f"  -> Skipping (ignored word): '{screen_word}'")
            return None, None
        if keys in IGNORED_WORDS or keys.lower() in IGNORED_WORDS:
            self.log(f"  -> Skipping (ignored word): '{keys}'")
            return None, None
        
        if self.tracked_language == 'hebrew':
            # Screen shows Hebrew - check if it should be English
            
            # Check if it's valid Hebrew
            if self.is_valid_hebrew(hebrew_version):
                self.log(f"  -> Valid Hebrew '{hebrew_version}', no fix")
                self.log_word(keys, screen_word, "no_fix (valid hebrew)")
                return None, None
            
            # For short words, only fix if it's a common English word
            if len(keys) <= 3:
                if english_version in COMMON_SHORT_ENGLISH:
                    self.log(f"  -> '{english_version}' is common short English - FIXING!")
                    self.log_word(keys, screen_word, f"FIX to english: {english_version}")
                    return english_version + punctuation, 'english'
                else:
                    self.log(f"  -> Short word, not in common list, skipping")
                    self.log_word(keys, screen_word, "no_fix (short, not common)")
                    return None, None
            
            # Check if the keys form a valid English word
            if self.is_valid_english(english_version):
                self.log(f"  -> '{english_version}' is valid English - FIXING!")
                self.log_word(keys, screen_word, f"FIX to english: {english_version}")
                return english_version + punctuation, 'english'
            
            self.log(f"  -> Not recognized")
            self.log_word(keys, screen_word, "no_fix (not recognized)")
            return None, None
        
        else:
            # Screen shows English - check if it should be Hebrew
            
            # Check if it's valid English
            if self.is_valid_english(english_version):
                self.log(f"  -> Valid English '{english_version}', no fix")
                self.log_word(keys, screen_word, "no_fix (valid english)")
                return None, None
            
            # Skip if Hebrew version is in ignored words
            if hebrew_version in IGNORED_WORDS:
                self.log(f"  -> Skipping (ignored word): '{hebrew_version}'")
                self.log_word(keys, screen_word, "no_fix (ignored)")
                return None, None
            
            # For short words, only fix if it's a common Hebrew word
            if len(keys) <= 3:
                if hebrew_version in SHORT_HEBREW_WORDS or hebrew_version in HEBREW_WORDS:
                    self.log(f"  -> '{hebrew_version}' is common short Hebrew - FIXING!")
                    self.log_word(keys, screen_word, f"FIX to hebrew: {hebrew_version}")
                    return hebrew_version + punctuation, 'hebrew'
                else:
                    self.log(f"  -> Short word, not in common list, skipping")
                    self.log_word(keys, screen_word, "no_fix (short, not common)")
                    return None, None
            
            # Check if it would be valid Hebrew
            if self.is_valid_hebrew(hebrew_version):
                self.log(f"  -> '{hebrew_version}' is valid Hebrew - FIXING!")
                self.log_word(keys, screen_word, f"FIX to hebrew: {hebrew_version}")
                return hebrew_version + punctuation, 'hebrew'
            
            self.log(f"  -> Not recognized")
            self.log_word(keys, screen_word, "no_fix (not recognized)")
            return None, None
    
    def get_active_window_title(self):
        """Get the title of the active window"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            return win32gui.GetWindowText(hwnd)
        except:
            return ""
    
    def is_outlook(self):
        """Check if Outlook is the active window"""
        title = self.get_active_window_title().lower()
        # Check for various Outlook window titles
        outlook_indicators = ['outlook', 'new mail', 'new message', 'message (html)', 'message (rich text)', 'message (plain text)']
        return any(indicator in title for indicator in outlook_indicators)
    
    def is_word(self):
        """Check if Word is the active window"""
        title = self.get_active_window_title().lower()
        return 'word' in title or '.docx' in title or '.doc' in title
    
    def align_left_outlook(self):
        """Click align left button in Outlook using image recognition"""
        try:
            self.log(f"  Searching for Align Left button at: {self.align_left_img}")
            
            if not os.path.exists(self.align_left_img):
                self.log(f"  ERROR: Image file not found!")
                return False
            
            # Try multiple confidence levels
            location = None
            for confidence in [0.9, 0.8, 0.7, 0.6, 0.5]:
                try:
                    location = pyautogui.locateOnScreen(self.align_left_img, confidence=confidence, grayscale=True)
                    if location:
                        self.log(f"  Found at confidence {confidence}")
                        break
                except TypeError:
                    # OpenCV not installed
                    location = pyautogui.locateOnScreen(self.align_left_img)
                    break
                except Exception:
                    continue
            
            if location:
                center = pyautogui.center(location)
                pyautogui.click(center)
                self.log(f"  Clicked Align Left at {center}")
                return True
            else:
                self.log("  Align Left button not found on screen (tried confidence 0.9-0.5)")
                return False
        except Exception as e:
            self.log(f"  Error clicking Align Left: {type(e).__name__}: {e}")
            return False
    
    def align_right_outlook(self):
        """Click align right button in Outlook using image recognition"""
        try:
            self.log(f"  Searching for Align Right button at: {self.align_right_img}")
            
            if not os.path.exists(self.align_right_img):
                self.log(f"  ERROR: Image file not found!")
                return False
            
            # Try multiple confidence levels
            location = None
            for confidence in [0.9, 0.8, 0.7, 0.6, 0.5]:
                try:
                    location = pyautogui.locateOnScreen(self.align_right_img, confidence=confidence, grayscale=True)
                    if location:
                        self.log(f"  Found at confidence {confidence}")
                        break
                except TypeError:
                    # OpenCV not installed
                    location = pyautogui.locateOnScreen(self.align_right_img)
                    break
                except Exception:
                    continue
            
            if location:
                center = pyautogui.center(location)
                pyautogui.click(center)
                self.log(f"  Clicked Align Right at {center}")
                return True
            else:
                self.log("  Align Right button not found on screen (tried confidence 0.9-0.5)")
                return False
        except Exception as e:
            self.log(f"  Error clicking Align Right: {type(e).__name__}: {e}")
            return False
    
    def set_alignment(self, language, wake_editor=False):
        """Set text direction based on language - only for first word of line"""
        if self.direction_set_for_line:
            return  # Already set direction for this line

        self.log(f"  [DIRECTION] Setting to: {language}")

        ctrl_sent = False
        shift_key = None
        try:
            # Send a space then backspace to "wake up" the editor (helps Outlook)
            # Only needed when no typing happened before this call
            if wake_editor:
                keyboard.send('space')
                time.sleep(0.02)
                keyboard.send('backspace')
                time.sleep(0.02)

            if language == 'english':
                # Ctrl+Left Shift = LTR (English)
                keyboard.press('ctrl')
                ctrl_sent = True
                shift_key = 'left shift'
                keyboard.press(shift_key)
                keyboard.release(shift_key)
                shift_key = None
                self.log("  [DIRECTION] Sent Ctrl+Left Shift (LTR)")
            else:
                # Ctrl+Right Shift = RTL (Hebrew)
                keyboard.press('ctrl')
                ctrl_sent = True
                shift_key = 'right shift'
                keyboard.press(shift_key)
                keyboard.release(shift_key)
                shift_key = None
                self.log("  [DIRECTION] Sent Ctrl+Right Shift (RTL)")
            self.direction_set_for_line = True
        except Exception as e:
            self.log(f"  [DIRECTION] Error: {e}")
        finally:
            if shift_key:
                try:
                    keyboard.release(shift_key)
                except:
                    pass
            if ctrl_sent:
                try:
                    keyboard.release('ctrl')
                except:
                    pass
    
    def switch_keyboard(self, to_hebrew):
        """Switch the keyboard layout"""
        now = time.time()
        if now - self.last_switch_time < 0.3:
            return
        target = self.LANG_HEBREW if to_hebrew else self.LANG_ENGLISH
        hwnd = win32gui.GetForegroundWindow()
        win32api.PostMessage(hwnd, self.WM_INPUTLANGCHANGEREQUEST, 0, target)
        self.last_switch_time = now
    
    def fix_word(self, corrected, target_lang, original_keys, is_first_word=False):
        """Erase the wrong word and type the correct one"""
        if self.is_fixing:
            return
        
        # Safety check: if correction equals what's on screen, skip
        screen_word = self.get_screen_word(original_keys)
        if corrected == screen_word:
            self.log(f"  -> Skip fix (already correct): '{corrected}'")
            self.current_word_keys = ""
            return
        
        # Capitalize first letter if first word of line and English
        if is_first_word and target_lang == 'english' and len(corrected) > 0:
            corrected = corrected[0].upper() + corrected[1:]
        
        self.is_fixing = True
        original_lang = self.tracked_language
        
        try:
            try:
                self.block_input(True)
            except:
                pass
            
            # Step 1: Delete the word + space using backspace
            delete_count = len(screen_word) + 1  # +1 for the space
            for _ in range(delete_count):
                keyboard.send('backspace')
            time.sleep(0.05)
            
            # Step 2: Switch keyboard to target language
            target_is_hebrew = (target_lang == 'hebrew')
            self.switch_keyboard(to_hebrew=target_is_hebrew)
            time.sleep(0.05)
            
            # Step 3: Type corrected word + space
            # Normalize Caps Lock to OFF before writing - keyboard.write() can leave it
            # stuck when BlockInput is active and GetKeyState returns stale state
            caps_was_on = bool(ctypes.windll.user32.GetKeyState(0x14) & 0x0001)
            if caps_was_on:
                keyboard.send('caps lock')
                time.sleep(0.02)
            for char in corrected:
                keyboard.write(char)
            keyboard.send('space')
            if caps_was_on:
                keyboard.send('caps lock')
                time.sleep(0.02)
            
            # Step 4: Set text direction (after typing, only for first word)
            # Must unblock input for Ctrl+Shift to reach the application
            if is_first_word:
                try:
                    self.block_input(False)
                except:
                    pass
                time.sleep(0.05)
                self.set_alignment(target_lang)
                time.sleep(0.05)
                try:
                    self.block_input(True)
                except:
                    pass
            
            # Step 5: Update tracked language to match what we switched to
            self.tracked_language = target_lang
            self.log(f"  -> Language now: {self.tracked_language.upper()}")
            
            # Step 6: Store fix info for undo detection
            self.last_fix_info = {
                'original_screen': screen_word,
                'corrected': corrected,
                'original_lang': original_lang,
                'target_lang': target_lang
            }
            
            # Step 7: Block delay
            block_delay_sec = CONFIG['block_delay_ms'] / 1000.0
            time.sleep(block_delay_sec)
            
        finally:
            try:
                self.block_input(False)
            except:
                pass
            
            self.is_fixing = False
            self.current_word_keys = ""
    
    def add_to_ignore_list(self, word):
        """Add a word to the ignore list permanently"""
        ignore_file = os.path.join(self.script_dir, 'ignore_words.txt')
        try:
            # Add to in-memory set (global variable)
            IGNORED_WORDS.add(word.lower())
            IGNORED_WORDS.add(word)
            
            # Add to file
            with open(ignore_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{word}")
            
            self.log(f"  [UNDO] Added '{word}' to ignore list")
        except Exception as e:
            self.log(f"  [UNDO] Error adding to ignore list: {e}")
    
    def handle_undo(self):
        """Handle Ctrl+` - undo last fix"""
        if not self.last_fix_info:
            return
        if self.is_fixing:
            return

        original_word = self.last_fix_info['original_screen']
        original_lang = self.last_fix_info['original_lang']
        corrected_word = self.last_fix_info['corrected']

        self.is_fixing = True
        self.current_word_keys = ""

        try:
            # Add the original word to ignore list
            self.add_to_ignore_list(original_word)

            # Wait for user to release hotkey, then release ctrl
            time.sleep(0.3)
            keyboard.release('left ctrl')
            keyboard.release('right ctrl')
            time.sleep(0.1)

            # Erase the corrected word + space
            delete_count = len(corrected_word) + 1  # +1 for the space
            for _ in range(delete_count):
                keyboard.send('backspace')
            time.sleep(0.1)

            # Switch language back to original
            self.switch_keyboard(to_hebrew=(original_lang == 'hebrew'))
            self.tracked_language = original_lang
            time.sleep(0.1)

            # Reset direction flag and set alignment (no wake_editor - typing will wake it)
            self.direction_set_for_line = False
            self.set_alignment(original_lang, wake_editor=False)

            # Type back the original word + space
            for char in original_word:
                keyboard.write(char)
            keyboard.send('space')

            # Clear the fix info
            self.last_fix_info = None
            self.last_unfixed_word = None
            self.log(f"  [UNDO] Restored '{original_word}'")
        finally:
            self.is_fixing = False
            self.current_word_keys = ""
    
    def handle_force_fix(self):
        """Handle Ctrl+` on unfixed word - force fix and learn"""
        if not self.last_unfixed_word:
            return
        if self.is_fixing:
            return

        keys = self.last_unfixed_word
        screen_word = self.get_screen_word(keys)
        was_first_word = self.last_unfixed_was_first

        # Determine target language (opposite of current)
        if self.tracked_language == 'english':
            target_lang = 'hebrew'
            corrected = self.keys_to_hebrew(keys)
        else:
            target_lang = 'english'
            corrected = keys.lower()

        self.log(f"  [FORCE FIX] '{screen_word}' -> '{corrected}' ({target_lang})")
        self.log(f"  [FORCE FIX] keys='{keys}' len={len(keys)} was_first={was_first_word}")

        # Save to learned words
        self.save_learned_word(keys.lower(), target_lang)

        self.is_fixing = True
        self.current_word_keys = ""

        try:
            # Wait for user to release hotkey
            time.sleep(0.3)
            keyboard.release('left ctrl')
            keyboard.release('right ctrl')
            time.sleep(0.1)

            # Erase the word + space (cursor is after space, so delete word length + 1 space)
            delete_count = len(screen_word) + 1
            self.log(f"  [FORCE FIX] Deleting {delete_count} characters (keys len={len(keys)})")
            for _ in range(delete_count):
                keyboard.send('backspace')
                time.sleep(0.02)  # Small delay between backspaces
            time.sleep(0.1)

            # Switch keyboard
            self.switch_keyboard(to_hebrew=(target_lang == 'hebrew'))
            self.tracked_language = target_lang
            time.sleep(0.1)

            # Type corrected word + space
            for char in corrected:
                keyboard.write(char)
            keyboard.send('space')

            # Set text direction only if this was the first word of line
            if was_first_word:
                time.sleep(0.05)
                self.set_alignment(target_lang, wake_editor=True)
                self.direction_set_for_line = True

            # Clear tracking
            self.last_unfixed_word = None
            self.last_unfixed_was_first = False
            self.log(f"  [FORCE FIX] Done - word learned for future")
        finally:
            self.is_fixing = False
            self.current_word_keys = ""
    
    def save_learned_word(self, keys, target_lang):
        """Save a word to learned_words.txt"""
        global LEARNED_WORDS
        learned_file = os.path.join(self.script_dir, 'learned_words.txt')
        try:
            with open(learned_file, 'a', encoding='utf-8') as f:
                f.write(f"{keys},{target_lang}\n")
            LEARNED_WORDS[keys.lower()] = target_lang
            self.log(f"  [LEARNED] Added '{keys}' -> {target_lang}")
            self.file_log(f"LEARNED: {keys} -> {target_lang}")
        except Exception as e:
            self.log(f"  [Error saving learned word: {e}]")
    
    def release_all_keys(self):
        """Release all potentially stuck modifier keys and normalize Caps Lock.
        Hotkey: Ctrl+Alt+R  (if Ctrl is stuck, just press Alt+R)"""
        for key_name in ['left ctrl', 'right ctrl', 'left shift', 'right shift', 'left alt', 'right alt']:
            try:
                keyboard.release(key_name)
            except:
                pass
        # Turn off Caps Lock if it's currently on
        try:
            if ctypes.windll.user32.GetKeyState(0x14) & 0x0001:
                keyboard.send('caps lock')
        except:
            pass
        # Reset internal modifier state
        self.ctrl_pressed = False
        self.alt_pressed = False
        self.shift_pressed = False
        self.log("  [RELEASE ALL] Modifier keys released, Caps Lock normalized")

    def toggle_language(self):
        """Toggle the tracked language"""
        old_lang = self.tracked_language
        if self.tracked_language == 'english':
            self.tracked_language = 'hebrew'
        else:
            self.tracked_language = 'english'
        self.log(f"  [Language toggled to: {self.tracked_language.upper()}]")
        self.file_log(f"LANG_TOGGLE: {old_lang} -> {self.tracked_language} (Alt+Shift)")
    
    def set_language(self, lang):
        """Manually set the tracked language"""
        old_lang = self.tracked_language
        self.tracked_language = lang
        self.log(f"  [Language manually set to: {self.tracked_language.upper()}]")
        self.file_log(f"LANG_SET: {old_lang} -> {self.tracked_language} (manual)")
    
    def on_key_press(self, key):
        if self.is_fixing:
            return
        
        try:
            # Check if window changed - clear buffer and reset first word tracking
            current_window = win32gui.GetForegroundWindow()
            if current_window != self.last_active_window:
                if self.current_word_keys:
                    print(f"  [Window changed - CLEARED buffer: '{self.current_word_keys}']")
                self.current_word_keys = ""
                self.last_active_window = current_window
                # Reset first word tracking - user switched to new window
                self.is_first_word_of_line = True
                self.direction_set_for_line = False
            
            # Track Ctrl for undo detection
            if key == Key.ctrl_l or key == Key.ctrl_r:
                self.ctrl_pressed = True
                return
            
            # Detect Ctrl+` (backtick) for undo OR force-fix - vk code 192
            if self.ctrl_pressed and hasattr(key, 'vk') and key.vk == 192:
                if self.last_fix_info:
                    # Undo last fix
                    threading.Thread(target=self.handle_undo, daemon=True).start()
                elif self.last_unfixed_word:
                    # Force-fix unfixed word
                    threading.Thread(target=self.handle_force_fix, daemon=True).start()
                return
            
            # Track Alt and Shift for language toggle detection
            if key == Key.alt_l or key == Key.alt_r:
                self.alt_pressed = True
                return
            if key == Key.shift_l or key == Key.shift_r:
                if self.alt_pressed:
                    # Alt+Shift detected - toggle language
                    self.toggle_language()
                    self.current_word_keys = ""
                    self.alt_pressed = False
                    return
                self.shift_pressed = True
                return
            
            # Get physical key from vk code
            key_char = None
            if hasattr(key, 'vk') and key.vk is not None:
                vk = key.vk
                if vk in self.VK_TO_ENGLISH:
                    key_char = self.VK_TO_ENGLISH[vk]
            
            if key_char:
                self.current_word_keys += key_char
            
            elif key == Key.enter:
                # Process current word first if any
                if len(self.current_word_keys) >= 2:
                    corrected, target = self.analyze_and_fix(self.current_word_keys)
                    
                    if corrected and target:
                        original_keys = self.current_word_keys
                        is_first = self.is_first_word_of_line
                        if self.is_first_word_of_line:
                            self.is_first_word_of_line = False
                        self.last_unfixed_word = None  # Clear - word was fixed
                        self.last_fix_info = None  # Will be set in fix_word
                        threading.Thread(
                            target=self.fix_word,
                            args=(corrected, target, original_keys, is_first),
                            daemon=True
                        ).start()
                    else:
                        # Word was NOT fixed - store for potential force-fix
                        self.last_unfixed_word = self.current_word_keys
                        self.last_unfixed_was_first = self.is_first_word_of_line  # Remember if it was first
                        self.last_fix_info = None  # No fix happened
                        # Word was correct - but still set direction if first word
                        if self.is_first_word_of_line:
                            detected_lang = self.detect_language(self.current_word_keys)
                            if detected_lang:
                                self.log(f"  [DIRECTION] First word correct, setting direction for: {detected_lang}")
                                # Run in thread with delay like fix_word does
                                def set_dir(lang):
                                    time.sleep(0.05)
                                    self.set_alignment(lang, wake_editor=True)
                                    self.direction_set_for_line = True
                                threading.Thread(target=set_dir, args=(detected_lang,), daemon=True).start()
                                self.last_unfixed_was_first = False  # Direction was set
                            else:
                                # Language not detected, direction NOT set - keep last_unfixed_was_first = True
                                pass
                            self.is_first_word_of_line = False
                        self.current_word_keys = ""
                else:
                    self.current_word_keys = ""
                
                # New line - reset first word tracking
                self.is_first_word_of_line = True
                self.direction_set_for_line = False
            
            elif key in [Key.space, Key.tab]:
                if len(self.current_word_keys) >= 2:
                    corrected, target = self.analyze_and_fix(self.current_word_keys)
                    
                    if corrected and target:
                        original_keys = self.current_word_keys
                        is_first = self.is_first_word_of_line
                        if self.is_first_word_of_line:
                            self.is_first_word_of_line = False
                        self.last_unfixed_word = None  # Clear - word was fixed
                        self.last_fix_info = None  # Will be set in fix_word
                        threading.Thread(
                            target=self.fix_word,
                            args=(corrected, target, original_keys, is_first),
                            daemon=True
                        ).start()
                    else:
                        # Word was NOT fixed - store for potential force-fix
                        self.last_unfixed_word = self.current_word_keys
                        self.last_unfixed_was_first = self.is_first_word_of_line  # Remember if it was first
                        self.last_fix_info = None  # No fix happened
                        # Word was correct - but still set direction if first word
                        if self.is_first_word_of_line:
                            detected_lang = self.detect_language(self.current_word_keys)
                            if detected_lang:
                                self.log(f"  [DIRECTION] First word correct, setting direction for: {detected_lang}")
                                # Run in thread with delay like fix_word does
                                def set_dir(lang):
                                    time.sleep(0.05)
                                    self.set_alignment(lang, wake_editor=True)
                                    self.direction_set_for_line = True
                                threading.Thread(target=set_dir, args=(detected_lang,), daemon=True).start()
                                self.last_unfixed_was_first = False  # Direction was set
                            else:
                                # Language not detected, direction NOT set - keep last_unfixed_was_first = True
                                pass
                            self.is_first_word_of_line = False
                        self.current_word_keys = ""
                else:
                    self.current_word_keys = ""
            
            elif key == Key.backspace:
                if self.current_word_keys:
                    self.current_word_keys = self.current_word_keys[:-1]
            
            elif key in [Key.left, Key.right, Key.up, Key.down, Key.home, Key.end]:
                self.current_word_keys = ""
            
        except Exception as e:
            if self.debug:
                print(f"Error: {e}")
    
    def on_key_release(self, key):
        """Track key releases for Alt+Shift detection"""
        if key == Key.alt_l or key == Key.alt_r:
            self.alt_pressed = False
        if key == Key.shift_l or key == Key.shift_r:
            self.shift_pressed = False
        if key == Key.ctrl_l or key == Key.ctrl_r:
            self.ctrl_pressed = False
    
    def start_session_monitor(self):
        """Monitor for session unlock events and reset to English"""
        import ctypes
        from ctypes import wintypes, Structure, WINFUNCTYPE, c_int, c_void_p, c_wchar_p, POINTER
        
        self.file_log("SESSION_MONITOR: Starting...")
        
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        wtsapi32 = ctypes.windll.wtsapi32
        
        # Set up DefWindowProcW properly for 64-bit
        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.DefWindowProcW.restype = ctypes.c_longlong
        
        # Window messages
        WM_WTSSESSION_CHANGE = 0x02B1
        WM_POWERBROADCAST = 0x0218
        
        # Session notification types
        WTS_SESSION_LOCK = 0x7
        WTS_SESSION_UNLOCK = 0x8
        WTS_SESSION_LOGON = 0x5
        WTS_SESSION_LOGOFF = 0x6
        
        # Power events
        PBT_APMRESUMEAUTOMATIC = 0x0012
        PBT_APMRESUMESUSPEND = 0x0007
        
        # Notification flags
        NOTIFY_FOR_THIS_SESSION = 0
        
        # Define WNDCLASSEXW structure
        class WNDCLASSEXW(Structure):
            _fields_ = [
                ("cbSize", wintypes.UINT),
                ("style", wintypes.UINT),
                ("lpfnWndProc", c_void_p),
                ("cbClsExtra", c_int),
                ("cbWndExtra", c_int),
                ("hInstance", wintypes.HINSTANCE),
                ("hIcon", wintypes.HICON),
                ("hCursor", wintypes.HICON),
                ("hbrBackground", wintypes.HBRUSH),
                ("lpszMenuName", c_wchar_p),
                ("lpszClassName", c_wchar_p),
                ("hIconSm", wintypes.HICON),
            ]
        
        # Use platform-appropriate types for WPARAM and LPARAM (64-bit on 64-bit Windows)
        if ctypes.sizeof(c_void_p) == 8:  # 64-bit
            WNDPROC = WINFUNCTYPE(ctypes.c_longlong, ctypes.c_void_p, wintypes.UINT, ctypes.c_ulonglong, ctypes.c_longlong)
        else:  # 32-bit
            WNDPROC = WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        
        def wnd_proc(hwnd, msg, wparam, lparam):
            try:
                if msg == WM_WTSSESSION_CHANGE:
                    self.file_log(f"SESSION_MONITOR: WM_WTSSESSION_CHANGE received, wparam={wparam}")
                    if wparam == WTS_SESSION_UNLOCK:
                        self.log("  [Session UNLOCK detected - resetting to ENGLISH]")
                        self.file_log(f"SESSION_UNLOCK: resetting tracked_language to english (was {self.tracked_language})")
                        self.tracked_language = 'english'
                        self.current_word_keys = ""
                    elif wparam == WTS_SESSION_LOGON:
                        self.log("  [Session LOGON detected - resetting to ENGLISH]")
                        self.file_log(f"SESSION_LOGON: resetting tracked_language to english (was {self.tracked_language})")
                        self.tracked_language = 'english'
                        self.current_word_keys = ""
                    elif wparam == WTS_SESSION_LOCK:
                        self.file_log(f"SESSION_LOCK: screen locked")
                
                elif msg == WM_POWERBROADCAST:
                    self.file_log(f"SESSION_MONITOR: WM_POWERBROADCAST received, wparam={wparam}")
                    if wparam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
                        self.log("  [Wake from sleep detected - resetting to ENGLISH]")
                        self.file_log(f"WAKE_FROM_SLEEP: resetting tracked_language to english (was {self.tracked_language})")
                        self.tracked_language = 'english'
                        self.current_word_keys = ""
            except Exception as e:
                self.file_log(f"SESSION_MONITOR: Error in wnd_proc: {e}")
            
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)
        
        # Keep reference to prevent garbage collection
        self._wnd_proc_cb = WNDPROC(wnd_proc)
        
        # Get module handle
        hInstance = kernel32.GetModuleHandleW(None)
        
        # Use unique class name
        import random
        class_name = f"SessionMonitor_{random.randint(10000, 99999)}"
        
        # Register window class using WNDCLASSEXW
        wc = WNDCLASSEXW()
        wc.cbSize = ctypes.sizeof(WNDCLASSEXW)
        wc.lpfnWndProc = ctypes.cast(self._wnd_proc_cb, c_void_p)
        wc.lpszClassName = class_name
        wc.hInstance = hInstance
        
        atom = user32.RegisterClassExW(ctypes.byref(wc))
        if not atom:
            error = kernel32.GetLastError()
            self.log(f"  [Warning: Could not register session monitor class, error={error}]")
            self.file_log(f"SESSION_MONITOR: FAILED to register class, error={error}")
            return
        
        # Create a hidden window (not message-only, WTS needs a real window)
        hwnd = user32.CreateWindowExW(
            0,                      # dwExStyle
            class_name,             # lpClassName  
            "SessionMonitor",       # lpWindowName
            0,                      # dwStyle (no visible style)
            0, 0, 0, 0,            # x, y, width, height
            None,                   # hWndParent (None = top-level)
            None,                   # hMenu
            hInstance,              # hInstance
            None                    # lpParam
        )
        
        if not hwnd:
            error = kernel32.GetLastError()
            self.log(f"  [Warning: Could not create session monitor window, error={error}]")
            self.file_log(f"SESSION_MONITOR: FAILED to create window, error={error}")
            return
        
        # Register for session notifications
        if not wtsapi32.WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION):
            error = kernel32.GetLastError()
            self.log(f"  [Warning: Could not register for session notifications, error={error}]")
            self.file_log(f"SESSION_MONITOR: FAILED to register session notifications, error={error}")
            user32.DestroyWindow(hwnd)
            return
        
        self.log("  [Session monitor started - will reset to English on unlock/logon]")
        self.file_log("SESSION_MONITOR: Started successfully, listening for unlock/logon events")
        
        # Message loop
        msg = wintypes.MSG()
        while self.is_running:
            # Use GetMessage for better CPU efficiency, but with timeout via PeekMessage
            if user32.PeekMessageW(ctypes.byref(msg), hwnd, 0, 0, 1):  # PM_REMOVE = 1
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.1)
        
        # Cleanup
        wtsapi32.WTSUnRegisterSessionNotification(hwnd)
        user32.DestroyWindow(hwnd)
    
    def on_mouse_click(self, x, y, button, pressed):
        """Clear buffer on mouse click and reset first word tracking"""
        if pressed:
            if self.current_word_keys:
                print(f"  [Mouse click - CLEARED buffer: '{self.current_word_keys}']")
                self.current_word_keys = ""
            # Reset first word tracking - user clicked somewhere new
            self.is_first_word_of_line = True
            self.direction_set_for_line = False
    
    def run(self):
        def quit_app():
            self.log("\nQuitting...")
            self.is_running = False
        
        def set_english():
            self.set_language('english')
        
        def set_hebrew():
            self.set_language('hebrew')
        
        def show_about():
            self.show_about_dialog()
        
        keyboard.add_hotkey('ctrl+alt+q', quit_app)
        keyboard.add_hotkey('ctrl+alt+e', set_english)
        keyboard.add_hotkey('ctrl+alt+h', set_hebrew)
        keyboard.add_hotkey('ctrl+alt+a', show_about)
        keyboard.add_hotkey('ctrl+alt+r', self.release_all_keys)
        
        # Start power monitor in background thread
        session_thread = threading.Thread(target=self.start_session_monitor, daemon=True)
        session_thread.start()
        
        # Start mouse listener in background
        mouse_listener = pynput_mouse.Listener(on_click=self.on_mouse_click)
        mouse_listener.start()
        
        with pynput_keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release
        ) as listener:
            while self.is_running:
                time.sleep(0.1)
            listener.stop()
        
        mouse_listener.stop()
    
    def show_about_dialog(self):
        """Show About dialog with copyright info"""
        import ctypes
        
        about_text = """Hebrew-English Auto Keyboard Switcher
Version 3.1.64

Automatically fixes words typed in the wrong keyboard language.

Created by Eitan Pollack

Hotkeys:
  Alt+Shift      - Toggle language
  Ctrl+`         - Undo fix / Force-fix
  Ctrl+Alt+E     - Force English
  Ctrl+Alt+H     - Force Hebrew
  Ctrl+Alt+A     - About
  Ctrl+Alt+R     - Release stuck keys
  Ctrl+Alt+Q     - Quit"""
        
        # Show message box (Windows)
        ctypes.windll.user32.MessageBoxW(
            0, 
            about_text, 
            "About Auto Switcher", 
            0x40  # MB_ICONINFORMATION
        )


def open_demo_on_first_run():
    """Open demo.html on first run"""
    import webbrowser
    
    # Get script directory
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check for first-run marker
    marker_file = os.path.join(script_dir, '.first_run_done')
    demo_file = os.path.join(script_dir, 'demo.html')
    
    if not os.path.exists(marker_file):
        # First run - open demo if it exists
        if os.path.exists(demo_file):
            # Use proper file URL format for Windows
            demo_url = 'file:///' + demo_file.replace('\\', '/')
            try:
                webbrowser.open(demo_url)
            except:
                pass
        
        # Create marker file
        try:
            with open(marker_file, 'w') as f:
                f.write('Demo shown')
        except:
            pass

def main():
    # Check for first run and open demo
    open_demo_on_first_run()
    
    debug_mode = '--debug' in sys.argv or '-d' in sys.argv  # Use --debug to enable logging
    switcher = HebrewEnglishSwitcher(debug=debug_mode)
    switcher.run()


if __name__ == "__main__":
    main()
