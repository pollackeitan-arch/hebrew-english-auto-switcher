================================================================================
    Hebrew-English Auto Keyboard Switcher v3.1.63
================================================================================

================================================================================
ENGLISH
================================================================================

WHAT IS THIS PROGRAM?
---------------------
Automatically fixes words typed in the wrong keyboard layout AND
automatically aligns text based on the first word's language.

Example 1: You want to type "שלום" but forgot to switch to Hebrew.
           You type "akuo" and press space.
           The program fixes it to "שלום" AND aligns right.

Example 2: You want to type "hello" but keyboard is on Hebrew.
           You type "יקךךם" and press space.
           The program fixes it to "Hello" AND aligns left.


WHAT'S NEW IN v3.1.63
---------------------
- FORCE-FIX & LEARN: Press Ctrl+` on unfixed words to fix them AND 
  teach the program to recognize them in the future
- CONTRACTIONS: Now recognizes i'm, don't, can't, won't, etc.
- AUTO-CAPITALIZE: First word of line corrected to English is capitalized
- DEBUG MODE: Run with run_debug.bat for troubleshooting (no logging by default)


HOW IT WORKS
------------
- Tracks language changes by detecting Alt+Shift
- Starts assuming English (your login language)
- Resets to English when waking from sleep
- After first word is typed/fixed, sets alignment automatically
- Works in Outlook and Word!


HOW TO RUN
----------
Option 1: Double-click "run.bat" (normal mode, no logging)

Option 2: Double-click "run_debug.bat" (debug mode, shows all actions)

Option 3: Start automatically with Windows
          Double-click "add_to_startup.bat" ONE TIME


HOTKEYS
-------
Alt+Shift     - Toggle language (auto-detected)
Ctrl+`        - UNDO last fix (adds word to ignore list)
              - OR FORCE-FIX unfixed word (adds to learned words)
Ctrl+Alt+E    - Force set to English
Ctrl+Alt+H    - Force set to Hebrew
Ctrl+Alt+Q    - Quit the program


USING CTRL+` (BACKTICK)
-----------------------
This key has TWO functions:

1. UNDO: After a word was auto-fixed incorrectly
   - Reverts to original word
   - Adds word to ignore_words.txt (won't fix again)

2. FORCE-FIX: After typing a word that WASN'T fixed (but should be)
   - Fixes the word to the other language
   - Adds word to learned_words.txt (will fix automatically next time)

Example of Force-Fix:
   You type "akuo" but it's not in the dictionary, so it stays "akuo"
   Press Ctrl+` and it becomes "שלום" AND learns for next time


FILES
-----
config.ini         - Settings (block delay)
ignore_words.txt   - Words to never change
learned_words.txt  - Words you taught via Ctrl+` (auto-fix next time)
hebrew_words.txt   - Hebrew dictionary (470,000 words)
align_left.png     - Image of align-left button (for Outlook)
align_right.png    - Image of align-right button (for Outlook)


OUTLOOK BUTTON IMAGES
---------------------
If alignment doesn't work in Outlook, the button images may not match.
Take new screenshots of the buttons and replace:
- align_left.png
- align_right.png


================================================================================
עברית
================================================================================

מה התוכנה עושה?
---------------
מתקנת אוטומטית מילים שהוקלדו בשפה הלא נכונה
וגם מיישרת את הטקסט אוטומטית לפי השפה.


מה חדש ב-v3.1.63
-----------------
- תיקון ולמידה: לחצו Ctrl+` על מילה שלא תוקנה כדי לתקן אותה 
  וללמד את התוכנה לזהות אותה בעתיד
- קיצורים: מזהה עכשיו i'm, don't, can't וכו'
- אות גדולה: מילה ראשונה בשורה שמתוקנת לאנגלית מקבלת אות גדולה


איך זה עובד
-----------
- עוקב אחרי שינויי שפה על ידי זיהוי Alt+Shift
- מתחיל בהנחה שהשפה היא אנגלית
- אחרי המילה הראשונה בכל שורה - מיישר אוטומטית


מקשי קיצור
----------
Alt+Shift     - החלפת שפה (זיהוי אוטומטי)
Ctrl+`        - ביטול תיקון אחרון (מוסיף לרשימת התעלמות)
              - או תיקון ידני + למידה (למילים שלא תוקנו)
Ctrl+Alt+E    - קביעה ידנית לאנגלית
Ctrl+Alt+H    - קביעה ידנית לעברית
Ctrl+Alt+Q    - יציאה מהתוכנה


================================================================================
