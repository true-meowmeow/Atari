# Module: config and shared runtime flags.
# Main: APP_DIR/DATA_PATH/SETTINGS_PATH, DEBUG_* flags, CAPTURE_OVERLAY_ACTIVE.
# Example: from atari.core import config; config.DATA_PATH

from pathlib import Path

# ---- Persistence ----
APP_DIR = Path.home() / ".macro_gui"
APP_DIR.mkdir(parents=True, exist_ok=True)
DATA_PATH = APP_DIR / "records.json"
SETTINGS_PATH = APP_DIR / "settings.json"

# ---- Debug / runtime tuning ----
DEBUG_STOP_WORD_OCR = False   # True = print OCR debug to console
STOP_WORD_POLL_SEC = 10.0     # poll interval for stop-word checks (seconds)
FOCUS_POLL_SEC = 0.5          # how often to restore focus to the game

# Flag for blocking global hotkeys during key capture.
CAPTURE_OVERLAY_ACTIVE = False
