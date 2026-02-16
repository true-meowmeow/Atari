# Module: project navigation map.
# Main: PROJECT_MAP string with module guide.
# Example: print(PROJECT_MAP)

PROJECT_MAP = """
Entry point
- main.py -> atari.runtime.app.main()

Core modules
- atari/core/config.py: paths + runtime flags
- atari/core/models.py: dataclasses + action serialization + OCR helpers
- atari/core/win32.py: Win32 helpers + SendInput wrappers
- atari/core/geometry.py: screen/rect helpers

UI / runtime
- atari/ui/main_window.py: MainWindow, all widgets and UI logic
- atari/ui/overlays.py: selection/capture overlays + spec_to_pretty
- atari/runtime/player.py: MacroPlayer thread and playback logic
- atari/runtime/hotkeys.py: global hotkeys + interval meter

Typical edits
- UI text/layout/buttons: atari/ui/main_window.py
- Macro playback behavior: atari/runtime/player.py
- Action models or serialization: atari/core/models.py
- OCR behavior: atari/core/models.py (ocr_text_in_rect)
- Win32 focus / input: atari/core/win32.py

Compatibility layer
- src/*.py re-export modules from atari/* to keep legacy imports working
"""
