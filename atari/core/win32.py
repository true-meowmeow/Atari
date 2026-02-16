# Module: Win32 helpers and SendInput wrappers.
# Main: resolve_hwnd_by_exe, _win_activate_hwnd, _win_send_key_batch.
# Example: from atari.core.win32 import resolve_hwnd_by_exe

import os
import sys
import ctypes
from ctypes import wintypes
from typing import List, Optional, Tuple

from PySide6.QtCore import QPoint, QRect

_user32 = None
_kernel32 = None
_GetDpiForWindow = None
_QueryFullProcessImageNameW = None

def _is_windows() -> bool:
    return sys.platform.startswith("win")


# ---------- Win32 helpers (без pywin32) ----------
if _is_windows():
    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    _user32.GetWindowThreadProcessId.restype = wintypes.DWORD

    _user32.WindowFromPoint.argtypes = [wintypes.POINT]
    _user32.WindowFromPoint.restype = wintypes.HWND

    _user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    _user32.GetAncestor.restype = wintypes.HWND

    _user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _user32.IsWindowVisible.restype = wintypes.BOOL

    _user32.IsIconic.argtypes = [wintypes.HWND]
    _user32.IsIconic.restype = wintypes.BOOL

    _user32.GetForegroundWindow.argtypes = []
    _user32.GetForegroundWindow.restype = wintypes.HWND

    _user32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    _user32.GetClientRect.restype = wintypes.BOOL

    _user32.ClientToScreen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
    _user32.ClientToScreen.restype = wintypes.BOOL

    # GetDpiForWindow доступен на Win10+, но на всякий случай — через getattr
    _GetDpiForWindow = getattr(_user32, "GetDpiForWindow", None)
    if _GetDpiForWindow:
        _GetDpiForWindow.argtypes = [wintypes.HWND]
        _GetDpiForWindow.restype = wintypes.UINT

    _kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    _kernel32.OpenProcess.restype = wintypes.HANDLE

    _kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    _kernel32.CloseHandle.restype = wintypes.BOOL

    # ---- foreground / focus ----
    _user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    _user32.SetForegroundWindow.restype = wintypes.BOOL

    _user32.BringWindowToTop.argtypes = [wintypes.HWND]
    _user32.BringWindowToTop.restype = wintypes.BOOL

    _user32.SetFocus.argtypes = [wintypes.HWND]
    _user32.SetFocus.restype = wintypes.HWND

    _user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    _user32.ShowWindow.restype = wintypes.BOOL

    _user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
    _user32.AttachThreadInput.restype = wintypes.BOOL

    _kernel32.GetCurrentThreadId.argtypes = []
    _kernel32.GetCurrentThreadId.restype = wintypes.DWORD

    # ---- SendInput / MapVirtualKey ----
    _user32.MapVirtualKeyW.argtypes = [wintypes.UINT, wintypes.UINT]
    _user32.MapVirtualKeyW.restype = wintypes.UINT


    _QueryFullProcessImageNameW = getattr(_kernel32, "QueryFullProcessImageNameW", None)
    if _QueryFullProcessImageNameW:
        _QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        _QueryFullProcessImageNameW.restype = wintypes.BOOL

    GA_ROOT = 2
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

def _hwnd_int(h) -> int:
    # ctypes HWND(NULL) -> None, это нормально
    return int(h) if h else 0

def _win_hwnd_from_point(x: int, y: int) -> int:
    if not _is_windows():
        return 0
    pt = wintypes.POINT(int(x), int(y))
    hwnd = _hwnd_int(_user32.WindowFromPoint(pt))
    if hwnd:
        hwnd = _hwnd_int(_user32.GetAncestor(hwnd, GA_ROOT))
    return hwnd



def _win_pid_from_hwnd(hwnd: int) -> int:
    if not _is_windows() or not hwnd:
        return 0
    pid = wintypes.DWORD(0)
    _user32.GetWindowThreadProcessId(wintypes.HWND(int(hwnd)), ctypes.byref(pid))
    return int(pid.value)


def _win_exe_from_pid(pid: int) -> str:
    if not _is_windows() or pid <= 0 or not _QueryFullProcessImageNameW:
        return ""
    h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, wintypes.DWORD(int(pid)))
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buf))
        ok = _QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
        return buf.value if ok else ""
    finally:
        _kernel32.CloseHandle(h)


def _win_dpi_for_hwnd(hwnd: int) -> int:
    if not _is_windows() or not hwnd or not _GetDpiForWindow:
        return 96
    try:
        dpi = int(_GetDpiForWindow(wintypes.HWND(int(hwnd))))
        return dpi if dpi > 0 else 96
    except Exception:
        return 96


def _win_client_rect_screen_dip(hwnd: int) -> Optional[QRect]:
    """Клиентская область окна в экранных координатах, В Qt-DIP (не физ пиксели)."""
    if not _is_windows() or not hwnd:
        return None

    # если окно свернуто — считаем что "нет процесса/окна"
    if _user32.IsIconic(wintypes.HWND(int(hwnd))):
        return None

    rc = wintypes.RECT()
    if not _user32.GetClientRect(wintypes.HWND(int(hwnd)), ctypes.byref(rc)):
        return None

    # top-left (0,0) -> screen
    tl = wintypes.POINT(0, 0)
    br = wintypes.POINT(rc.right, rc.bottom)
    if not _user32.ClientToScreen(wintypes.HWND(int(hwnd)), ctypes.byref(tl)):
        return None
    if not _user32.ClientToScreen(wintypes.HWND(int(hwnd)), ctypes.byref(br)):
        return None

    # Win32 br — "конец", для QRect с bottomRight-инклюзивным: -1
    left_phys, top_phys = int(tl.x), int(tl.y)
    right_phys, bot_phys = int(br.x) - 1, int(br.y) - 1

    dpi = _win_dpi_for_hwnd(hwnd)
    scale = 96.0 / float(dpi)  # физ -> DIP

    left = int(round(left_phys * scale))
    top = int(round(top_phys * scale))
    right = int(round(right_phys * scale))
    bottom = int(round(bot_phys * scale))

    r = QRect(QPoint(left, top), QPoint(right, bottom)).normalized()
    if not r.isValid() or r.width() < 5 or r.height() < 5:
        return None
    return r


def _win_enum_top_windows() -> List[int]:
    if not _is_windows():
        return []
    res: List[int] = []

    EnumWindows = _user32.EnumWindows
    EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
    EnumWindows.restype = wintypes.BOOL

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, lparam):
        res.append(int(hwnd))
        return True

    EnumWindows(cb, 0)
    return res


def resolve_hwnd_by_exe(exe_path: str) -> int:
    """Пытаемся найти живое видимое окно процесса по exe path."""
    if not _is_windows():
        return 0
    exe_path = (exe_path or "").strip()
    if not exe_path:
        return 0

    exe_norm = os.path.normcase(exe_path)

    # 1) foreground — лучший кандидат
    fg = _hwnd_int(_user32.GetForegroundWindow())
    if fg:
        pid = _win_pid_from_hwnd(fg)
        if pid:
            pexe = _win_exe_from_pid(pid)
            if pexe and os.path.normcase(pexe) == exe_norm:
                if _user32.IsWindowVisible(wintypes.HWND(fg)) and not _user32.IsIconic(wintypes.HWND(fg)):
                    return fg

    # 2) иначе перебираем все top-level окна
    for hwnd in _win_enum_top_windows():
        if not hwnd:
            continue
        if not _user32.IsWindowVisible(wintypes.HWND(hwnd)):
            continue
        if _user32.IsIconic(wintypes.HWND(hwnd)):
            continue

        pid = _win_pid_from_hwnd(hwnd)
        if not pid:
            continue
        pexe = _win_exe_from_pid(pid)
        if pexe and os.path.normcase(pexe) == exe_norm:
            return hwnd

    return 0


def resolve_bound_base_rect_dip(exe_path: str) -> Optional[QRect]:
    hwnd = resolve_hwnd_by_exe(exe_path)
    if not hwnd:
        return None
    return _win_client_rect_screen_dip(hwnd)

# ---------- Win32 SendInput keyboard (SCANCODE) + Foreground activate ----------
if _is_windows():
    # pointer-sized ULONG_PTR
    if ctypes.sizeof(ctypes.c_void_p) == 8:
        _ULONG_PTR = ctypes.c_ulonglong
    else:
        _ULONG_PTR = ctypes.c_ulong

    INPUT_KEYBOARD = 1

    KEYEVENTF_EXTENDEDKEY = 0x0001
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_SCANCODE = 0x0008

    MAPVK_VK_TO_VSC = 0

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", _ULONG_PTR),
        ]

    class _INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("union", _INPUT_UNION)]

    _user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
    _user32.SendInput.restype = wintypes.UINT

    SW_RESTORE = 9

    # RU -> "US key" (чтобы запись, сделанная на RU раскладке, тоже могла отработать)
    _RU_TO_US = {
        "ё": "`",
        "й": "q", "ц": "w", "у": "e", "к": "r", "е": "t", "н": "y", "г": "u", "ш": "i", "щ": "o", "з": "p", "х": "[", "ъ": "]",
        "ф": "a", "ы": "s", "в": "d", "а": "f", "п": "g", "р": "h", "о": "j", "л": "k", "д": "l", "ж": ";", "э": "'",
        "я": "z", "ч": "x", "с": "c", "м": "v", "и": "b", "т": "n", "ь": "m", "б": ",", "ю": ".", ".": "/",
    }

    # VK_OEM mapping (US layout physical keys)
    _VK_OEM = {
        ";": 0xBA,  # VK_OEM_1
        "=": 0xBB,  # VK_OEM_PLUS
        ",": 0xBC,  # VK_OEM_COMMA
        "-": 0xBD,  # VK_OEM_MINUS
        ".": 0xBE,  # VK_OEM_PERIOD
        "/": 0xBF,  # VK_OEM_2
        "`": 0xC0,  # VK_OEM_3
        "[": 0xDB,  # VK_OEM_4
        "\\": 0xDC, # VK_OEM_5
        "]": 0xDD,  # VK_OEM_6
        "'": 0xDE,  # VK_OEM_7
    }

    _VK_SPECIAL = {
        "Escape": 0x1B,
        "Enter": 0x0D,
        "Tab": 0x09,
        "Backspace": 0x08,
        "Delete": 0x2E,
        "Insert": 0x2D,
        "Space": 0x20,
        "Home": 0x24,
        "End": 0x23,
        "PageUp": 0x21,
        "PageDown": 0x22,
        "Up": 0x26,
        "Down": 0x28,
        "Left": 0x25,
        "Right": 0x27,
        "CapsLock": 0x14,
        "PrintScreen": 0x2C,
        "Pause": 0x13,

        # modifiers (лучше “левые”, чтобы игры стабильнее ловили)
        "Ctrl": 0xA2,   # VK_LCONTROL
        "Shift": 0xA0,  # VK_LSHIFT
        "Alt": 0xA4,    # VK_LMENU
        "Meta": 0x5B,   # VK_LWIN
    }

    _VK_EXTENDED = {
        0x21, 0x22, 0x23, 0x24,  # PgUp PgDn End Home
        0x25, 0x26, 0x27, 0x28,  # arrows
        0x2D, 0x2E,              # Insert Delete
        0x5B, 0x5C,              # Win keys
        0x2C,                    # PrintScreen
    }

    def _win_vk_from_name(name: str) -> tuple[int, bool]:
        n = (name or "").strip()
        if not n:
            return (0, False)

        # normalize mods spelling
        low = n.casefold()
        if low == "control":
            n = "Ctrl"
        elif low == "shift":
            n = "Shift"
        elif low == "alt":
            n = "Alt"
        elif low in ("meta", "win", "cmd"):
            n = "Meta"

        # F keys
        if n.startswith("F") and n[1:].isdigit():
            k = int(n[1:])
            if 1 <= k <= 24:
                vk = 0x70 + (k - 1)
                return (vk, False)

        # specials
        if n in _VK_SPECIAL:
            vk = _VK_SPECIAL[n]
            return (vk, vk in _VK_EXTENDED)

        # single character: latin/digit/oem, plus RU->US mapping
        if len(n) == 1:
            ch = n
            ch_low = ch.lower()

            # RU char -> US key
            if ch_low in _RU_TO_US:
                ch_low = _RU_TO_US[ch_low]

            if "a" <= ch_low <= "z":
                return (ord(ch_low.upper()), False)
            if "0" <= ch_low <= "9":
                return (ord(ch_low), False)
            if ch_low in _VK_OEM:
                vk = _VK_OEM[ch_low]
                return (vk, False)

        # unknown
        return (0, False)

    def _win_send_key_by_name(name: str, is_down: bool) -> bool:
        vk, is_ext = _win_vk_from_name(name)
        if not vk:
            return False

        sc = int(_user32.MapVirtualKeyW(int(vk), MAPVK_VK_TO_VSC) or 0)

        # ✅ Fallback: модификаторы иногда дают scancode=0 через MapVirtualKey
        if sc <= 0:
            if vk == 0xA0:       # VK_LSHIFT
                sc = 0x2A        # LShift scancode
            elif vk == 0xA1:     # VK_RSHIFT
                sc = 0x36
            elif vk == 0xA2:     # VK_LCONTROL
                sc = 0x1D
            elif vk == 0xA3:     # VK_RCONTROL (обычно extended)
                sc = 0x1D
                is_ext = True
            elif vk == 0xA4:     # VK_LMENU (Alt)
                sc = 0x38
            elif vk == 0xA5:     # VK_RMENU (AltGr, extended)
                sc = 0x38
                is_ext = True
            elif vk == 0x5B:     # VK_LWIN (Meta)
                sc = 0x5B
                is_ext = True
            elif vk == 0x5C:     # VK_RWIN
                sc = 0x5C
                is_ext = True
            else:
                return False

        flags = KEYEVENTF_SCANCODE
        if is_ext:
            flags |= KEYEVENTF_EXTENDEDKEY
        if not is_down:
            flags |= KEYEVENTF_KEYUP

        inp = INPUT()
        inp.type = INPUT_KEYBOARD
        inp.union.ki = KEYBDINPUT(wVk=0, wScan=sc, dwFlags=flags, time=0, dwExtraInfo=0)

        sent = int(_user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) or 0)
        return sent == 1

    def _win_make_key_input(name: str, is_down: bool) -> Optional[INPUT]:
        vk, is_ext = _win_vk_from_name(name)
        if not vk:
            return None

        sc = int(_user32.MapVirtualKeyW(int(vk), MAPVK_VK_TO_VSC) or 0)

        # fallback для модификаторов / спецклавиш
        if sc <= 0:
            if vk == 0xA0:       # VK_LSHIFT
                sc = 0x2A
            elif vk == 0xA1:     # VK_RSHIFT
                sc = 0x36
            elif vk == 0xA2:     # VK_LCONTROL
                sc = 0x1D
            elif vk == 0xA3:     # VK_RCONTROL
                sc = 0x1D
                is_ext = True
            elif vk == 0xA4:     # VK_LMENU (Alt)
                sc = 0x38
            elif vk == 0xA5:     # VK_RMENU
                sc = 0x38
                is_ext = True
            elif vk == 0x5B:     # VK_LWIN
                sc = 0x5B
                is_ext = True
            elif vk == 0x5C:     # VK_RWIN
                sc = 0x5C
                is_ext = True
            else:
                return None

        flags = KEYEVENTF_SCANCODE
        if is_ext:
            flags |= KEYEVENTF_EXTENDEDKEY
        if not is_down:
            flags |= KEYEVENTF_KEYUP

        inp = INPUT()
        inp.type = INPUT_KEYBOARD
        inp.union.ki = KEYBDINPUT(
            wVk=0,
            wScan=sc,
            dwFlags=flags,
            time=0,
            dwExtraInfo=0
        )
        return inp

    def _win_send_key_batch(names: List[str], is_down: bool) -> bool:
        inputs: List[INPUT] = []
        for n in names:
            inp = _win_make_key_input(n, is_down)
            if inp is not None:
                inputs.append(inp)

        if not inputs:
            return False

        arr = (INPUT * len(inputs))(*inputs)
        sent = int(_user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT)) or 0)
        return sent == len(inputs)


    def _win_activate_hwnd(hwnd: int) -> bool:
        if not hwnd:
            return False
        h = wintypes.HWND(int(hwnd))

        try:
            if _user32.IsIconic(h):
                _user32.ShowWindow(h, SW_RESTORE)
        except Exception:
            pass

        try:
            fg = _hwnd_int(_user32.GetForegroundWindow())
            if fg == hwnd:
                return True

            # attach thread input trick
            cur_tid = int(_kernel32.GetCurrentThreadId())

            pid = wintypes.DWORD(0)
            fg_tid = int(_user32.GetWindowThreadProcessId(wintypes.HWND(fg), ctypes.byref(pid)) or 0)

            pid2 = wintypes.DWORD(0)
            target_tid = int(_user32.GetWindowThreadProcessId(h, ctypes.byref(pid2)) or 0)

            if fg_tid:
                _user32.AttachThreadInput(cur_tid, fg_tid, True)
            if target_tid:
                _user32.AttachThreadInput(cur_tid, target_tid, True)

            _user32.BringWindowToTop(h)
            _user32.SetForegroundWindow(h)
            _user32.SetFocus(h)

            if target_tid:
                _user32.AttachThreadInput(cur_tid, target_tid, False)
            if fg_tid:
                _user32.AttachThreadInput(cur_tid, fg_tid, False)

        except Exception:
            try:
                _user32.SetForegroundWindow(h)
            except Exception:
                pass

        return _hwnd_int(_user32.GetForegroundWindow()) == hwnd



if not _is_windows():
    def _win_vk_from_name(name: str) -> tuple[int, bool]:
        return (0, False)

    def _win_send_key_by_name(name: str, is_down: bool) -> bool:
        return False

    def _win_make_key_input(name: str, is_down: bool):
        return None

    def _win_send_key_batch(names: List[str], is_down: bool) -> bool:
        return False

    def _win_activate_hwnd(hwnd: int) -> bool:
        return False
