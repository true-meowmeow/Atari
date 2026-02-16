# Module: macro playback thread and signals.
# Main: MacroPlayer, PlayerSignals.
# Example: from atari.runtime.player import MacroPlayer

import random
import threading
import time
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import QObject, QRect, QThread, Signal

from atari.core.config import DEBUG_STOP_WORD_OCR, FOCUS_POLL_SEC, STOP_WORD_POLL_SEC
from atari.core.geometry import rect_to_rel, rel_to_rect, virtual_geometry
from atari.core.models import (
    Action, AreaAction, BaseAreaAction, Delay, KeyAction, MouseButtonName,
    Record, WaitAction, WaitEventAction, WordAreaAction,
    DEFAULT_TRIGGER, action_from_dict, action_to_display, is_ocr_available, normalize_trigger,
)
from atari.core.win32 import (
    _hwnd_int, _is_windows, _user32, _win_activate_hwnd, _win_client_rect_screen_dip,
    _win_dpi_for_hwnd, _win_send_key_batch, _win_send_key_by_name, resolve_hwnd_by_exe,
)

# ---- Playback Thread ----
class PlayerSignals(QObject):
    status = Signal(str)
    progress = Signal(int, int)
    current = Signal(str)
    cycle = Signal(str)
    action_row = Signal(int)

    # NEW:
    action_error = Signal(int, str)  # row, message  (РєСЂР°СЃРёРј СЃС‚СЂРѕРєСѓ, РїСЂРёС‡РёРЅР°)
    action_ok = Signal(int)  # row           (СЃРЅРёРјР°РµРј РєСЂР°СЃРЅС‹Р№)
    paused = Signal(bool, str)  # is_paused, reason

    finished = Signal()


class MacroPlayer(QThread):
    def __init__(self, record: Record, bound_exe: str = "", stop_word_cfg: Optional[dict] = None,
                 stop_word_enabled_event: Optional[threading.Event] = None, start_index: int = 0):
        super().__init__()
        self.record = record

        self.signals = PlayerSignals()
        self._stop = threading.Event()
        self._pause = threading.Event()
        self._pause.set()
        self._pause_lock = threading.Lock()
        self.end_state = "done"  # done | stopped | error
        self.end_reason = ""

        self._bound_exe = (bound_exe or "").strip()
        self._stop_word = WordAreaAction.from_dict(stop_word_cfg) if isinstance(stop_word_cfg, dict) else None
        self._stop_word_interval = float(STOP_WORD_POLL_SEC)
        if isinstance(stop_word_cfg, dict):
            try:
                interval = float(stop_word_cfg.get("interval_sec", STOP_WORD_POLL_SEC))
            except Exception:
                interval = float(STOP_WORD_POLL_SEC)
            if interval <= 0:
                interval = float(STOP_WORD_POLL_SEC)
            self._stop_word_interval = interval
        self._stop_word_enabled_event = stop_word_enabled_event
        self._start_index = max(0, int(start_index))

        # NEW: stop-word worker (С‡С‚РѕР±С‹ OCR РЅРµ РґС‘СЂРіР°Р» РѕСЃРЅРѕРІРЅРѕР№ РїРѕС‚РѕРє)
        self._sw_lock = threading.Lock()
        self._sw_base: Optional[QRect] = None
        self._sw_dpr: float = 1.0
        self._sw_thread: Optional[threading.Thread] = None


    def _mark_end(self, state: str, reason: str):
        if getattr(self, "end_state", "done") == "done":
            self.end_state = state
            self.end_reason = reason

    def stop(self, reason: str = "Остановлено пользователем"):
        self._mark_end("stopped", reason)
        self._stop.set()
        self._pause.set()  # С‡С‚РѕР±С‹ РЅРµ Р·Р°РІРёСЃ РІ РїР°СѓР·Рµ

    def pause(self, reason: str = "Пауза"):
        with self._pause_lock:
            if not self._pause.is_set():
                return
            self._pause.clear()
        self.signals.paused.emit(True, reason)

    def resume(self):
        with self._pause_lock:
            if self._pause.is_set():
                return
            self._pause.set()
        self.signals.paused.emit(False, "")

    def is_paused(self) -> bool:
        return not self._pause.is_set()

    def run(self):

        try:
            # Lazy import here (avoid import errors on GUI init)
            try:
                from pynput.mouse import Controller as MouseController, Button
                from pynput.keyboard import Controller as KeyboardController, Key
            except Exception as ex:
                self.signals.status.emit(f"Ошибка pynput: {ex}")
                self.signals.finished.emit()
                return

            mouse = MouseController()
            kb = KeyboardController()

            base_area: Optional[QRect] = None
            last_base: Optional[QRect] = None
            ocr_available = is_ocr_available(force_refresh=False)

            def _sw_set_snapshot(base: Optional[QRect], dpr: float):
                with self._sw_lock:
                    self._sw_base = QRect(base) if (base and base.isValid()) else None
                    self._sw_dpr = float(dpr) if (dpr and dpr > 0) else 1.0

            def _start_stopword_worker():
                if (not self._stop_word) or (not ocr_available):
                    return

                def worker():
                    next_check = 0.0
                    while not self._stop.is_set():
                        # enabled?
                        if self._stop_word_enabled_event is not None and not self._stop_word_enabled_event.is_set():
                            time.sleep(0.05)
                            continue

                        now = time.perf_counter()
                        if now < next_check:
                            time.sleep(0.05)
                            continue
                        next_check = now + float(self._stop_word_interval)

                        with self._sw_lock:
                            b = QRect(self._sw_base) if self._sw_base else None
                            dpr = float(self._sw_dpr)

                        if b and b.isValid():
                            found = self._stop_word.resolve_target_rect_global(
                                b,
                                debug=bool(DEBUG_STOP_WORD_OCR),
                                debug_tag="[STOP_WORD] ",
                                dpr_override=dpr,  # NEW
                            )
                            ok = bool(found and found.isValid() and found.width() > 2 and found.height() > 2)
                            if ok:
                                w = getattr(self._stop_word, "word", "")
                                self._mark_end("stopped", f"Остановлено: найдено стоп-слово '{w}'")
                                self.signals.status.emit(f"Стоп-слово найдено: '{w}' — остановка.")
                                self._stop.set()
                                self._pause.set()
                                break

                        time.sleep(0.05)

                self._sw_thread = threading.Thread(target=worker, daemon=True)
                self._sw_thread.start()

            next_stop_check = time.perf_counter() + 1.0

            next_focus_check = 0.0

            def _ensure_focus(force: bool = False):
                nonlocal next_focus_check
                if not self._bound_exe or not _is_windows():
                    return
                now = time.perf_counter()
                if (not force) and now < next_focus_check:
                    return
                next_focus_check = now + float(FOCUS_POLL_SEC)

                hwnd = resolve_hwnd_by_exe(self._bound_exe)
                if not hwnd:
                    return

                fg = _hwnd_int(_user32.GetForegroundWindow())
                if force or fg != hwnd:
                    _win_activate_hwnd(hwnd)

            def _resolve_runtime_base() -> tuple[Optional[QRect], float, int]:
                """(rect_dip, dpr, hwnd)"""
                if self._bound_exe and _is_windows():
                    hwnd = resolve_hwnd_by_exe(self._bound_exe)
                    if not hwnd:
                        return None, 1.0, 0
                    rect = _win_client_rect_screen_dip(hwnd)
                    dpr = float(_win_dpi_for_hwnd(hwnd)) / 96.0
                    if dpr <= 0:
                        dpr = 1.0
                    return rect, dpr, hwnd
                # fallback
                return None, 1.0, 0

            # --- init base when record has no base_area (С‡С‚РѕР±С‹ rel-РєРѕРѕСЂРґРёРЅР°С‚С‹ СЂР°Р±РѕС‚Р°Р»Рё) ---
            runtime_rect, current_dpr, runtime_hwnd = _resolve_runtime_base()

            if self._bound_exe and (runtime_rect is None or not runtime_rect.isValid()):
                msg = "База привязана, но окно/процесс не найден — остановка, чтобы не кликать по экрану."
                self._mark_end("error", msg)
                self.signals.status.emit(msg)
                self.signals.finished.emit()
                return

            base_area = runtime_rect if (runtime_rect and runtime_rect.isValid()) else virtual_geometry()
            last_base = QRect(base_area) if base_area and base_area.isValid() else None
            current_area = base_area

            # snapshot + СЃС‚Р°СЂС‚ С„РѕРЅРѕРІРѕР№ РїСЂРѕРІРµСЂРєРё СЃС‚РѕРї-СЃР»РѕРІР°
            _sw_set_snapshot(base_area, current_dpr)
            _start_stopword_worker()


            # NEW: РїСЂРё СЃС‚Р°СЂС‚Рµ вЂ” СЃСЂР°Р·Сѓ РїРѕРґРЅСЏС‚СЊ РѕРєРЅРѕ РёРіСЂС‹
            _ensure_focus(force=True)

            def _maybe_update_base(current_area_ref: Optional[QRect]) -> Tuple[Optional[QRect], Optional[QRect]]:
                nonlocal last_base
                rb_rect, rb_dpr, rb_hwnd = _resolve_runtime_base()
                if rb_rect and rb_rect.isValid():
                    if last_base and last_base.isValid() and current_area_ref and current_area_ref.isValid():
                        if (rb_rect.left() != last_base.left() or rb_rect.top() != last_base.top() or
                                rb_rect.width() != last_base.width() or rb_rect.height() != last_base.height()):
                            rx1, ry1, rx2, ry2 = rect_to_rel(last_base, current_area_ref)
                            current_area_ref = rel_to_rect(rb_rect, rx1, ry1, rx2, ry2)
                    last_base = QRect(rb_rect)
                    return rb_rect, current_area_ref
                return last_base, current_area_ref

            # NEW: С„РѕРєСѓСЃ + СЃРЅР°РїС€РѕС‚
            _ensure_focus()
            _sw_set_snapshot(base_area, current_dpr)

            # build action objects
            actions: List[Action] = []
            for ad in self.record.actions:
                if isinstance(ad, dict):
                    actions.append(action_from_dict(ad))

            has_ocr_actions = bool(self._stop_word) or any(
                isinstance(a, (WordAreaAction, WaitEventAction))
                for a in actions
            )
            if (not ocr_available) and has_ocr_actions:
                self.signals.status.emit(
                    "OCR недоступен: Tesseract не найден. OCR-действия будут пропущены."
                )

            start_index = self._start_index
            if start_index < 0 or start_index >= len(actions):
                start_index = 0

            def _normalize_key_press_mode(mode: Any) -> str:
                return "long" if str(mode or "").strip().lower() == "long" else "normal"

            def _normalize_delay_cfg(raw: Any, default_sec: float) -> Dict[str, Any]:
                try:
                    if isinstance(raw, Delay):
                        d = raw
                    elif isinstance(raw, dict):
                        d = Delay.from_dict(raw)
                    else:
                        d = Delay("fixed", float(default_sec), float(default_sec))
                except Exception:
                    d = Delay("fixed", float(default_sec), float(default_sec))

                mode = "range" if str(getattr(d, "mode", "fixed")).lower() == "range" else "fixed"
                try:
                    a = max(0.0, float(getattr(d, "a", default_sec)))
                except Exception:
                    a = max(0.0, float(default_sec))
                try:
                    b = max(0.0, float(getattr(d, "b", a)))
                except Exception:
                    b = a
                if mode == "fixed":
                    b = a
                return {"mode": mode, "a": float(a), "b": float(b)}

            def _sample_delay_cfg(raw: Any, default_sec: float) -> float:
                cfg = _normalize_delay_cfg(raw, default_sec)
                a = float(cfg.get("a", default_sec))
                b = float(cfg.get("b", a))
                if cfg.get("mode") == "range":
                    lo = min(a, b)
                    hi = max(a, b)
                    return random.uniform(lo, hi)
                return a

            def _normalize_long_press_item(raw: Any) -> Optional[Dict[str, Any]]:
                if not isinstance(raw, dict):
                    return None

                trigger_raw = raw.get("trigger")
                hold_raw = raw.get("hold")
                activate_mode_raw = raw.get("activate_mode", "after_prev")
                start_delay_raw = raw.get("start_delay")

                # Backward compatibility: legacy long entries stored regular KeyAction dicts.
                if not isinstance(trigger_raw, dict):
                    try:
                        legacy_action = action_from_dict(raw)
                    except Exception:
                        legacy_action = None
                    if isinstance(legacy_action, KeyAction):
                        trigger_raw = {
                            "kind": legacy_action.kind,
                            "keys": list(getattr(legacy_action, "keys", []) or []),
                            "mouse_button": getattr(legacy_action, "mouse_button", None),
                        }
                        hold_raw = legacy_action.delay.to_dict()

                if not isinstance(trigger_raw, dict):
                    return None

                return {
                    "trigger": normalize_trigger(trigger_raw),
                    "hold": _normalize_delay_cfg(hold_raw, 0.2),
                    "activate_mode": "from_start" if str(activate_mode_raw).strip().lower() == "from_start" else "after_prev",
                    "start_delay": _normalize_delay_cfg(start_delay_raw, 0.0),
                }

            def _key_long_action_items(a_key: KeyAction) -> List[Dict[str, Any]]:
                cfg = getattr(a_key, "long_press", {})
                if not isinstance(cfg, dict):
                    return []
                raw = cfg.get("actions", [])
                if not isinstance(raw, list):
                    return []
                out: List[Dict[str, Any]] = []
                for ad in raw:
                    item = _normalize_long_press_item(ad)
                    if item is not None:
                        out.append(item)
                return out

            def _build_inline_action_seq(raw_actions: List[Dict[str, Any]]) -> List[Action]:
                seq: List[Action] = []
                for ad in (raw_actions or []):
                    if not isinstance(ad, dict):
                        continue
                    try:
                        seq.append(action_from_dict(ad))
                    except Exception:
                        continue
                return seq

            def _estimate_action_progress_item(it: Action, depth: int = 0) -> int:
                if depth >= 8:
                    return 0
                if isinstance(it, KeyAction):
                    mode = _normalize_key_press_mode(getattr(it, "press_mode", "normal"))
                    if mode == "long":
                        return len(_key_long_action_items(it))
                    try:
                        mult = int(getattr(it, "multiplier", 1))
                    except Exception:
                        mult = 1
                    return max(1, mult)
                if isinstance(it, (WaitAction, WaitEventAction)):
                    return 1
                return 0

            def _estimate_extra_progress_items(seq_actions: List[Action]) -> int:
                extra = 0
                for it in seq_actions:
                    extra += _estimate_action_progress_item(it)
                return extra

            total = max(1, _estimate_extra_progress_items(actions))

            def k_to_pynput(name: str):
                n = (name or "").strip()
                low = n.casefold()
                if low == "ctrl":
                    n = "Ctrl"
                elif low == "shift":
                    n = "Shift"
                elif low == "alt":
                    n = "Alt"
                elif low == "meta":
                    n = "Meta"

                name_norm = n  # <-- РІР°Р¶РЅРѕ!

                specials = {
                    "Escape": Key.esc,
                    "Enter": Key.enter,
                    "Tab": Key.tab,
                    "Backspace": Key.backspace,
                    "Delete": Key.delete,
                    "Insert": Key.insert,
                    "Space": Key.space,
                    "Home": Key.home,
                    "End": Key.end,
                    "PageUp": Key.page_up,
                    "PageDown": Key.page_down,
                    "Up": Key.up,
                    "Down": Key.down,
                    "Left": Key.left,
                    "Right": Key.right,
                    "CapsLock": Key.caps_lock,
                    "PrintScreen": Key.print_screen,
                    "Pause": Key.pause,
                    "Ctrl": Key.ctrl,
                    "Shift": Key.shift,
                    "Alt": Key.alt,
                    "Meta": Key.cmd,
                }
                if name_norm in specials:
                    return specials[name_norm]
                if name_norm.startswith("F") and name_norm[1:].isdigit():
                    n = int(name_norm[1:])
                    # pynput supports Key.f1..f20
                    attr = f"f{n}"
                    return getattr(Key, attr, None)
                # letters/digits
                if len(name_norm) == 1:
                    # pynput wants lowercase for letters
                    if name_norm.isalpha():
                        return name_norm.lower()
                    return name_norm
                return None

            def perform_trigger(spec: dict):
                spec = normalize_trigger(spec)
                kind = spec["kind"]
                keys = list(spec.get("keys") or [])
                mb = spec.get("mouse_button")

                if kind == "mouse":
                    if mb not in ("left", "middle", "right"):
                        mb = "left"

                    mods_only = [k for k in keys if k in ("Shift", "Ctrl", "Alt", "Meta")]

                    for mname in mods_only:
                        if self._stop.is_set():
                            return
                        _kbd_down(mname)
                        _sleep_checked(0.005)

                    _sleep_checked(0.020)
                    if not self._stop.is_set():
                        _ensure_focus()
                        click(mb)

                    for mname in reversed(mods_only):
                        try:
                            _kbd_up(mname)
                        except Exception:
                            pass
                    return

                # kind == "keys"
                press_combo(keys)

            def perform_trigger_times(spec: dict, times: int, delay: Optional[Delay] = None):
                reps = max(1, int(times))
                for i in range(reps):
                    if self._stop.is_set():
                        return
                    if delay is not None:
                        t = delay.sample()
                        if t > 0:
                            if not _sleep_checked(t):
                                return
                    perform_trigger(spec)
                    if delay is None and i + 1 < reps:
                        if not _sleep_checked(0.03):
                            return

            def _press_hold_trigger(spec: dict) -> Dict[str, Any]:
                spec = normalize_trigger(spec)
                kind = spec["kind"]
                keys = list(spec.get("keys") or [])
                mb = spec.get("mouse_button")

                mod_names = {"Shift", "Ctrl", "Alt", "Meta"}
                mod_order = {"Ctrl": 0, "Alt": 1, "Shift": 2, "Meta": 3}

                if kind == "mouse":
                    if mb not in ("left", "middle", "right"):
                        mb = "left"
                    mods_only = [k for k in keys if k in mod_names]
                    mods_only.sort(key=lambda x: mod_order.get(x, 99))

                    for mname in mods_only:
                        if self._stop.is_set():
                            break
                        _kbd_down(mname)
                        _sleep_checked(0.001)

                    if mb == "left":
                        btn = Button.left
                    elif mb == "middle":
                        btn = Button.middle
                    else:
                        btn = Button.right

                    _ensure_focus()
                    try:
                        mouse.press(btn)
                    except Exception:
                        pass
                    return {"kind": "mouse", "mods": list(mods_only), "btn": btn}

                # kind == "keys"
                if not keys:
                    return {"kind": "noop"}

                mods = [k for k in keys if k in mod_names]
                normals = [k for k in keys if k not in mod_names]
                mods.sort(key=lambda x: mod_order.get(x, 99))
                downs = mods + normals
                ups = list(reversed(normals)) + list(reversed(mods))

                _ensure_focus()
                if _is_windows():
                    ok_down = _win_send_key_batch(downs, True)
                    if not ok_down:
                        for k in downs:
                            _kbd_down(k)
                            _sleep_checked(0.001)
                else:
                    for k in downs:
                        _kbd_down(k)
                        _sleep_checked(0.001)

                return {"kind": "keys", "ups": list(ups)}

            def _release_hold_trigger(handle: Optional[Dict[str, Any]]):
                if not isinstance(handle, dict):
                    return
                kind = str(handle.get("kind", "")).lower()

                if kind == "mouse":
                    btn = handle.get("btn")
                    mods = list(handle.get("mods") or [])
                    if btn is not None:
                        try:
                            mouse.release(btn)
                        except Exception:
                            pass
                    for mname in reversed(mods):
                        try:
                            _kbd_up(str(mname))
                            _sleep_checked(0.001)
                        except Exception:
                            pass
                    return

                if kind == "keys":
                    ups = [str(k) for k in (handle.get("ups") or [])]
                    if _is_windows():
                        ok_up = _win_send_key_batch(ups, False) if ups else True
                        if not ok_up:
                            for k in ups:
                                try:
                                    _kbd_up(k)
                                    _sleep_checked(0.001)
                                except Exception:
                                    pass
                    else:
                        for k in ups:
                            try:
                                _kbd_up(k)
                                _sleep_checked(0.001)
                            except Exception:
                                pass

            def smooth_move_to(x: int, y: int, duration: float):
                duration = max(0.0, float(duration))
                if duration <= 0.001:
                    mouse.position = (x, y)
                    return
                start = mouse.position
                sx, sy = start[0], start[1]
                dx = x - sx
                dy = y - sy
                steps = max(1, int(duration * 120))
                step_sleep = duration / steps

                for i in range(1, steps + 1):
                    if self._stop.is_set():
                        return
                    if not _wait_if_paused():
                        return

                    t = i / steps
                    tt = t * t * (3 - 2 * t)  # ease in-out
                    nx = int(sx + dx * tt)
                    ny = int(sy + dy * tt)
                    mouse.position = (nx, ny)

                    if not _sleep_checked(step_sleep):
                        return

            def click(btn_name: MouseButtonName):
                if btn_name == "left":
                    b = Button.left
                elif btn_name == "middle":
                    b = Button.middle
                else:
                    b = Button.right
                mouse.press(b)
                _sleep_checked(0.01)
                mouse.release(b)

            def _wait_if_paused() -> bool:
                while not self._stop.is_set() and not self._pause.is_set():
                    time.sleep(0.05)
                return not self._stop.is_set()

            def _sleep_checked(sec: float) -> bool:
                """Sleep РєСѓСЃРєР°РјРё, СѓС‡РёС‚С‹РІР°СЏ stop() Рё PAUSE (РІРѕ РІСЂРµРјСЏ РїР°СѓР·С‹ РІСЂРµРјСЏ РЅРµ "СЃСЉРµРґР°РµС‚СЃСЏ")."""
                remaining = max(0.0, float(sec))
                while remaining > 0:
                    if self._stop.is_set():
                        return False
                    if not _wait_if_paused():
                        return False

                    chunk = min(0.01, remaining)
                    t0 = time.perf_counter()

                    time.sleep(chunk)
                    remaining -= (time.perf_counter() - t0)
                return True

            def _kbd_down(name: str) -> bool:
                if _is_windows():
                    if _win_send_key_by_name(name, True):
                        return True
                pk = k_to_pynput(name)
                if pk is not None:
                    kb.press(pk)
                    return True
                return False

            def _kbd_up(name: str) -> bool:
                if _is_windows():
                    if _win_send_key_by_name(name, False):
                        return True
                pk = k_to_pynput(name)
                if pk is not None:
                    kb.release(pk)
                    return True
                return False

            def press_combo(keys: List[str]):
                _ensure_focus()

                mod_names = {"Shift", "Ctrl", "Alt", "Meta"}
                mod_order = {"Ctrl": 0, "Alt": 1, "Shift": 2, "Meta": 3}

                mods = [k for k in keys if k in mod_names]
                normals = [k for k in keys if k not in mod_names]

                mods.sort(key=lambda x: mod_order.get(x, 99))

                # С…РѕС‚РёРј: mods down -> normals down -> normals up -> mods up
                downs = mods + normals
                ups = list(reversed(normals)) + list(reversed(mods))

                KEY_HOLD_TIME = 0.060  # <--- РµСЃР»Рё РёРіСЂР° РєР°РїСЂРёР·РЅР°СЏ, РјРѕР¶РЅРѕ 0.08-0.12
                AFTER_UP_DELAY = 0.005

                # Windows: РїР°С‡РєРѕР№ (СЃР°РјРѕРµ РІР°Р¶РЅРѕРµ РґР»СЏ РёРіСЂ)
                if _is_windows():
                    ok_down = _win_send_key_batch(downs, True)
                    if not ok_down:
                        # fallback РЅР° СЃС‚Р°СЂС‹Р№ СЃРїРѕСЃРѕР±, РµСЃР»Рё РІРґСЂСѓРі
                        for k in downs:
                            _kbd_down(k)
                            _sleep_checked(0.001)

                    if not _sleep_checked(KEY_HOLD_TIME):
                        # РµСЃР»Рё РѕСЃС‚Р°РЅРѕРІРёР»Рё РІРѕ РІСЂРµРјСЏ СѓРґРµСЂР¶Р°РЅРёСЏ вЂ” РіР°СЂР°РЅС‚РёСЂРѕРІР°РЅРЅРѕ РѕС‚РїСѓСЃС‚РёРј
                        if _is_windows():
                            _win_send_key_batch(ups, False)
                        else:
                            for k in ups:
                                try:
                                    _kbd_up(k)
                                except Exception:
                                    pass
                        return

                    # РѕС‚РїСѓСЃРєР°РµРј: СЃРЅР°С‡Р°Р»Р° E, РїРѕС‚РѕРј ShiftF
                    ok_up = _win_send_key_batch(ups, False)
                    if not ok_up:
                        for k in ups:
                            try:
                                _kbd_up(k)
                                _sleep_checked(0.001)
                            except Exception:
                                pass

                    _sleep_checked(AFTER_UP_DELAY)
                    return

                # РЅРµ-Windows fallback
                for k in downs:
                    _kbd_down(k)
                    _sleep_checked(0.001)

                if not _sleep_checked(KEY_HOLD_TIME):
                    for k in ups:
                        try:
                            _kbd_up(k)
                        except Exception:
                            pass
                    return

                for k in ups:
                    try:
                        _kbd_up(k)
                        _sleep_checked(0.001)
                    except Exception:
                        pass

                _sleep_checked(AFTER_UP_DELAY)

            current_area: Optional[QRect] = base_area
            move_mouse_ready = bool(getattr(self.record, "move_mouse", True))
            cycles_left = int(self.record.repeat.count)
            infinite = bool(self.record.repeat.enabled and cycles_left == 0)

            def cycle_label():
                if not self.record.repeat.enabled:
                    return "Цикл: выключен"
                if infinite:
                    return "Цикл: бесконечно"
                return f"Цикл: осталось {cycles_left}"

            self.signals.status.emit("Проигрывание запущено")
            self.signals.cycle.emit(cycle_label())

            done = 0
            cycle_num = 0
            restart_from_beginning = False
            ocr_skip_notice_emitted = False

            def _emit_ocr_skip_notice():
                nonlocal ocr_skip_notice_emitted
                if ocr_skip_notice_emitted:
                    return
                ocr_skip_notice_emitted = True
                self.signals.current.emit("OCR недоступен: OCR-зависимые действия пропускаются.")

            def _on_fail_post_mode(a_word: WordAreaAction) -> str:
                mode = str(getattr(a_word, "on_fail_post_mode", "none") or "none")
                if mode not in ("none", "stop", "repeat"):
                    return "none"
                return mode

            def _apply_on_fail_post_mode(a_word: WordAreaAction) -> bool:
                nonlocal restart_from_beginning, start_index

                mode = _on_fail_post_mode(a_word)
                if mode == "stop":
                    stop_msg = (
                        "Остановлено: завершены дополнительные действия "
                        "в ветке «Если не найдено» (режим «Остановиться»)."
                    )
                    self._mark_end("stopped", stop_msg)
                    self.signals.status.emit(stop_msg)
                    self._stop.set()
                    return False

                if mode == "repeat":
                    restart_from_beginning = True
                    start_index = 0
                    self.signals.current.emit(
                        "Повтор с начала: завершены дополнительные действия "
                        "в ветке «Если не найдено» (режим «Повторить»)."
                    )
                return True

            def _execute_key_long_press_items(items: List[Dict[str, Any]], owner_idx: int) -> bool:
                nonlocal done, total

                if not items:
                    self.signals.current.emit("Список действий для продолжительного нажатия пуст.")
                    return True

                plan: List[Dict[str, Any]] = []
                prev_end = 0.0

                # Build an absolute timeline for each list item.
                # from_start: start time is absolute from long-action start.
                # after_prev: start time is tied to end of the previous list item.
                for idx, item in enumerate(items):
                    hold_sec = max(0.0, float(_sample_delay_cfg(item.get("hold"), 0.2)))
                    activate_mode = "from_start" if str(item.get("activate_mode", "")).lower() == "from_start" else "after_prev"
                    if activate_mode == "from_start":
                        start_sec = max(0.0, float(_sample_delay_cfg(item.get("start_delay"), 0.0)))
                    else:
                        start_sec = max(0.0, float(prev_end))
                    end_sec = start_sec + hold_sec
                    prev_end = end_sec
                    plan.append({
                        "index": int(idx),
                        "trigger": normalize_trigger(item.get("trigger", DEFAULT_TRIGGER)),
                        "start_sec": float(start_sec),
                        "end_sec": float(end_sec),
                        "hold_sec": float(hold_sec),
                        "mode": activate_mode,
                    })

                total_items = len(plan)
                self.signals.current.emit(f"Выполняю действия продолжительного нажатия: {total_items}")

                events: List[Tuple[float, int, int, str]] = []
                for p in plan:
                    idx = int(p["index"])
                    # Event sorting key:
                    # 1) time
                    # 2) item index (preserve list order ties)
                    # 3) event order (start before end for same item/time)
                    events.append((float(p["start_sec"]), idx, 0, "start"))
                    events.append((float(p["end_sec"]), idx, 1, "end"))
                events.sort(key=lambda e: (e[0], e[1], e[2]))

                active_holds: Dict[int, Dict[str, Any]] = {}
                started_at = time.perf_counter()

                try:
                    for rel_sec, idx, _ord, ev_type in events:
                        if self._stop.is_set():
                            return False
                        if restart_from_beginning:
                            return True
                        if not _wait_if_paused():
                            return False

                        target_at = started_at + float(rel_sec)
                        wait_sec = max(0.0, target_at - time.perf_counter())
                        if wait_sec > 0 and (not _sleep_checked(wait_sec)):
                            return False

                        if self._stop.is_set():
                            return False
                        if restart_from_beginning:
                            return True

                        p = plan[idx]
                        if ev_type == "start":
                            self.signals.current.emit(
                                f"Продолжительное нажатие {idx + 1}/{total_items}: старт +{p['start_sec']:.3f}s, удержание {p['hold_sec']:.3f}s"
                            )
                            active_holds[idx] = _press_hold_trigger(p["trigger"])
                            continue

                        hold_handle = active_holds.pop(idx, None)
                        _release_hold_trigger(hold_handle)
                        done += 1
                        self.signals.progress.emit(done, total)

                    return not self._stop.is_set()
                finally:
                    # Ensure no key/button remains pressed on stop/error.
                    for idx in sorted(list(active_holds.keys()), reverse=True):
                        _release_hold_trigger(active_holds.get(idx))
                    active_holds.clear()

            def _execute_inline_action(a_inline: Action, owner_idx: int) -> bool:
                nonlocal current_area, move_mouse_ready, done, total, base_area, current_dpr

                if self._stop.is_set():
                    return False

                if isinstance(a_inline, BaseAreaAction):
                    base_area = a_inline.rect()
                    current_dpr = 1.0
                    current_area = base_area
                    self.signals.current.emit(
                        f"Базовая область: {base_area.left()},{base_area.top()} → {base_area.right()},{base_area.bottom()}"
                    )
                    return True

                if isinstance(a_inline, AreaAction):
                    area_rect = a_inline.rect_global(base_area)
                    current_area = area_rect
                    move_mouse_ready = True
                    self.signals.current.emit(
                        f"Область: {area_rect.left()},{area_rect.top()} → {area_rect.right()},{area_rect.bottom()}"
                    )
                    if a_inline.click and area_rect.isValid() and area_rect.width() > 2 and area_rect.height() > 2:
                        cx = (area_rect.left() + area_rect.right()) // 2
                        cy = (area_rect.top() + area_rect.bottom()) // 2
                        smooth_move_to(cx, cy, 0.05)
                        perform_trigger_times(
                            getattr(a_inline, "trigger", DEFAULT_TRIGGER),
                            a_inline.multiplier,
                            a_inline.delay,
                        )
                    return True

                if isinstance(a_inline, WaitEventAction):
                    if not ocr_available:
                        _emit_ocr_skip_notice()
                        self.signals.current.emit("Ожидание события пропущено: OCR/Tesseract не установлен.")
                        done += 1
                        self.signals.progress.emit(done, total)
                        return True

                    watch_rect = a_inline.rect_global(base_area)
                    if not watch_rect or not watch_rect.isValid() or watch_rect.width() < 2 or watch_rect.height() < 2:
                        msg = "Ожидание события: область не задана или пуста."
                        self.signals.action_error.emit(owner_idx, msg)
                        self.signals.current.emit(msg)
                        done += 1
                        self.signals.progress.emit(done, total)
                        return True

                    desc = action_to_display(a_inline)[1]
                    poll = max(0.1, float(getattr(a_inline, "poll", 1.0)))
                    self.signals.current.emit(f"Ожидание события: {desc}")
                    while not self._stop.is_set():
                        if not _wait_if_paused():
                            return False
                        found = a_inline.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                        ok = bool(found and found.isValid() and found.width() > 2 and found.height() > 2)
                        if ok:
                            self.signals.current.emit(f"Ожидание события выполнено: {desc}")
                            if found and found.isValid():
                                current_area = found
                            elif watch_rect and watch_rect.isValid():
                                current_area = watch_rect
                            done += 1
                            self.signals.progress.emit(done, total)
                            return True
                        if not _sleep_checked(poll):
                            return False
                    return False

                if isinstance(a_inline, WaitAction):
                    t = a_inline.delay.sample()
                    self.signals.current.emit(f"Ожидание: {t:.3f}с")
                    if not _sleep_checked(t):
                        return False
                    done += 1
                    self.signals.progress.emit(done, total)
                    return True

                if isinstance(a_inline, WordAreaAction):
                    if not ocr_available:
                        _emit_ocr_skip_notice()
                        self.signals.current.emit(f"Область слова '{a_inline.word}' пропущена: OCR/Tesseract не установлен.")
                        return True

                    RETRY_DELAY = 0.12
                    ROUND_DELAY = 5.0
                    try:
                        max_tries = int(getattr(a_inline, "search_max_tries", 100))
                    except Exception:
                        max_tries = 100
                    if max_tries < 1:
                        max_tries = 1
                    search_infinite = bool(getattr(a_inline, "search_infinite", True))
                    on_fail = str(getattr(a_inline, "search_on_fail", "retry") or "retry")
                    if on_fail not in ("retry", "error", "action"):
                        on_fail = "retry"

                    if search_infinite:
                        attempt = 0
                        found = None
                        while True:
                            if self._stop.is_set():
                                return False
                            if not _wait_if_paused():
                                return False
                            attempt += 1
                            found = a_inline.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                            if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                break
                            self.signals.current.emit(
                                f"Слово '{a_inline.word}' не найдено (попытка {attempt}) - повтор..."
                            )
                            if RETRY_DELAY > 0 and (not _sleep_checked(RETRY_DELAY)):
                                return False
                        if found and found.isValid() and found.width() > 2 and found.height() > 2:
                            current_area = found
                            move_mouse_ready = True
                            self.signals.current.emit(
                                f"Область слова: '{a_inline.word}' №{a_inline.index} -> "
                                f"{found.left()},{found.top()} → {found.right()},{found.bottom()}"
                            )
                            if a_inline.click:
                                cx = (found.left() + found.right()) // 2
                                cy = (found.top() + found.bottom()) // 2
                                smooth_move_to(cx, cy, 0.05)
                                perform_trigger_times(
                                    getattr(a_inline, "trigger", DEFAULT_TRIGGER),
                                    a_inline.multiplier,
                                    a_inline.delay,
                                )
                        return True

                    while True:
                        if self._stop.is_set():
                            return False
                        if not _wait_if_paused():
                            return False
                        found = None
                        for attempt in range(1, max_tries + 1):
                            if self._stop.is_set():
                                return False
                            if not _wait_if_paused():
                                return False
                            found = a_inline.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                            if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                break
                            if attempt < max_tries:
                                self.signals.current.emit(
                                    f"Слово '{a_inline.word}' не найдено (попытка {attempt}/{max_tries}) - повтор..."
                                )
                                if RETRY_DELAY > 0 and (not _sleep_checked(RETRY_DELAY)):
                                    return False

                        if found and found.isValid() and found.width() > 2 and found.height() > 2:
                            current_area = found
                            move_mouse_ready = True
                            self.signals.current.emit(
                                f"Область слова: '{a_inline.word}' №{a_inline.index} -> "
                                f"{found.left()},{found.top()} → {found.right()},{found.bottom()}"
                            )
                            if a_inline.click:
                                cx = (found.left() + found.right()) // 2
                                cy = (found.top() + found.bottom()) // 2
                                smooth_move_to(cx, cy, 0.05)
                                perform_trigger_times(
                                    getattr(a_inline, "trigger", DEFAULT_TRIGGER),
                                    a_inline.multiplier,
                                    a_inline.delay,
                                )
                            return True

                        msg = f"Не смог найти слово '{a_inline.word}' в зоне после {max_tries} попыток."
                        self.signals.action_error.emit(owner_idx, msg)
                        self.signals.current.emit(msg)
                        if on_fail == "error":
                            self._mark_end("error", msg)
                            self.signals.status.emit(msg)
                            self._stop.set()
                            return False
                        if on_fail == "action":
                            if not _execute_inline_actions(getattr(a_inline, "on_fail_actions", []) or [], owner_idx):
                                return False
                            if restart_from_beginning:
                                return True
                            if not _apply_on_fail_post_mode(a_inline):
                                return False
                            return True

                        wait_msg = f"Повтор поиска через {int(ROUND_DELAY)}с..."
                        self.signals.current.emit(wait_msg)
                        if not _sleep_checked(ROUND_DELAY):
                            return False

                # KeyAction
                detail = action_to_display(a_inline)[1]
                self.signals.current.emit(detail)
                mode = _normalize_key_press_mode(getattr(a_inline, "press_mode", "normal"))
                if mode == "long":
                    if not _execute_key_long_press_items(_key_long_action_items(a_inline), owner_idx):
                        return False
                    return True

                reps = max(1, int(a_inline.multiplier))
                for _r_i in range(reps):
                    if self._stop.is_set():
                        return False
                    t = a_inline.delay.sample()
                    if (move_mouse_ready and current_area and current_area.isValid() and
                            current_area.width() > 2 and current_area.height() > 2):
                        tx = random.randint(current_area.left() + 1, current_area.right() - 1)
                        ty = random.randint(current_area.top() + 1, current_area.bottom() - 1)
                        smooth_move_to(tx, ty, t)
                    else:
                        if t > 0 and (not _sleep_checked(t)):
                            return False

                    if self._stop.is_set():
                        return False

                    if a_inline.kind == "mouse" and a_inline.mouse_button:
                        mods_only = [k for k in a_inline.keys if k in ("Shift", "Ctrl", "Alt", "Meta")]
                        mod_pks = []
                        for mname in mods_only:
                            pk = k_to_pynput(mname)
                            if pk is not None:
                                mod_pks.append(pk)
                        for pk in mod_pks:
                            if self._stop.is_set():
                                return False
                            kb.press(pk)
                            _sleep_checked(0.005)
                        _sleep_checked(0.020)
                        if not self._stop.is_set():
                            click(a_inline.mouse_button)
                        for pk in reversed(mod_pks):
                            try:
                                kb.release(pk)
                            except Exception:
                                pass
                    else:
                        press_combo(a_inline.keys)

                    done += 1
                    self.signals.progress.emit(done, total)
                return True

            def _execute_inline_actions(
                raw_actions: List[Dict[str, Any]],
                owner_idx: int,
                source_label: str = "ветки «Если не найдено»",
                extend_total: bool = True,
            ) -> bool:
                nonlocal total
                seq = _build_inline_action_seq(raw_actions)
                source = str(source_label or "").strip() or "дополнительных действий"

                if not seq:
                    self.signals.current.emit(f"Список действий для {source} пуст.")
                    return True

                if extend_total:
                    total += _estimate_extra_progress_items(seq)
                self.signals.current.emit(f"Выполняю действий из {source}: {len(seq)}")
                for sub_a in seq:
                    if self._stop.is_set():
                        return False
                    if restart_from_beginning:
                        return True
                    if not _wait_if_paused():
                        return False
                    if not _execute_inline_action(sub_a, owner_idx):
                        return False
                return not self._stop.is_set()

            while True:
                if self._stop.is_set():
                    break

                cycle_num += 1
                self.signals.current.emit(f"Старт цикла #{cycle_num}")
                # run actions
                for idx in range(start_index, len(actions)):
                    if restart_from_beginning:
                        break
                    a = actions[idx]
                    if self._stop.is_set():
                        break
                    # NEW: РѕР±РЅРѕРІР»СЏРµРј СЃРЅР°РїС€РѕС‚ РґР»СЏ СЃС‚РѕРї-СЃР»РѕРІР°
                    _sw_set_snapshot(base_area, locals().get("current_dpr", 1.0))


                    if isinstance(a, BaseAreaAction):
                        runtime_rect, current_dpr, runtime_hwnd = _resolve_runtime_base()

                        if self._bound_exe and (runtime_rect is None or not runtime_rect.isValid()):
                            msg = "База: процесс/окно не найдено — остановка чтобы не кликать по экрану."
                            self._mark_end("error", msg)
                            self.signals.status.emit(msg)
                            self._stop.set()
                            break

                        if runtime_rect and runtime_rect.isValid():
                            base_area = runtime_rect
                        else:
                            base_area = a.rect()
                            current_dpr = 1.0

                        current_area = base_area
                        _sw_set_snapshot(base_area, current_dpr)

                        self.signals.current.emit(
                            f"Базовая область: {base_area.left()},{base_area.top()} → {base_area.right()},{base_area.bottom()}"
                        )

                        if a.click and base_area.isValid() and base_area.width() > 2 and base_area.height() > 2:
                            cx = (base_area.left() + base_area.right()) // 2
                            cy = (base_area.top() + base_area.bottom()) // 2
                            smooth_move_to(cx, cy, 0.05)
                            perform_trigger_times(getattr(a, "trigger", DEFAULT_TRIGGER), 1)
                        continue

                    self.signals.action_row.emit(idx)  # <-- Р’РћРў РўРЈРў, РЅР° РѕРґРЅРѕРј СѓСЂРѕРІРЅРµ СЃ if

                    if isinstance(a, WordAreaAction):
                        if not ocr_available:
                            _emit_ocr_skip_notice()
                            self.signals.action_ok.emit(idx)
                            self.signals.current.emit(f"Область слова '{a.word}' пропущена: OCR/Tesseract не установлен.")
                            continue

                        RETRY_DELAY = 0.12
                        ROUND_DELAY = 5.0
                        try:
                            max_tries = int(getattr(a, "search_max_tries", 100))
                        except Exception:
                            max_tries = 100
                        if max_tries < 1:
                            max_tries = 1
                        search_infinite = bool(getattr(a, "search_infinite", True))
                        on_fail = str(getattr(a, "search_on_fail", "retry") or "retry")
                        if on_fail not in ("retry", "error", "action"):
                            on_fail = "retry"

                        if search_infinite:
                            attempt = 0
                            found = None
                            while True:
                                if self._stop.is_set():
                                    break
                                if not _wait_if_paused():
                                    break

                                attempt += 1
                                found = a.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                                if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                    break

                                self.signals.current.emit(
                                    f"Слово '{a.word}' не найдено (попытка {attempt}) - повтор..."
                                )
                                if RETRY_DELAY > 0:
                                    if not _sleep_checked(RETRY_DELAY):
                                        break

                            if self._stop.is_set():
                                break

                            if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                current_area = found
                                move_mouse_ready = True
                                self.signals.action_ok.emit(idx)
                                self.signals.current.emit(
                                    f"Область слова: '{a.word}' №{a.index} -> "
                                    f"{found.left()},{found.top()} → {found.right()},{found.bottom()}"
                                )

                                if a.click:
                                    cx = (found.left() + found.right()) // 2
                                    cy = (found.top() + found.bottom()) // 2
                                    smooth_move_to(cx, cy, 0.05)
                                    perform_trigger_times(getattr(a, "trigger", DEFAULT_TRIGGER), a.multiplier, a.delay)
                            continue

                        while True:
                            if self._stop.is_set():
                                break
                            if not _wait_if_paused():
                                break

                            found = None

                            for attempt in range(1, max_tries + 1):
                                if self._stop.is_set():
                                    break
                                if not _wait_if_paused():
                                    break

                                found = a.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                                if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                    break

                                if attempt < max_tries:
                                    self.signals.current.emit(
                                        f"Слово '{a.word}' не найдено (попытка {attempt}/{max_tries}) - повтор..."
                                    )
                                    if RETRY_DELAY > 0:
                                        if not _sleep_checked(RETRY_DELAY):
                                            break

                            if self._stop.is_set():
                                break

                            if found and found.isValid() and found.width() > 2 and found.height() > 2:
                                current_area = found
                                move_mouse_ready = True
                                self.signals.action_ok.emit(idx)
                                self.signals.current.emit(
                                    f"Область слова: '{a.word}' №{a.index} -> "
                                    f"{found.left()},{found.top()} → {found.right()},{found.bottom()}"
                                )

                                if a.click:
                                    cx = (found.left() + found.right()) // 2
                                    cy = (found.top() + found.bottom()) // 2
                                    smooth_move_to(cx, cy, 0.05)
                                    perform_trigger_times(getattr(a, "trigger", DEFAULT_TRIGGER), a.multiplier, a.delay)
                                break

                            msg = f"Не смог найти слово '{a.word}' в зоне после {max_tries} попыток."
                            self.signals.action_error.emit(idx, msg)
                            self.signals.current.emit(msg)

                            if on_fail == "error":
                                self._mark_end("error", msg)
                                self.signals.status.emit(msg)
                                self._stop.set()
                                break
                            if on_fail == "action":
                                if not _execute_inline_actions(getattr(a, "on_fail_actions", []) or [], idx):
                                    break
                                if restart_from_beginning:
                                    break
                                if not _apply_on_fail_post_mode(a):
                                    break
                                break

                            wait_msg = f"Повтор поиска через {int(ROUND_DELAY)}с..."
                            self.signals.current.emit(wait_msg)
                            if not _sleep_checked(ROUND_DELAY):
                                break

                        if self._stop.is_set():
                            break
                        if restart_from_beginning:
                            break
                        continue
                        # Р¶РґС‘Рј resume/stop
                        if not _wait_if_paused():
                            break

                        if self._stop.is_set():
                            break

                    if isinstance(a, AreaAction):
                        current_area = a.rect_global(base_area)
                        move_mouse_ready = True
                        self.signals.current.emit(
                            f"Область: {current_area.left()},{current_area.top()} → {current_area.right()},{current_area.bottom()}")

                        if a.click and current_area.isValid() and current_area.width() > 2 and current_area.height() > 2:
                            cx = (current_area.left() + current_area.right()) // 2
                            cy = (current_area.top() + current_area.bottom()) // 2
                            smooth_move_to(cx, cy, 0.05)
                            perform_trigger_times(getattr(a, "trigger", DEFAULT_TRIGGER), a.multiplier, a.delay)

                        continue

                    if isinstance(a, WaitEventAction):
                        if not ocr_available:
                            _emit_ocr_skip_notice()
                            self.signals.action_ok.emit(idx)
                            self.signals.current.emit("Ожидание события пропущено: OCR/Tesseract не установлен.")
                            done += 1
                            self.signals.progress.emit(done, total)
                            continue

                        watch_rect = a.rect_global(base_area)
                        if not watch_rect or not watch_rect.isValid() or watch_rect.width() < 2 or watch_rect.height() < 2:
                            msg = "Ожидание события: область не задана или пуста."
                            self.signals.action_error.emit(idx, msg)
                            self.signals.current.emit(msg)
                            done += 1
                            self.signals.progress.emit(done, total)
                            continue

                        desc = action_to_display(a)[1]
                        poll = max(0.1, float(getattr(a, "poll", 1.0)))
                        self.signals.current.emit(f"Ожидание события: {desc}")

                        while not self._stop.is_set():
                            if not _wait_if_paused():
                                break

                            found = a.resolve_target_rect_global(base_area, dpr_override=current_dpr)
                            ok = bool(found and found.isValid() and found.width() > 2 and found.height() > 2)

                            if ok:
                                self.signals.action_ok.emit(idx)
                                self.signals.current.emit(f"Ожидание события выполнено: {desc}")
                                if found and found.isValid():
                                    current_area = found
                                elif watch_rect and watch_rect.isValid():
                                    current_area = watch_rect
                                done += 1
                                self.signals.progress.emit(done, total)
                                break

                            if not _sleep_checked(poll):
                                break

                        continue

                    if isinstance(a, WaitAction):
                        t = a.delay.sample()
                        self.signals.current.emit(f"Ожидание: {t:.3f}с")
                        if not _sleep_checked(t):
                            break
                        done += 1
                        self.signals.progress.emit(done, total)

                        continue

                    # KeyAction
                    detail = action_to_display(a)[1]
                    self.signals.current.emit(detail)
                    mode = _normalize_key_press_mode(getattr(a, "press_mode", "normal"))
                    if mode == "long":
                        if not _execute_key_long_press_items(_key_long_action_items(a), idx):
                            break
                        continue

                    reps = max(1, int(a.multiplier))
                    for r_i in range(reps):
                        if self._stop.is_set():
                            break

                        t = a.delay.sample()

                        # move smoothly during the delay time
                        if (move_mouse_ready and current_area and current_area.isValid() and
                                current_area.width() > 2 and current_area.height() > 2):
                            tx = random.randint(current_area.left() + 1, current_area.right() - 1)
                            ty = random.randint(current_area.top() + 1, current_area.bottom() - 1)
                            smooth_move_to(tx, ty, t)
                        else:
                            if t > 0:
                                if not _sleep_checked(t):
                                    break

                        if self._stop.is_set():
                            break

                        # do the actual action
                        if a.kind == "mouse" and a.mouse_button:
                            mods_only = [k for k in a.keys if k in ("Shift", "Ctrl", "Alt", "Meta")]

                            # Р·Р°Р¶РёРјР°РµРј РјРѕРґС‹ РІСЂСѓС‡РЅСѓСЋ, РєР»РёРєР°РµРј, РѕС‚РїСѓСЃРєР°РµРј
                            mod_pks = []
                            for mname in mods_only:
                                pk = k_to_pynput(mname)
                                if pk is not None:
                                    mod_pks.append(pk)

                            for pk in mod_pks:
                                if self._stop.is_set():
                                    break
                                kb.press(pk)
                                _sleep_checked(0.005)

                            _sleep_checked(0.020)  # РґР°С‚СЊ РёРіСЂРµ СѓРІРёРґРµС‚СЊ РјРѕРґС‹
                            if not self._stop.is_set():
                                click(a.mouse_button)

                            for pk in reversed(mod_pks):
                                try:
                                    kb.release(pk)
                                except Exception:
                                    pass
                        else:
                            press_combo(a.keys)

                        done += 1
                        self.signals.progress.emit(done, total)

                if self._stop.is_set():
                    break

                if restart_from_beginning:
                    restart_from_beginning = False

                    if self.record.repeat.enabled and (not infinite):
                        cycles_left -= 1
                        self.signals.cycle.emit(cycle_label())
                        if cycles_left <= 0:
                            break
                    continue

                if not self.record.repeat.enabled:
                    break

                # decrement cycles if finite
                if not infinite:
                    cycles_left -= 1
                    self.signals.cycle.emit(cycle_label())
                    if cycles_left <= 0:
                        break

                # wait between cycles
                d = self.record.repeat.delay.sample()
                if d > 0:
                    self.signals.current.emit(f"Пауза перед повтором: {d:.3f}с")
                    if not _sleep_checked(d):
                        break

        except Exception as ex:
            msg = f"Ошибка в проигрывателе: {type(ex).__name__}: {ex}"
            self._mark_end("error", msg)
            self.signals.status.emit(msg)
        finally:
            # РµСЃР»Рё РїРѕС‚РѕРє РѕСЃС‚Р°РЅРѕРІРёР»Рё, РЅРѕ РїСЂРёС‡РёРЅСѓ РЅРµ РїСЂРѕСЃС‚Р°РІРёР»Рё
            if self._stop.is_set() and getattr(self, "end_state", "done") == "done":
                self._mark_end("stopped", "Проигрывание остановлено.")
            self.signals.finished.emit()
