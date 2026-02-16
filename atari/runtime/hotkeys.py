# Module: global hotkeys and interval meter.
# Main: GlobalHotkeyListener, IntervalMeter.
# Example: from atari.runtime.hotkeys import GlobalHotkeyListener

import threading
import time
from typing import Optional

from PySide6.QtCore import QObject, Signal

class GlobalHotkeyListener(QObject):
    esc_pressed = Signal()
    f1_pressed = Signal()
    f2_pressed = Signal()
    f3_pressed = Signal()
    f4_pressed = Signal()
    f5_pressed = Signal()

    def __init__(self):
        super().__init__()
        self._listener = None

    def start(self):
        if self._listener is not None:
            return
        try:
            from pynput import keyboard
        except Exception:
            return

        def on_press(key):
            try:
                if key == keyboard.Key.esc:
                    self.esc_pressed.emit()
                elif key == keyboard.Key.f1:
                    self.f1_pressed.emit()
                elif key == keyboard.Key.f2:
                    self.f2_pressed.emit()
                elif key == keyboard.Key.f3:
                    self.f3_pressed.emit()
                elif key == keyboard.Key.f4:
                    self.f4_pressed.emit()
                elif key == keyboard.Key.f5:
                    self.f5_pressed.emit()
            except Exception:
                pass

        self._listener = keyboard.Listener(on_press=on_press)
        self._listener.daemon = True
        self._listener.start()

    def stop(self):
        if self._listener is None:
            return
        try:
            self._listener.stop()
        except Exception:
            pass
        self._listener = None



class IntervalMeter(QObject):
    interval = Signal(float, str, int)  # dt_sec, name, index
    status = Signal(str)
    stopped = Signal()

    def __init__(self):
        super().__init__()
        self._kb_listener = None
        self._ms_listener = None

        self._lock = threading.Lock()
        self._enabled = False
        self._armed = False  # ждём выбор целевой кнопки

        self._target_kind: Optional[str] = None  # "key" | "mouse"
        self._target_sig: Optional[tuple] = None

        self._last_ts: Optional[float] = None
        self.intervals: List[float] = []
        self._count = 0

        self._keys_down = set()
        self._mouse_down = set()

    # ---- normalize / pretty ----
    def _key_sig(self, kb, key) -> tuple:
        # чтобы сравнение работало стабильно (Key / KeyCode)
        try:
            if key in (kb.Key.esc, kb.Key.f4, kb.Key.f8):
                return ("key", key)
        except Exception:
            pass

        # Key (special)
        if hasattr(kb, "Key") and isinstance(key, kb.Key):
            return ("key", key)

        # KeyCode
        ch = getattr(key, "char", None)
        if ch:
            return ("char", str(ch).lower())
        vk = getattr(key, "vk", None)
        if vk is not None:
            return ("vk", int(vk))
        return ("str", str(key))

    def _mouse_sig(self, button) -> tuple:
        return ("mouse", button)

    def _key_to_name(self, key) -> str:
        ch = getattr(key, "char", None)
        if ch:
            return ch.upper() if ch.isalpha() else ch
        s = str(key)
        return s.replace("Key.", "")

    def _mouse_to_name(self, button) -> str:
        try:
            name = getattr(button, "name", "")
        except Exception:
            name = str(button)
        mapping = {"left": "ЛКМ", "middle": "СКМ", "right": "ПКМ"}
        return mapping.get(name, str(button))

    def _target_pretty(self) -> str:
        if self._target_kind == "mouse":
            # сигнатура ("mouse", Button.left)
            return self._mouse_to_name(self._target_sig[1]) if self._target_sig else "мышь"
        if self._target_kind == "key":
            # показать хотя бы тип
            return "клавиша"
        return "—"

    # ---- public API ----
    def start(self):
        # стартуем слушатели один раз
        if self._kb_listener is None or self._ms_listener is None:
            try:
                from pynput import keyboard, mouse
            except Exception as ex:
                self.status.emit(f"Замер: ошибка pynput: {ex}")
                return

            def on_key_press(key):
                from pynput import keyboard as kb

                # F8 — стоп
                if key == kb.Key.f8:
                    self.stop()
                    return

                # не ловим хоткеи приложения, чтобы не мешали
                if key in (kb.Key.esc, kb.Key.f1, kb.Key.f2, kb.Key.f3, kb.Key.f4, kb.Key.f5):
                    return

                sig = self._key_sig(kb, key)

                # фильтр автоповтора: пока держишь клавишу — не спамим
                with self._lock:
                    if sig in self._keys_down:
                        return
                    self._keys_down.add(sig)

                self._handle_event(kind="key", sig=sig, name=self._key_to_name(key))

            def on_key_release(key):
                from pynput import keyboard as kb
                sig = self._key_sig(kb, key)
                with self._lock:
                    self._keys_down.discard(sig)

            def on_click(x, y, button, pressed):
                # считаем только "нажатие", отпускание нужно для анти-дребезга
                if pressed:
                    sig = self._mouse_sig(button)
                    with self._lock:
                        if sig in self._mouse_down:
                            return
                        self._mouse_down.add(sig)
                    self._handle_event(kind="mouse", sig=sig, name=self._mouse_to_name(button))
                else:
                    sig = self._mouse_sig(button)
                    with self._lock:
                        self._mouse_down.discard(sig)

            self._kb_listener = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
            self._kb_listener.daemon = True
            self._kb_listener.start()

            self._ms_listener = mouse.Listener(on_click=on_click)
            self._ms_listener.daemon = True
            self._ms_listener.start()

        with self._lock:
            self._enabled = True
            self._armed = True
            self._target_kind = None
            self._target_sig = None

            self._last_ts = None
            self.intervals.clear()
            self._count = 0
            self._keys_down.clear()
            self._mouse_down.clear()

        self.status.emit(
            "Замер: нажми ОДНУ кнопку для выбора (например ЛКМ). Потом считаю интервалы только по ней. F8 — стоп.")

    def stop(self):
        with self._lock:
            if not self._enabled:
                return
            self._enabled = False
            self._armed = False
        self.status.emit("Замер: остановлен.")
        self.stopped.emit()

    def is_running(self) -> bool:
        with self._lock:
            return bool(self._enabled)

    # ---- core ----
    def _handle_event(self, kind: str, sig: tuple, name: str):
        ts = time.perf_counter()

        with self._lock:
            if not self._enabled:
                return

            # 1) Выбор целевой кнопки
            if self._armed:
                self._armed = False
                self._target_kind = kind
                self._target_sig = sig

                # считаем это первым нажатием (от него пойдёт первый интервал)
                self._last_ts = ts
                chosen = name
                # для клавы иногда nice показать имя
                if kind == "key":
                    chosen = name
                self.status.emit(f"Замер: выбрано «{chosen}». Теперь нажимай ЕЁ же — считаю интервалы. F8 — стоп.")
                return

            # 2) Фильтр: считаем только выбранную кнопку
            if kind != self._target_kind or sig != self._target_sig:
                return

            # 3) Интервал
            if self._last_ts is None:
                self._last_ts = ts
                return

            dt = ts - self._last_ts
            self._last_ts = ts

            self.intervals.append(dt)
            self._count += 1
            idx = self._count

        # имя для таблицы — выбранная кнопка (фиксированная)
        self.interval.emit(dt, name, idx)


