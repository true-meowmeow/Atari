# Module: overlay widgets and key helper formatting.
# Main: AreaSelectOverlay, KeyCaptureOverlay, spec_to_pretty.
# Example: from atari.ui.overlays import AreaSelectOverlay

from typing import Dict, Optional

from PySide6.QtCore import QEventLoop, QPoint, QRect, QSize, Qt, Signal, QTimer
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QWidget

from atari.core import config
from atari.core.geometry import virtual_geometry
from atari.core.models import KeyAction, MouseButtonName
from atari.core.win32 import _is_windows, _win_activate_hwnd
from atari.localization import i18n

class AreaSelectOverlay(QWidget):
    accepted = Signal(QRect)  # global rect
    canceled = Signal()

    def __init__(self, initial_global: Optional[QRect] = None):
        super().__init__(None)
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.Tool |
            Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setMouseTracking(True)

        self._vg = virtual_geometry()
        self.setGeometry(self._vg)

        self._rect_local: Optional[QRect] = None
        if initial_global and initial_global.isValid():
            self._rect_local = QRect(
                self.mapFromGlobal(initial_global.topLeft()),
                self.mapFromGlobal(initial_global.bottomRight()),
            ).normalized()

        self._dragging = False
        self._mode: Optional[str] = None  # "create"|"move"|"resize"
        self._resize_handle: Optional[str] = None
        self._anchor = QPoint()
        self._last_pos = QPoint()

        self._handle_size = 10
        self._edge_margin = 8

        self._help = (
            "Выбор области\n"
            "— Тяни ЛКМ чтобы создать прямоугольник\n"
            "— Потяни за края/углы чтобы изменить размер, или перетащи внутри чтобы двигать\n"
            "F1: применить   F2: сбросить   F3: выйти"
        )

        self.setFocusPolicy(Qt.StrongFocus)

    def show_and_block(self) -> Optional[QRect]:
        self.showFullScreen()
        self.raise_()
        self.activateWindow()
        self.setFocus(Qt.ActiveWindowFocusReason)
        if _is_windows():
            try:
                _win_activate_hwnd(int(self.winId()))
            except Exception:
                pass
        try:
            self.grabKeyboard()
            self.grabMouse()
        except Exception:
            pass

        result: Dict[str, Optional[QRect]] = {"rect": None}
        loop = QEventLoop()

        def on_accept(r: QRect):
            result["rect"] = r
            loop.quit()

        def on_cancel():
            loop.quit()

        self.accepted.connect(on_accept)
        self.canceled.connect(on_cancel)

        loop.exec()
        try:
            self.releaseKeyboard()
            self.releaseMouse()
        except Exception:
            pass
        self.close()
        self.deleteLater()
        return result["rect"]

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_F1:
            if self._rect_local and self._rect_local.isValid() and self._rect_local.width() >= 5 and self._rect_local.height() >= 5:
                tl = self.mapToGlobal(self._rect_local.topLeft())
                br = self.mapToGlobal(self._rect_local.bottomRight())
                self.accepted.emit(QRect(tl, br).normalized())
            else:
                # no rect
                pass
            return
        if e.key() == Qt.Key_F2:
            self._rect_local = None
            self.update()
            return
        if e.key() == Qt.Key_F3 or e.key() == Qt.Key_Escape:
            self.canceled.emit()
            return
        super().keyPressEvent(e)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        pos = e.position().toPoint()
        self._last_pos = pos

        if self._rect_local and self._rect_local.isValid():
            handle = self._hit_test_handle(pos, self._rect_local)
            if handle:
                self._mode = "resize"
                self._resize_handle = handle
                self._dragging = True
                self._anchor = pos
                return
            if self._rect_local.contains(pos):
                self._mode = "move"
                self._dragging = True
                self._anchor = pos
                return

        # create new
        self._mode = "create"
        self._dragging = True
        self._anchor = pos
        self._rect_local = QRect(pos, pos)
        self.update()

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        self._last_pos = pos

        if not self._dragging or not self._mode or not self._rect_local:
            self.update()
            return

        if self._mode == "create":
            self._rect_local = QRect(self._anchor, pos).normalized()
            self.update()
            return

        if self._mode == "move":
            delta = pos - self._anchor
            self._anchor = pos
            r = QRect(self._rect_local)
            r.translate(delta)
            # keep inside overlay
            r = self._clamp_rect(r)
            self._rect_local = r
            self.update()
            return

        if self._mode == "resize":
            r = QRect(self._rect_local)
            self._apply_resize(r, pos, self._resize_handle or "")
            r = r.normalized()
            r = self._clamp_rect(r)
            self._rect_local = r
            self.update()
            return

    def mouseReleaseEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        self._dragging = False
        self._mode = None
        self._resize_handle = None
        self.update()

    def _clamp_rect(self, r: QRect) -> QRect:
        bounds = QRect(0, 0, self.width(), self.height())
        # clamp by translation if outside
        dx = 0
        dy = 0
        if r.left() < bounds.left():
            dx = bounds.left() - r.left()
        if r.top() < bounds.top():
            dy = bounds.top() - r.top()
        if r.right() > bounds.right():
            dx = bounds.right() - r.right()
        if r.bottom() > bounds.bottom():
            dy = bounds.bottom() - r.bottom()
        r.translate(dx, dy)
        # final clamp sizes
        r.setLeft(max(bounds.left(), r.left()))
        r.setTop(max(bounds.top(), r.top()))
        r.setRight(min(bounds.right(), r.right()))
        r.setBottom(min(bounds.bottom(), r.bottom()))
        return r

    def _hit_test_handle(self, p: QPoint, r: QRect) -> Optional[str]:
        hs = self._handle_size
        # corners
        handles = {
            "tl": QRect(r.topLeft() - QPoint(hs // 2, hs // 2), QSize(hs, hs)),
            "tr": QRect(QPoint(r.right(), r.top()) - QPoint(hs // 2, hs // 2), QSize(hs, hs)),
            "bl": QRect(QPoint(r.left(), r.bottom()) - QPoint(hs // 2, hs // 2), QSize(hs, hs)),
            "br": QRect(r.bottomRight() - QPoint(hs // 2, hs // 2), QSize(hs, hs)),
        }
        for k, hr in handles.items():
            if hr.contains(p):
                return k
        # edges
        em = self._edge_margin
        if abs(p.x() - r.left()) <= em and r.top() <= p.y() <= r.bottom():
            return "l"
        if abs(p.x() - r.right()) <= em and r.top() <= p.y() <= r.bottom():
            return "r"
        if abs(p.y() - r.top()) <= em and r.left() <= p.x() <= r.right():
            return "t"
        if abs(p.y() - r.bottom()) <= em and r.left() <= p.x() <= r.right():
            return "b"
        return None

    def _apply_resize(self, r: QRect, pos: QPoint, handle: str):
        minw, minh = 10, 10
        if handle in ("l", "tl", "bl"):
            r.setLeft(pos.x())
        if handle in ("r", "tr", "br"):
            r.setRight(pos.x())
        if handle in ("t", "tl", "tr"):
            r.setTop(pos.y())
        if handle in ("b", "bl", "br"):
            r.setBottom(pos.y())
        rn = r.normalized()
        if rn.width() < minw:
            # expand back
            if handle in ("l", "tl", "bl"):
                r.setLeft(r.right() - minw)
            else:
                r.setRight(r.left() + minw)
        if rn.height() < minh:
            if handle in ("t", "tl", "tr"):
                r.setTop(r.bottom() - minh)
            else:
                r.setBottom(r.top() + minh)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        # dark overlay
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(0, 0, 0, 160))
        p.drawRect(self.rect())

        # clear selection "hole"
        if self._rect_local and self._rect_local.isValid():
            p.setCompositionMode(QPainter.CompositionMode_Clear)
            p.drawRect(self._rect_local)
            p.setCompositionMode(QPainter.CompositionMode_SourceOver)

            # border + handles
            p.setPen(QPen(QColor(255, 255, 255, 230), 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(self._rect_local)

            # handles
            hs = self._handle_size
            p.setBrush(QBrush(QColor(255, 255, 255, 230)))
            for pt in [self._rect_local.topLeft(), QPoint(self._rect_local.right(), self._rect_local.top()),
                       QPoint(self._rect_local.left(), self._rect_local.bottom()), self._rect_local.bottomRight()]:
                hr = QRect(pt - QPoint(hs // 2, hs // 2), QSize(hs, hs))
                p.drawRect(hr)

        # help text
        p.setPen(QColor(255, 255, 255, 230))
        p.setFont(QFont("Segoe UI", 12))
        p.drawText(QRect(20, 20, 560, 180), Qt.TextWordWrap, i18n.tr(self._help))


class KeyCaptureOverlay(QWidget):
    accepted = Signal(dict)  # dict for KeyAction fields: kind, keys, mouse_button
    canceled = Signal()

    # Вставь внутрь класса KeyCaptureOverlay
    def _start_pynput_capture(self):
        if self._kb_listener is not None or self._ms_listener is not None:
            return
        keyboard = None
        mouse = None
        try:
            from pynput import keyboard as _kb
            keyboard = _kb
        except Exception:
            keyboard = None
        try:
            from pynput import mouse as _ms
            mouse = _ms
        except Exception:
            mouse = None
        if keyboard is None and mouse is None:
            return

        def on_press(key):
            name = self._pynput_key_to_name(key)
            if name:
                QTimer.singleShot(0, lambda n=name: self._on_pynput_key(n, True))

        def on_release(key):
            name = self._pynput_key_to_name(key)
            if name:
                QTimer.singleShot(0, lambda n=name: self._on_pynput_key(n, False))

        def on_click(_x, _y, button, pressed):
            if not pressed:
                return
            name = getattr(button, "name", "")
            if name in ("left", "middle", "right"):
                QTimer.singleShot(0, lambda n=name: self._on_pynput_mouse(n))

        if keyboard is not None:
            self._kb_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
            self._kb_listener.daemon = True
            self._kb_listener.start()

        if mouse is not None:
            self._ms_listener = mouse.Listener(on_click=on_click)
            self._ms_listener.daemon = True
            self._ms_listener.start()

    def _stop_pynput_capture(self):
        if self._kb_listener is not None:
            try:
                self._kb_listener.stop()
            except Exception:
                pass
            self._kb_listener = None
        if self._ms_listener is not None:
            try:
                self._ms_listener.stop()
            except Exception:
                pass
            self._ms_listener = None

    def _pynput_key_to_name(self, key) -> Optional[str]:
        try:
            from pynput import keyboard as kb
        except Exception:
            return None

        try:
            if isinstance(key, kb.Key):
                specials = {
                    kb.Key.esc: "Escape",
                    kb.Key.enter: "Enter",
                    kb.Key.tab: "Tab",
                    kb.Key.backspace: "Backspace",
                    kb.Key.delete: "Delete",
                    kb.Key.insert: "Insert",
                    kb.Key.space: "Space",
                    kb.Key.home: "Home",
                    kb.Key.end: "End",
                    kb.Key.page_up: "PageUp",
                    kb.Key.page_down: "PageDown",
                    kb.Key.up: "Up",
                    kb.Key.down: "Down",
                    kb.Key.left: "Left",
                    kb.Key.right: "Right",
                    kb.Key.caps_lock: "CapsLock",
                    kb.Key.print_screen: "PrintScreen",
                    kb.Key.pause: "Pause",
                    kb.Key.ctrl: "Ctrl",
                    kb.Key.ctrl_l: "Ctrl",
                    kb.Key.ctrl_r: "Ctrl",
                    kb.Key.shift: "Shift",
                    kb.Key.shift_l: "Shift",
                    kb.Key.shift_r: "Shift",
                    kb.Key.alt: "Alt",
                    kb.Key.alt_l: "Alt",
                    kb.Key.alt_r: "Alt",
                    kb.Key.alt_gr: "Alt",
                    kb.Key.cmd: "Meta",
                    kb.Key.cmd_l: "Meta",
                    kb.Key.cmd_r: "Meta",
                }
                if key in specials:
                    return specials[key]
                name = getattr(key, "name", "")
                if name.startswith("f") and name[1:].isdigit():
                    return f"F{int(name[1:])}"
                return None

            ch = getattr(key, "char", None)
            if ch and len(ch) == 1:
                return ch.upper() if ch.isalpha() else ch
        except Exception:
            return None

        return None

    def _on_pynput_key(self, name: str, pressed: bool):
        if getattr(self, "_closing", False):
            return

        if pressed:
            if name in self._keys_down:
                return
            self._keys_down.add(name)
        else:
            self._keys_down.discard(name)

        if pressed:
            if name == "F1":
                self.accepted.emit(self._build_spec())
                return
            if name == "F2":
                self._pressed_mods.clear()
                self._combo_mods = None
                self._last_key = None
                self._last_mouse = None
                self.update()
                return
            if name in ("F3", "Escape"):
                self.canceled.emit()
                return

        if name in ("Shift", "Ctrl", "Alt", "Meta"):
            if pressed:
                self._pressed_mods.add(name)
            else:
                self._pressed_mods.discard(name)
            self.update()
            return

        if pressed:
            self._last_key = name
            self._last_mouse = None
            self._combo_mods = set(self._pressed_mods)
            self.update()

    def _on_pynput_mouse(self, name: str):
        if getattr(self, "_closing", False):
            return
        self._combo_mods = set(self._pressed_mods)
        self._last_mouse = name
        self._last_key = None
        self.update()

    def _mods_from_event(self, e):
        m = e.modifiers()
        mods = set()
        if m & Qt.ShiftModifier:
            mods.add("Shift")
        if m & Qt.ControlModifier:
            mods.add("Ctrl")
        if m & Qt.AltModifier:
            mods.add("Alt")
        if m & Qt.MetaModifier:
            mods.add("Meta")
        return mods

    def keyPressEvent(self, e):
        # служебные клавиши режима захвата
        if e.key() == Qt.Key_F1:
            self.accepted.emit(self._build_spec())
            return
        if e.key() == Qt.Key_F2:
            self._pressed_mods.clear()
            self._combo_mods = None  # <-- NEW
            self._last_key = None
            self._last_mouse = None
            self.update()
            return
        if e.key() in (Qt.Key_F3, Qt.Key_Escape):
            self.canceled.emit()
            return

        if e.isAutoRepeat():
            return

        # обновляем модификаторы (Shift/Ctrl/Alt/Meta) безопасно
        self._pressed_mods = self._mods_from_event(e)

        # если нажали сам модификатор — просто обновили состояние и всё
        if e.key() in (Qt.Key_Shift, Qt.Key_Control, Qt.Key_Alt, Qt.Key_Meta):
            self.update()
            return

        # иначе это "основная" клавиша комбинации
        name = qt_key_to_name(e.key(), e.text())
        if name:
            self._last_key = name
            self._last_mouse = None
            self._combo_mods = set(self._pressed_mods)  # <-- NEW: фиксируем Shift/Ctrl/Alt/Meta на момент нажатия

        self.update()

    def keyReleaseEvent(self, e):
        if e.isAutoRepeat():
            return

        # после отпускания пересчитываем модификаторы
        self._pressed_mods = self._mods_from_event(e)

        # иногда Qt ещё "видит" отпускаемый модификатор в modifiers(),
        # поэтому гарантированно убираем его вручную:
        if e.key() == Qt.Key_Shift:
            self._pressed_mods.discard("Shift")
        elif e.key() == Qt.Key_Control:
            self._pressed_mods.discard("Ctrl")
        elif e.key() == Qt.Key_Alt:
            self._pressed_mods.discard("Alt")
        elif e.key() == Qt.Key_Meta:
            self._pressed_mods.discard("Meta")

        self.update()

    def __init__(self, area_global: Optional[QRect], initial: Optional[KeyAction] = None):
        super().__init__(None)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setMouseTracking(True)

        self._vg = virtual_geometry()
        self.setGeometry(self._vg)
        self.setFocusPolicy(Qt.StrongFocus)

        self._area_local: Optional[QRect] = None
        if area_global and area_global.isValid():
            self._area_local = QRect(self.mapFromGlobal(area_global.topLeft()),
                                     self.mapFromGlobal(area_global.bottomRight())).normalized()

        self._pressed_mods: set[str] = set()
        self._last_key: Optional[str] = None
        self._last_mouse: Optional[MouseButtonName] = None
        self._combo_mods: Optional[set[str]] = None  # <-- NEW: модификаторы, зафиксированные в момент выбора
        self._kb_listener = None
        self._ms_listener = None
        self._keys_down: set[str] = set()
        self._closing = False

        if initial:
            if initial.kind == "mouse" and initial.mouse_button:
                self._last_mouse = initial.mouse_button
            else:
                # take last non-mod as base
                base = None
                mods = set()
                for k in initial.keys:
                    if k in ("Shift", "Ctrl", "Alt", "Meta"):
                        mods.add(k)
                    else:
                        base = k
                self._pressed_mods = mods
                self._last_key = base

        self._help = (
            "Захват клавиши/комбинации\n"
            "— Нажми (например Shift+E) или кликни мышью (ЛКМ/СКМ/ПКМ)\n"
            "F1: применить   F2: сбросить   F3: выйти"
        )

    def show_and_block(self) -> Optional[dict]:
        config.CAPTURE_OVERLAY_ACTIVE = True
        self.showFullScreen()
        self.raise_()
        self.activateWindow()
        self.setFocus(Qt.ActiveWindowFocusReason)
        if _is_windows():
            try:
                _win_activate_hwnd(int(self.winId()))
            except Exception:
                pass
        # Force input focus even if the window manager denies activation.
        try:
            self.grabKeyboard()
            self.grabMouse()
        except Exception:
            pass
        self._start_pynput_capture()

        result: Dict[str, Optional[dict]] = {"spec": None}
        loop = QEventLoop()

        def on_accept(spec: dict):
            if self._closing:
                return
            self._closing = True
            result["spec"] = spec
            self._stop_pynput_capture()
            loop.quit()

        def on_cancel():
            if self._closing:
                return
            self._closing = True
            self._stop_pynput_capture()
            loop.quit()

        self.accepted.connect(on_accept)
        self.canceled.connect(on_cancel)

        loop.exec()
        self._stop_pynput_capture()
        try:
            self.releaseKeyboard()
            self.releaseMouse()
        except Exception:
            pass
        config.CAPTURE_OVERLAY_ACTIVE = False
        self.close()
        self.deleteLater()
        return result["spec"]

    def _build_spec(self) -> dict:
        # если уже выбрана базовая клавиша/мышь — используем зафиксированные моды,
        # чтобы отпускание Shift до F1 не ломало комбинацию
        mods_src = None
        if (self._last_key or self._last_mouse) and isinstance(self._combo_mods, set):
            mods_src = self._combo_mods
        else:
            mods_src = self._pressed_mods

        mods = list(sorted(mods_src, key=lambda x: ("Ctrl", "Alt", "Shift", "Meta").index(x)
        if x in ("Ctrl", "Alt", "Shift", "Meta") else 99))

        if self._last_mouse:
            return {"kind": "mouse", "keys": mods, "mouse_button": self._last_mouse}
        if self._last_key:
            return {"kind": "keys", "keys": mods + [self._last_key], "mouse_button": None}

        # only modifiers?
        if mods:
            return {"kind": "keys", "keys": mods, "mouse_button": None}
        return {"kind": "keys", "keys": [], "mouse_button": None}

    def mousePressEvent(self, e):
        self._pressed_mods = self._mods_from_event(e)  # <-- ВАЖНО
        self._combo_mods = set(self._pressed_mods)  # <-- NEW
        if e.button() == Qt.LeftButton:
            self._last_mouse = "left"
        elif e.button() == Qt.MiddleButton:
            self._last_mouse = "middle"
        elif e.button() == Qt.RightButton:
            self._last_mouse = "right"
        else:
            return
        self._last_key = None
        self.update()

    def closeEvent(self, _):
        self._stop_pynput_capture()
        try:
            self.releaseKeyboard()
            self.releaseMouse()
        except Exception:
            pass
        config.CAPTURE_OVERLAY_ACTIVE = False

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        # dark overlay
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(0, 0, 0, 160))
        p.drawRect(self.rect())

        # clear area hole (like area selection mode)
        if self._area_local and self._area_local.isValid():
            p.setCompositionMode(QPainter.CompositionMode_Clear)
            p.drawRect(self._area_local)
            p.setCompositionMode(QPainter.CompositionMode_SourceOver)
            p.setPen(QPen(QColor(255, 255, 255, 230), 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(self._area_local)

        # text
        p.setPen(QColor(255, 255, 255, 230))
        p.setFont(QFont("Segoe UI", 12))
        p.drawText(QRect(20, 20, 640, 160), Qt.TextWordWrap, i18n.tr(self._help))

        # current selection
        spec = self._build_spec()
        pretty = spec_to_pretty(spec)
        p.setFont(QFont("Segoe UI", 16, QFont.DemiBold))
        p.drawText(QRect(20, 190, 900, 60), Qt.AlignLeft | Qt.AlignVCenter, i18n.tr(f"Текущее: {pretty}"))


def qt_key_to_name(qt_key: int, text: str) -> Optional[str]:
    # ✅ Надёжно: буквы A..Z и цифры 0..9 берём по коду клавиши
    if Qt.Key_A <= qt_key <= Qt.Key_Z:
        return chr(int(qt_key))  # 'A'..'Z'
    if Qt.Key_0 <= qt_key <= Qt.Key_9:
        return chr(int(qt_key))  # '0'..'9'

    # fallback на текст (для знаков и т.п.)
    if text and len(text) == 1:
        ch = text
        if ch.isalpha():
            return ch.upper()
        return ch

    mapping = {
        Qt.Key_Escape: "Escape",
        Qt.Key_Return: "Enter",
        Qt.Key_Enter: "Enter",
        Qt.Key_Tab: "Tab",
        Qt.Key_Backspace: "Backspace",
        Qt.Key_Delete: "Delete",
        Qt.Key_Insert: "Insert",
        Qt.Key_Space: "Space",
        Qt.Key_Home: "Home",
        Qt.Key_End: "End",
        Qt.Key_PageUp: "PageUp",
        Qt.Key_PageDown: "PageDown",
        Qt.Key_Up: "Up",
        Qt.Key_Down: "Down",
        Qt.Key_Left: "Left",
        Qt.Key_Right: "Right",
        Qt.Key_CapsLock: "CapsLock",
        Qt.Key_Print: "PrintScreen",
        Qt.Key_Pause: "Pause",
    }

    if Qt.Key_F1 <= qt_key <= Qt.Key_F35:
        return f"F{qt_key - Qt.Key_F1 + 1}"

    return mapping.get(qt_key, None)



def spec_to_pretty(spec: dict) -> str:
    mods = spec.get("keys", [])
    kind = spec.get("kind")
    if kind == "mouse":
        btn = {
            "left": i18n.tr("ЛКМ"),
            "middle": i18n.tr("СКМ"),
            "right": i18n.tr("ПКМ"),
        }.get(spec.get("mouse_button"), i18n.tr("Мышь"))
        if mods:
            return " + ".join(mods + [btn])
        return btn
    keys = spec.get("keys", [])
    return " + ".join(keys) if keys else i18n.tr("(пусто)")


