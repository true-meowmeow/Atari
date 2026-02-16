# Module: main Qt UI window and layout helpers.
# Main: MainWindow, FlowLayout.
# Example: from atari.ui.main_window import MainWindow

import copy
import json
import os
import random
import re
import sys
import threading
import time
import ctypes
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from PySide6.QtCore import (
    QAbstractAnimation, QEasingCurve, QEvent, QEventLoop, QPoint, QRect, QSize, Qt, QThread, QTimer, QUrl, Signal, QVariantAnimation
)
from PySide6.QtGui import (
    QAction, QActionGroup, QBrush, QColor, QDesktopServices, QFont, QGuiApplication, QKeySequence, QPainter, QPen, QShortcut
)
from PySide6.QtWidgets import (
    QAbstractItemView, QAbstractSpinBox, QApplication, QButtonGroup, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog, QFormLayout,
    QDialog,
    QFrame, QGroupBox, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLayout,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMenu, QMessageBox,
    QPushButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter, QTableWidget,
    QTableWidgetItem, QTableWidgetSelectionRange, QToolButton, QVBoxLayout, QWidget,
)

from atari.core import config
from atari.core.geometry import rect_to_rel, rel_to_rect, virtual_geometry
from atari.core.models import (
    Action, AreaAction, BaseAreaAction, Delay, DEFAULT_TRIGGER,
    KeyAction, Record, RepeatSettings, WaitAction, WaitEventAction, WordAreaAction,
    action_from_dict, action_to_display, get_installed_ocr_languages, normalize_ocr_lang_spec, normalize_trigger,
)
from atari.core.win32 import (
    _is_windows, _win_exe_from_pid, _win_hwnd_from_point, _win_pid_from_hwnd,
    resolve_bound_base_rect_dip, resolve_hwnd_by_exe,
)
from atari.localization import i18n
from atari.runtime.hotkeys import GlobalHotkeyListener, IntervalMeter
from atari.runtime.player import MacroPlayer
from atari.ui.overlays import AreaSelectOverlay, KeyCaptureOverlay, spec_to_pretty

i18n.install_qt_translation_hooks()

class FlowLayout(QLayout):
    def __init__(self, parent=None, margin=0, hspacing=8, vspacing=8):
        super().__init__(parent)
        self._items = []
        self._h = hspacing
        self._v = vspacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index):
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientations(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize(0, 0)
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        m = self.contentsMargins()
        size.setWidth(size.width() + m.left() + m.right())
        size.setHeight(size.height() + m.top() + m.bottom())
        return size

    def _do_layout(self, rect, test_only):
        x = rect.x()
        y = rect.y()
        line_h = 0
        right = rect.x() + rect.width()

        for item in self._items:
            hint = item.sizeHint()
            next_x = x + hint.width()

            if next_x > right and line_h > 0:
                x = rect.x()
                y += line_h + self._v
                next_x = x + hint.width()
                line_h = 0

            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))

            x = next_x + self._h
            line_h = max(line_h, hint.height())

        return (y + line_h) - rect.y()


class ChipCheckBox(QCheckBox):
    """Checkbox with full-rect click hit area for chip-like styling."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Chip-like toggles should not show native dotted focus rectangle.
        self.setFocusPolicy(Qt.NoFocus)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(30)

    def hitButton(self, pos):
        return self.rect().contains(pos)


class ClickableLabel(QLabel):
    clicked = Signal()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()
            e.accept()
            return
        super().mousePressEvent(e)


class PressModeSlider(QWidget):
    """Two-state segmented slider with animated thumb."""

    modeChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mode = "normal"
        self._pos = 0.0
        self._labels = {
            "normal": "Обычное нажатие",
            "long": "Продолжительное нажатие",
        }
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumHeight(38)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._anim = QVariantAnimation(self)
        self._anim.setDuration(220)
        self._anim.setEasingCurve(QEasingCurve.InOutCubic)
        self._anim.valueChanged.connect(self._on_anim_value_changed)

    def mode(self) -> str:
        return self._mode

    def set_mode(self, mode: str, *, animate: bool = True, emit_signal: bool = False):
        mode_norm = "long" if str(mode or "").strip().lower() == "long" else "normal"
        target = 1.0 if mode_norm == "long" else 0.0
        changed = (mode_norm != self._mode)
        self._mode = mode_norm

        # Keep running animation if the same mode is requested again.
        if not changed:
            if animate:
                return
            if self._anim.state() != QAbstractAnimation.Stopped:
                return
            self._pos = float(target)
            self.update()
            return

        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(float(self._pos))
            self._anim.setEndValue(float(target))
            self._anim.start()
        else:
            self._anim.stop()
            self._pos = float(target)
            self.update()

        if changed and emit_signal:
            self.modeChanged.emit(self._mode)

    def _on_anim_value_changed(self, value):
        try:
            self._pos = float(value)
        except Exception:
            self._pos = 0.0
        self.update()

    def sizeHint(self):
        return QSize(460, 40)

    def mousePressEvent(self, e):
        if not self.isEnabled():
            return super().mousePressEvent(e)
        new_mode = "normal" if e.position().x() < (self.width() / 2.0) else "long"
        self.set_mode(new_mode, animate=True, emit_signal=True)
        e.accept()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Left, Qt.Key_A):
            self.set_mode("normal", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Right, Qt.Key_D):
            self.set_mode("long", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.set_mode("long" if self._mode == "normal" else "normal", animate=True, emit_signal=True)
            e.accept()
            return
        super().keyPressEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        r = self.rect().adjusted(1, 1, -1, -1)
        if r.width() <= 8 or r.height() <= 8:
            return

        enabled = self.isEnabled()
        track_bg = QColor("#141821")
        track_border = QColor("#2a3447")
        thumb_bg = QColor("#284162")
        thumb_border = QColor("#4d78ad")
        text_on = QColor("#f4f8ff")
        text_off = QColor("#a9b3c7")
        if not enabled:
            track_bg = QColor("#11161f")
            track_border = QColor("#1f2633")
            thumb_bg = QColor("#1d2a3c")
            thumb_border = QColor("#31465f")
            text_on = QColor("#92a0b5")
            text_off = QColor("#6f7f95")

        p.setPen(QPen(track_border, 1))
        p.setBrush(track_bg)
        p.drawRoundedRect(r, 12, 12)

        pad = 3
        inner = r.adjusted(pad, pad, -pad, -pad)
        half_w = max(1, inner.width() // 2)
        thumb_w = half_w
        thumb_x = int(round(inner.left() + (inner.width() - thumb_w) * float(self._pos)))
        thumb = QRect(thumb_x, inner.top(), thumb_w, inner.height())

        p.setPen(QPen(thumb_border, 1))
        p.setBrush(thumb_bg)
        p.drawRoundedRect(thumb, 9, 9)

        left = QRect(inner.left(), inner.top(), half_w, inner.height())
        right = QRect(inner.left() + half_w, inner.top(), inner.width() - half_w, inner.height())

        p.setPen(text_on if self._mode == "normal" else text_off)
        p.drawText(left, Qt.AlignCenter, i18n.tr(self._labels["normal"]))
        p.setPen(text_on if self._mode == "long" else text_off)
        p.drawText(right, Qt.AlignCenter, i18n.tr(self._labels["long"]))


class WaitModeSlider(QWidget):
    """Two-state segmented slider for wait action mode."""

    modeChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mode = "time"
        self._pos = 0.0
        self._labels = {
            "time": "Время",
            "event": "Событие",
        }
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumHeight(38)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._anim = QVariantAnimation(self)
        self._anim.setDuration(220)
        self._anim.setEasingCurve(QEasingCurve.InOutCubic)
        self._anim.valueChanged.connect(self._on_anim_value_changed)

    def mode(self) -> str:
        return self._mode

    def set_mode(self, mode: str, *, animate: bool = True, emit_signal: bool = False):
        mode_norm = "event" if str(mode or "").strip().lower() == "event" else "time"
        target = 1.0 if mode_norm == "event" else 0.0
        changed = (mode_norm != self._mode)
        self._mode = mode_norm

        # If the same mode is requested while an animation is already running,
        # keep the current animation to avoid jump-then-animate artifacts.
        if not changed:
            if animate:
                return
            if self._anim.state() != QAbstractAnimation.Stopped:
                return
            self._pos = float(target)
            self.update()
            return

        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(float(self._pos))
            self._anim.setEndValue(float(target))
            self._anim.start()
        else:
            self._anim.stop()
            self._pos = float(target)
            self.update()

        if changed and emit_signal:
            self.modeChanged.emit(self._mode)

    def _on_anim_value_changed(self, value):
        try:
            self._pos = float(value)
        except Exception:
            self._pos = 0.0
        self.update()

    def sizeHint(self):
        return QSize(460, 40)

    def mousePressEvent(self, e):
        if not self.isEnabled():
            return super().mousePressEvent(e)
        new_mode = "time" if e.position().x() < (self.width() / 2.0) else "event"
        self.set_mode(new_mode, animate=True, emit_signal=True)
        e.accept()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Left, Qt.Key_A):
            self.set_mode("time", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Right, Qt.Key_D):
            self.set_mode("event", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.set_mode("event" if self._mode == "time" else "time", animate=True, emit_signal=True)
            e.accept()
            return
        super().keyPressEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        r = self.rect().adjusted(1, 1, -1, -1)
        if r.width() <= 8 or r.height() <= 8:
            return

        enabled = self.isEnabled()
        track_bg = QColor("#141821")
        track_border = QColor("#2a3447")
        thumb_bg = QColor("#284162")
        thumb_border = QColor("#4d78ad")
        text_on = QColor("#f4f8ff")
        text_off = QColor("#a9b3c7")
        if not enabled:
            track_bg = QColor("#11161f")
            track_border = QColor("#1f2633")
            thumb_bg = QColor("#1d2a3c")
            thumb_border = QColor("#31465f")
            text_on = QColor("#92a0b5")
            text_off = QColor("#6f7f95")

        p.setPen(QPen(track_border, 1))
        p.setBrush(track_bg)
        p.drawRoundedRect(r, 12, 12)

        pad = 3
        inner = r.adjusted(pad, pad, -pad, -pad)
        half_w = max(1, inner.width() // 2)
        thumb_w = half_w
        thumb_x = int(round(inner.left() + (inner.width() - thumb_w) * float(self._pos)))
        thumb = QRect(thumb_x, inner.top(), thumb_w, inner.height())

        p.setPen(QPen(thumb_border, 1))
        p.setBrush(thumb_bg)
        p.drawRoundedRect(thumb, 9, 9)

        left = QRect(inner.left(), inner.top(), half_w, inner.height())
        right = QRect(inner.left() + half_w, inner.top(), inner.width() - half_w, inner.height())

        p.setPen(text_on if self._mode == "time" else text_off)
        p.drawText(left, Qt.AlignCenter, i18n.tr(self._labels["time"]))
        p.setPen(text_on if self._mode == "event" else text_off)
        p.drawText(right, Qt.AlignCenter, i18n.tr(self._labels["event"]))


class AreaModeSlider(QWidget):
    """Two-state segmented slider for area action mode."""

    modeChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mode = "screen"
        self._pos = 0.0
        self._labels = {
            "screen": "Экран",
            "text": "Текст",
        }
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumHeight(38)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._anim = QVariantAnimation(self)
        self._anim.setDuration(220)
        self._anim.setEasingCurve(QEasingCurve.InOutCubic)
        self._anim.valueChanged.connect(self._on_anim_value_changed)

    def mode(self) -> str:
        return self._mode

    def set_mode(self, mode: str, *, animate: bool = True, emit_signal: bool = False):
        mode_norm = "text" if str(mode or "").strip().lower() == "text" else "screen"
        target = 1.0 if mode_norm == "text" else 0.0
        changed = (mode_norm != self._mode)
        self._mode = mode_norm

        if not changed:
            if animate:
                return
            if self._anim.state() != QAbstractAnimation.Stopped:
                return
            self._pos = float(target)
            self.update()
            return

        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(float(self._pos))
            self._anim.setEndValue(float(target))
            self._anim.start()
        else:
            self._anim.stop()
            self._pos = float(target)
            self.update()

        if changed and emit_signal:
            self.modeChanged.emit(self._mode)

    def _on_anim_value_changed(self, value):
        try:
            self._pos = float(value)
        except Exception:
            self._pos = 0.0
        self.update()

    def sizeHint(self):
        return QSize(460, 40)

    def mousePressEvent(self, e):
        if not self.isEnabled():
            return super().mousePressEvent(e)
        new_mode = "screen" if e.position().x() < (self.width() / 2.0) else "text"
        self.set_mode(new_mode, animate=True, emit_signal=True)
        e.accept()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Left, Qt.Key_A):
            self.set_mode("screen", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Right, Qt.Key_D):
            self.set_mode("text", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.set_mode("text" if self._mode == "screen" else "screen", animate=True, emit_signal=True)
            e.accept()
            return
        super().keyPressEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        r = self.rect().adjusted(1, 1, -1, -1)
        if r.width() <= 8 or r.height() <= 8:
            return

        enabled = self.isEnabled()
        track_bg = QColor("#141821")
        track_border = QColor("#2a3447")
        thumb_bg = QColor("#284162")
        thumb_border = QColor("#4d78ad")
        text_on = QColor("#f4f8ff")
        text_off = QColor("#a9b3c7")
        if not enabled:
            track_bg = QColor("#11161f")
            track_border = QColor("#1f2633")
            thumb_bg = QColor("#1d2a3c")
            thumb_border = QColor("#31465f")
            text_on = QColor("#92a0b5")
            text_off = QColor("#6f7f95")

        p.setPen(QPen(track_border, 1))
        p.setBrush(track_bg)
        p.drawRoundedRect(r, 12, 12)

        pad = 3
        inner = r.adjusted(pad, pad, -pad, -pad)
        half_w = max(1, inner.width() // 2)
        thumb_w = half_w
        thumb_x = int(round(inner.left() + (inner.width() - thumb_w) * float(self._pos)))
        thumb = QRect(thumb_x, inner.top(), thumb_w, inner.height())

        p.setPen(QPen(thumb_border, 1))
        p.setBrush(thumb_bg)
        p.drawRoundedRect(thumb, 9, 9)

        left = QRect(inner.left(), inner.top(), half_w, inner.height())
        right = QRect(inner.left() + half_w, inner.top(), inner.width() - half_w, inner.height())

        p.setPen(text_on if self._mode == "screen" else text_off)
        p.drawText(left, Qt.AlignCenter, i18n.tr(self._labels["screen"]))
        p.setPen(text_on if self._mode == "text" else text_off)
        p.drawText(right, Qt.AlignCenter, i18n.tr(self._labels["text"]))


class LongActivationSlider(QWidget):
    """Two-state segmented slider for long-press item activation mode."""

    modeChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mode = "after_prev"
        self._pos = 0.0
        self._labels = {
            "after_prev": "После предыдущего",
            "from_start": "От старта действия",
        }
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumHeight(30)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._anim = QVariantAnimation(self)
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QEasingCurve.InOutCubic)
        self._anim.valueChanged.connect(self._on_anim_value_changed)

    def mode(self) -> str:
        return self._mode

    def set_mode(self, mode: str, *, animate: bool = True, emit_signal: bool = False):
        mode_norm = "from_start" if str(mode or "").strip().lower() == "from_start" else "after_prev"
        target = 1.0 if mode_norm == "from_start" else 0.0
        changed = (mode_norm != self._mode)
        self._mode = mode_norm

        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(float(self._pos))
            self._anim.setEndValue(float(target))
            self._anim.start()
        else:
            self._pos = float(target)
            self.update()

        if changed and emit_signal:
            self.modeChanged.emit(self._mode)

    def _on_anim_value_changed(self, value):
        try:
            self._pos = float(value)
        except Exception:
            self._pos = 0.0
        self.update()

    def sizeHint(self):
        return QSize(360, 32)

    def mousePressEvent(self, e):
        if not self.isEnabled():
            return super().mousePressEvent(e)
        new_mode = "after_prev" if e.position().x() < (self.width() / 2.0) else "from_start"
        self.set_mode(new_mode, animate=True, emit_signal=True)
        e.accept()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Left, Qt.Key_A):
            self.set_mode("after_prev", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Right, Qt.Key_D):
            self.set_mode("from_start", animate=True, emit_signal=True)
            e.accept()
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Space):
            self.set_mode("from_start" if self._mode == "after_prev" else "after_prev", animate=True, emit_signal=True)
            e.accept()
            return
        super().keyPressEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)

        r = self.rect().adjusted(1, 1, -1, -1)
        if r.width() <= 8 or r.height() <= 8:
            return

        enabled = self.isEnabled()
        track_bg = QColor("#141821")
        track_border = QColor("#2a3447")
        thumb_bg = QColor("#284162")
        thumb_border = QColor("#4d78ad")
        text_on = QColor("#f4f8ff")
        text_off = QColor("#a9b3c7")
        if not enabled:
            track_bg = QColor("#11161f")
            track_border = QColor("#1f2633")
            thumb_bg = QColor("#1d2a3c")
            thumb_border = QColor("#31465f")
            text_on = QColor("#92a0b5")
            text_off = QColor("#6f7f95")

        p.setPen(QPen(track_border, 1))
        p.setBrush(track_bg)
        p.drawRoundedRect(r, 10, 10)

        pad = 3
        inner = r.adjusted(pad, pad, -pad, -pad)
        half_w = max(1, inner.width() // 2)
        thumb_w = half_w
        thumb_x = int(round(inner.left() + (inner.width() - thumb_w) * float(self._pos)))
        thumb = QRect(thumb_x, inner.top(), thumb_w, inner.height())

        p.setPen(QPen(thumb_border, 1))
        p.setBrush(thumb_bg)
        p.drawRoundedRect(thumb, 8, 8)

        left = QRect(inner.left(), inner.top(), half_w, inner.height())
        right = QRect(inner.left() + half_w, inner.top(), inner.width() - half_w, inner.height())

        p.setPen(text_on if self._mode == "after_prev" else text_off)
        p.drawText(left, Qt.AlignCenter, i18n.tr(self._labels["after_prev"]))
        p.setPen(text_on if self._mode == "from_start" else text_off)
        p.drawText(right, Qt.AlignCenter, i18n.tr(self._labels["from_start"]))



# ---- Main UI ----
class MainWindow(QMainWindow):
    class ClickableGroupBox(QGroupBox):
        """QGroupBox, который сворачивается/разворачивается кликом по заголовку."""

        def mousePressEvent(self, e):
            try:
                y = e.position().y()
            except Exception:
                y = e.pos().y()

            # клики по верхней зоне (примерно заголовок)
            title_h = self.fontMetrics().height() + 14
            if self.isCheckable() and y <= title_h:
                self.setChecked(not self.isChecked())
                e.accept()
                return
            super().mousePressEvent(e)

    class DarkTitleDialog(QDialog):
        def __init__(self, owner, parent=None, on_close=None):
            super().__init__(parent)
            self._owner = owner
            self._on_close = on_close

        def showEvent(self, e):
            super().showEvent(e)
            if self._owner is not None:
                self._owner._schedule_dark_titlebar(self)

        def closeEvent(self, e):
            self.hide()
            if callable(self._on_close):
                try:
                    self._on_close(0)
                except Exception:
                    pass
            e.ignore()

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, "_startup_paint_synced", False):
            self._startup_paint_synced = True
            self.ensurePolished()
            style = self.style()
            if style is not None:
                style.unpolish(self)
                style.polish(self)
            self.update()
            self.repaint()
        if getattr(self, "_dark_titlebar_done", False):
            return
        self._dark_titlebar_done = True
        self._apply_dark_titlebar_windows()

    def _show_dark_info_dialog(
        self,
        title: str,
        text: str,
        accept_text: str = "OK",
        reject_text: str = "Отмена",
        *,
        danger_accept: bool = False,
    ) -> bool:
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.setModal(True)
        dlg.setMinimumWidth(460)
        self._prepare_dark_dialog(dlg)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        lbl = QLabel(text)
        lbl.setWordWrap(True)
        lay.addWidget(lbl)

        row = QHBoxLayout()
        row.addStretch(1)

        red_btn_style = (
            "QPushButton { background: #612121; border: 1px solid #7e2d2d; color: #f4dede; }"
            "QPushButton:hover { background: #742828; }"
            "QPushButton:pressed { background: #4f1a1a; }"
            "QPushButton:focus { outline: none; }"
        )
        green_btn_style = (
            "QPushButton { background: #1f5a35; border: 1px solid #2d7747; color: #dff4e7; }"
            "QPushButton:hover { background: #266d41; }"
            "QPushButton:pressed { background: #184a2b; }"
            "QPushButton:focus { outline: none; }"
        )
        cancel_style = green_btn_style if danger_accept else red_btn_style
        accept_style = red_btn_style if danger_accept else green_btn_style

        btn_cancel = QPushButton(reject_text)
        btn_cancel.setStyleSheet(cancel_style)
        btn_cancel.clicked.connect(dlg.reject)

        btn_ok = QPushButton(accept_text)
        btn_ok.setStyleSheet(accept_style)
        btn_ok.clicked.connect(dlg.accept)
        row.addWidget(btn_cancel)
        row.addWidget(btn_ok)
        lay.addLayout(row)

        i18n.retranslate_widget_tree(dlg)
        return dlg.exec() == QDialog.Accepted

    def _prepare_dark_dialog(self, dlg: QDialog):
        if dlg is None:
            return
        dlg.setStyleSheet(self.styleSheet())
        dlg.setAttribute(Qt.WA_NativeWindow, True)
        dlg.winId()
        self._apply_dark_titlebar_widget(dlg)
        self._schedule_dark_titlebar(dlg)

    def _input_text_dark(self, title: str, label: str, text: str = "") -> Tuple[str, bool]:
        dlg = QInputDialog(self)
        dlg.setWindowTitle(title)
        dlg.setLabelText(label)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setTextValue(str(text or ""))
        self._prepare_dark_dialog(dlg)
        ok = dlg.exec() == QDialog.Accepted
        return dlg.textValue(), ok

    def _input_int_dark(
        self,
        title: str,
        label: str,
        value: int,
        min_value: int,
        max_value: int,
        step: int = 1,
    ) -> Tuple[int, bool]:
        dlg = QInputDialog(self)
        dlg.setWindowTitle(title)
        dlg.setLabelText(label)
        dlg.setInputMode(QInputDialog.IntInput)
        dlg.setIntRange(int(min_value), int(max_value))
        dlg.setIntStep(max(1, int(step)))
        dlg.setIntValue(int(value))
        self._prepare_dark_dialog(dlg)
        ok = dlg.exec() == QDialog.Accepted
        return int(dlg.intValue()), ok

    def _input_double_dark(
        self,
        title: str,
        label: str,
        value: float,
        min_value: float,
        max_value: float,
        decimals: int = 1,
    ) -> Tuple[float, bool]:
        dlg = QInputDialog(self)
        dlg.setWindowTitle(title)
        dlg.setLabelText(label)
        dlg.setInputMode(QInputDialog.DoubleInput)
        dlg.setDoubleRange(float(min_value), float(max_value))
        dlg.setDoubleDecimals(max(0, int(decimals)))
        dlg.setDoubleValue(float(value))
        self._prepare_dark_dialog(dlg)
        ok = dlg.exec() == QDialog.Accepted
        return float(dlg.doubleValue()), ok

    def _input_item_dark(
        self,
        title: str,
        label: str,
        items: List[str],
        current: int = 0,
        editable: bool = False,
    ) -> Tuple[str, bool]:
        vals = [str(x) for x in (items or [])]
        dlg = QInputDialog(self)
        dlg.setWindowTitle(title)
        dlg.setLabelText(label)
        dlg.setInputMode(QInputDialog.TextInput)
        dlg.setComboBoxItems(vals)
        dlg.setComboBoxEditable(bool(editable))
        if vals:
            idx = min(max(int(current), 0), len(vals) - 1)
            dlg.setTextValue(vals[idx])
        self._prepare_dark_dialog(dlg)
        ok = dlg.exec() == QDialog.Accepted
        return dlg.textValue(), ok

    def _read_settings_payload(self) -> Dict[str, Any]:
        cached = getattr(self, "_settings_payload_cache", None)
        if isinstance(cached, dict):
            return dict(cached)
        try:
            if config.SETTINGS_PATH.exists():
                payload = json.loads(config.SETTINGS_PATH.read_text(encoding="utf-8"))
            else:
                payload = {}
        except Exception:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}
        self._settings_payload_cache = dict(payload)
        return dict(payload)

    def _retranslate_all_ui(self):
        i18n.retranslate_widget_tree(self)
        for attr in ("repeat_dialog", "measure_dialog", "app_settings_dialog"):
            dlg = getattr(self, attr, None)
            if dlg is not None:
                i18n.retranslate_widget_tree(dlg)

        action_row = self._selected_action_row() if hasattr(self, "actions_table") else None
        fail_row = self._selected_fail_action_row() if hasattr(self, "fail_actions_table") else None
        key_long_row = self._selected_key_long_action_row() if hasattr(self, "key_long_actions_table") else None

        if hasattr(self, "menu_bind_app"):
            self._refresh_bind_app_menu(force=True)
        if hasattr(self, "record_list"):
            self._refresh_record_list()
        if hasattr(self, "actions_table"):
            self._refresh_actions(select_row=action_row)
        if hasattr(self, "fail_actions_table"):
            self._refresh_fail_actions(select_row=fail_row)
        if hasattr(self, "key_long_actions_table"):
            self._refresh_key_long_actions(select_row=key_long_row)
        if hasattr(self, "lw_app_history"):
            self._refresh_app_settings_list()
        if hasattr(self, "btn_bind_base"):
            self._refresh_global_buttons()
        if hasattr(self, "repeat_box"):
            self._refresh_repeat_ui()
        if hasattr(self, "actions_table"):
            self._on_action_selected()

        for attr in (
            "key_press_mode_slider",
            "wait_mode_slider",
            "area_mode_slider",
            "key_long_activation_slider",
        ):
            widget = getattr(self, attr, None)
            if widget is not None:
                widget.update()

        self.update()

    def _apply_ui_language(
        self,
        language: Optional[str],
        *,
        persist: bool = True,
        force_retranslate: bool = False,
    ):
        prev_lang = getattr(self, "_ui_language", i18n.DEFAULT_LANGUAGE)
        lang = i18n.set_language(language)
        self._ui_language = lang

        if hasattr(self, "app_lang_group"):
            self._set_single_choice_selector(self.app_lang_group, lang, default_code=i18n.DEFAULT_LANGUAGE)

        ui_ready = bool(getattr(self, "_ui_ready_for_language_change", False))
        if ui_ready and (force_retranslate or prev_lang != lang):
            self._retranslate_all_ui()

        if persist and ui_ready and not bool(getattr(self, "_suspend_settings_autosave", False)):
            self._save_settings()

    def _on_app_language_changed(self):
        if not hasattr(self, "app_lang_group"):
            return
        lang = self._get_single_choice_selector(self.app_lang_group, default_code=i18n.DEFAULT_LANGUAGE)
        self._apply_ui_language(lang, persist=True)

    def _load_installed_ocr_languages(self, force_refresh: bool = False) -> List[str]:
        langs = [str(x).strip().lower() for x in get_installed_ocr_languages(force_refresh=force_refresh) if str(x).strip()]
        self._ocr_lang_codes = langs
        return list(self._ocr_lang_codes)

    def _default_ocr_lang(self, available: Optional[List[str]] = None) -> str:
        avail = [str(x).strip().lower() for x in (available or []) if str(x).strip()]
        if not avail:
            avail = self._load_installed_ocr_languages(force_refresh=False)
        if "rus" in avail:
            return "rus"
        if avail:
            return avail[0]
        return "rus"

    def _normalize_ocr_lang(self, lang: str, available: Optional[List[str]] = None) -> str:
        avail = [str(x).strip().lower() for x in (available or []) if str(x).strip()]
        if not avail:
            avail = self._load_installed_ocr_languages(force_refresh=False)
        return normalize_ocr_lang_spec(lang, available=avail)

    def _ocr_lang_display_name(self, code: str) -> str:
        key = str(code or "").strip().lower()
        pretty = {
            "rus": "Русский",
            "eng": "English",
            "ukr": "Українська",
            "bel": "Беларуская",
            "deu": "Deutsch",
            "fra": "Français",
            "spa": "Español",
            "ita": "Italiano",
            "por": "Português",
            "pol": "Polski",
            "tur": "Türkçe",
            "nld": "Nederlands",
            "ces": "Čeština",
            "slk": "Slovenčina",
            "slv": "Slovenščina",
            "hun": "Magyar",
            "ron": "Română",
            "fin": "Suomi",
            "swe": "Svenska",
            "dan": "Dansk",
            "nor": "Norsk",
            "ell": "Ελληνικά",
            "bul": "Български",
            "srp": "Српски",
            "hrv": "Hrvatski",
            "lit": "Lietuvių",
            "lav": "Latviešu",
            "est": "Eesti",
            "ara": "العربية",
            "heb": "עברית",
            "hin": "हिन्दी",
            "ben": "বাংলা",
            "tam": "தமிழ்",
            "tel": "తెలుగు",
            "jpn": "日本語",
            "kor": "한국어",
            "chi_sim": "中文 (简体)",
            "chi_tra": "中文 (繁體)",
            "vie": "Tiếng Việt",
            "tha": "ไทย",
            "ind": "Bahasa Indonesia",
            "msa": "Bahasa Melayu",
        }.get(key, "")
        return pretty if pretty else key

    def _clear_layout_widgets(self, layout: Optional[QLayout]):
        if layout is None:
            return
        while layout.count():
            item = layout.takeAt(0)
            child_widget = item.widget()
            child_layout = item.layout()
            if child_layout is not None:
                self._clear_layout_widgets(child_layout)
                child_layout.deleteLater()
            if child_widget is not None:
                child_widget.deleteLater()

    def _rebuild_ocr_lang_selector(
        self,
        row: QWidget,
        on_change,
        selected_lang: Optional[str] = None,
    ) -> Optional[QButtonGroup]:
        row_l = row.layout()
        if row_l is None:
            row_l = QHBoxLayout(row)
            row_l.setContentsMargins(0, 0, 0, 0)
            row_l.setSpacing(6)
        else:
            self._clear_layout_widgets(row_l)

        installed = self._load_installed_ocr_languages(force_refresh=False)
        if not installed:
            btn_help = QPushButton("Инструкция установки")
            btn_help.setCursor(Qt.PointingHandCursor)
            btn_help.setStyleSheet(
                "QPushButton { background: #612121; border: 1px solid #7e2d2d; color: #f4dede; }"
                "QPushButton:hover { background: #742828; }"
                "QPushButton:pressed { background: #4f1a1a; }"
                "QPushButton:focus { outline: none; }"
            )
            btn_help.clicked.connect(self._show_ocr_install_help_dialog)
            row_l.addWidget(btn_help)
            row_l.addStretch(1)
            return None

        group = QButtonGroup(row)
        group.setExclusive(False)

        for code in installed:
            btn = QToolButton(row)
            btn.setObjectName("lang_chip")
            btn.setText(self._ocr_lang_display_name(code))
            btn.setCheckable(True)
            btn.setAutoRaise(False)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setMinimumHeight(30)
            btn.setProperty("lang_code", code)
            btn.toggled.connect(
                lambda checked, g=group, b=btn, cb=on_change: self._on_ocr_lang_selector_toggled(g, b, checked, cb)
            )
            group.addButton(btn)
            row_l.addWidget(btn)

        self._set_ocr_lang_selector(group, selected_lang or self._default_ocr_lang(installed))
        row_l.addStretch(1)
        return group

    def _refresh_ocr_lang_selectors(self, force_refresh: bool = False) -> bool:
        if force_refresh:
            self._load_installed_ocr_languages(force_refresh=True)
        else:
            self._load_installed_ocr_languages(force_refresh=False)

        has_langs = bool(self._ocr_lang_codes)
        if hasattr(self, "area_ocr_lang_row"):
            prev = self._get_ocr_lang_selector(getattr(self, "area_ocr_lang_group", None))
            self.area_ocr_lang_group = self._rebuild_ocr_lang_selector(
                self.area_ocr_lang_row, self._apply_area_params, selected_lang=prev
            )
        if hasattr(self, "wait_event_ocr_lang_row"):
            prev = self._get_ocr_lang_selector(getattr(self, "wait_event_ocr_lang_group", None))
            self.wait_event_ocr_lang_group = self._rebuild_ocr_lang_selector(
                self.wait_event_ocr_lang_row, self._apply_wait_event_params, selected_lang=prev
            )

        if hasattr(self, "actions_table"):
            self._on_action_selected()
        return has_langs

    def _build_ocr_lang_selector(self, on_change) -> Tuple[QWidget, Optional[QButtonGroup]]:
        row = QWidget()
        row_l = QHBoxLayout(row)
        row_l.setContentsMargins(0, 0, 0, 0)
        row_l.setSpacing(6)
        group = self._rebuild_ocr_lang_selector(row, on_change, selected_lang=self._default_ocr_lang())
        return row, group

    def _show_ocr_install_help_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Настройка OCR (Tesseract)")
        dlg.setModal(True)
        dlg.setMinimumWidth(640)
        self._prepare_dark_dialog(dlg)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        title = QLabel("Для работы OCR установите Tesseract и языковые пакеты.")
        title.setWordWrap(True)
        lay.addWidget(title)

        steps = QLabel(
            "1. Скачайте и установите последнюю версию Tesseract OCR.\n"
            "2. Во время установки выберите нужные языки OCR.\n"
            "3. Запомните путь к папке установки (обычно содержит tesseract.exe).\n"
            "4. Откройте параметры Windows: «Переменные среды».\n"
            "5. В переменной Path (для пользователя или системы) добавьте путь к папке Tesseract.\n"
            "6. Откройте новую консоль (Win+R -> cmd).\n"
            "7. Выполните команду: tesseract --list-langs\n"
            "8. Убедитесь, что в списке есть установленные языки.\n"
            "9. Нажмите «Закрыть приложение» ниже и запустите приложение снова вручную."
        )
        steps.setWordWrap(True)
        steps.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(steps)

        row = QHBoxLayout()
        row.addStretch(1)

        btn_skip = QPushButton("Пропустить")
        btn_skip.setStyleSheet(
            "QPushButton { background: #612121; border: 1px solid #7e2d2d; color: #f4dede; }"
            "QPushButton:hover { background: #742828; }"
            "QPushButton:pressed { background: #4f1a1a; }"
            "QPushButton:focus { outline: none; }"
        )
        btn_skip.clicked.connect(dlg.reject)
        row.addWidget(btn_skip)

        btn_close_app = QPushButton("Закрыть приложение")
        btn_close_app.setStyleSheet(
            "QPushButton { background: #1f5a35; border: 1px solid #2d7747; color: #dff4e7; }"
            "QPushButton:hover { background: #266d41; }"
            "QPushButton:pressed { background: #184a2b; }"
            "QPushButton:focus { outline: none; }"
        )

        def on_close_app():
            dlg.accept()
            app = QApplication.instance()
            if app is not None:
                app.quit()

        btn_close_app.clicked.connect(on_close_app)
        row.addWidget(btn_close_app)
        lay.addLayout(row)

        i18n.retranslate_widget_tree(dlg)
        dlg.exec()

    def _on_ocr_lang_selector_toggled(
        self,
        group: Optional[QButtonGroup],
        changed_btn: Optional[QToolButton],
        checked: bool,
        on_change,
    ):
        if group is None or changed_btn is None:
            if callable(on_change):
                on_change()
            return

        # Не даём снять последнюю активную кнопку языка.
        if not checked:
            has_any_other_checked = any(
                b is not changed_btn and b.isChecked()
                for b in group.buttons()
            )
            if not has_any_other_checked:
                changed_btn.blockSignals(True)
                changed_btn.setChecked(True)
                changed_btn.blockSignals(False)

        if callable(on_change):
            on_change()

    def _set_ocr_lang_selector(self, group: Optional[QButtonGroup], lang: str):
        if group is None:
            return
        buttons = [b for b in group.buttons() if isinstance(b, QToolButton)]
        if not buttons:
            return

        available = []
        for btn in buttons:
            code = str(btn.property("lang_code") or "").strip().lower()
            if code:
                available.append(code)

        norm = self._normalize_ocr_lang(lang, available=available)
        targets = [x for x in str(norm).split("+") if x]
        if not targets:
            targets = [self._default_ocr_lang(available)]
        targets_set = set(targets)

        for btn in buttons:
            btn.blockSignals(True)
        for btn in buttons:
            code = str(btn.property("lang_code") or "").strip().lower()
            btn.setChecked(code in targets_set)
        if not any(btn.isChecked() for btn in buttons):
            fallback = self._default_ocr_lang(available)
            for btn in buttons:
                code = str(btn.property("lang_code") or "").strip().lower()
                if code == fallback:
                    btn.setChecked(True)
                    break
        for btn in buttons:
            btn.blockSignals(False)

    def _get_ocr_lang_selector(self, group: Optional[QButtonGroup]) -> str:
        if group is None:
            return self._default_ocr_lang()

        selected: List[str] = []
        available: List[str] = []
        for btn in group.buttons():
            code = str(btn.property("lang_code") or "").strip().lower()
            if not code:
                continue
            available.append(code)
            if btn.isChecked():
                selected.append(code)

        if not selected:
            selected = [self._default_ocr_lang(available)]
        return "+".join(selected)

    def _build_single_choice_selector(
        self,
        items: List[Tuple[str, str]],
        on_change,
        default_code: str,
    ) -> Tuple[QWidget, QButtonGroup]:
        row = QWidget()
        row_l = QHBoxLayout(row)
        row_l.setContentsMargins(0, 0, 0, 0)
        row_l.setSpacing(6)

        group = QButtonGroup(row)
        group.setExclusive(True)

        first_btn = None
        default_btn = None
        for code, text in items:
            btn = QToolButton(row)
            btn.setObjectName("lang_chip")
            btn.setText(str(text))
            btn.setCheckable(True)
            btn.setAutoRaise(False)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setMinimumHeight(30)
            btn.setProperty("choice_code", str(code))
            btn.toggled.connect(
                lambda checked, cb=on_change: (cb() if checked and callable(cb) else None)
            )
            group.addButton(btn)
            row_l.addWidget(btn)
            if first_btn is None:
                first_btn = btn
            if str(code) == str(default_code):
                default_btn = btn

        if default_btn is None:
            default_btn = first_btn
        if default_btn is not None:
            default_btn.setChecked(True)

        row_l.addStretch(1)
        return row, group

    def _set_single_choice_selector(self, group: Optional[QButtonGroup], code: str, default_code: str):
        if group is None:
            return
        target_code = str(code or default_code)
        target_btn = None
        first_btn = None
        for btn in group.buttons():
            if first_btn is None:
                first_btn = btn
            if str(btn.property("choice_code") or "") == target_code:
                target_btn = btn
        if target_btn is None:
            target_btn = first_btn
        if target_btn is None:
            return
        for btn in group.buttons():
            btn.blockSignals(True)
        target_btn.setChecked(True)
        for btn in group.buttons():
            btn.blockSignals(False)

    def _get_single_choice_selector(self, group: Optional[QButtonGroup], default_code: str) -> str:
        if group is None:
            return str(default_code)
        btn = group.checkedButton()
        if btn is None:
            return str(default_code)
        code = str(btn.property("choice_code") or "").strip()
        return code if code else str(default_code)

    def _is_our_window_widget(self, widget: Optional[QWidget]) -> bool:
        if widget is None:
            return False
        win = widget.window()
        if win is self:
            return True
        for attr_name in ("repeat_dialog", "measure_dialog", "app_settings_dialog"):
            dlg = getattr(self, attr_name, None)
            if isinstance(dlg, QDialog) and win is dlg:
                return True
        return False

    def _resolve_edit_widget(self, widget: Optional[QWidget]) -> Optional[QWidget]:
        cur = widget
        while cur is not None:
            if isinstance(cur, QLineEdit):
                return cur
            if isinstance(cur, QAbstractSpinBox):
                return cur
            if isinstance(cur, QComboBox) and cur.isEditable():
                return cur
            cur = cur.parentWidget()
        return None

    def _clear_edit_focus(self, widget: Optional[QWidget]):
        edit = self._resolve_edit_widget(widget)
        if edit is None:
            return

        if isinstance(edit, QAbstractSpinBox):
            le = edit.lineEdit()
            if le is not None:
                le.deselect()
        elif isinstance(edit, QLineEdit):
            edit.deselect()

        edit.clearFocus()
        win = edit.window()
        if isinstance(win, QWidget):
            win.setFocus(Qt.ActiveWindowFocusReason)

    def _step_spinbox(self, spin: QAbstractSpinBox, step: int):
        if spin is None:
            return
        if int(step) < 0:
            spin.stepDown()
        else:
            spin.stepUp()

        le = spin.lineEdit()
        if le is not None:
            le.deselect()

        btn = self.sender()
        if isinstance(btn, QWidget):
            btn.setFocus(Qt.MouseFocusReason)

    def eventFilter(self, obj, event):
        etype = event.type()
        if etype not in (QEvent.MouseButtonPress, QEvent.KeyPress):
            return super().eventFilter(obj, event)

        obj_widget = obj if isinstance(obj, QWidget) else None
        if obj_widget is None or not self._is_our_window_widget(obj_widget):
            return super().eventFilter(obj, event)

        focused = QApplication.focusWidget()
        if not isinstance(focused, QWidget) or not self._is_our_window_widget(focused):
            return super().eventFilter(obj, event)

        focused_edit = self._resolve_edit_widget(focused)
        if focused_edit is None:
            return super().eventFilter(obj, event)

        if etype == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self._clear_edit_focus(focused_edit)
            return True

        if etype == QEvent.MouseButtonPress:
            clicked_edit = self._resolve_edit_widget(obj_widget)
            if clicked_edit is None:
                self._clear_edit_focus(focused_edit)

        return super().eventFilter(obj, event)

    def _normalize_bound_app_lists(self):
        seen = set()
        recent = []
        for raw in self._bound_exe_recent:
            exe = str(raw or "").strip()
            if not exe:
                continue
            key = os.path.normcase(exe)
            if key in seen:
                continue
            seen.add(key)
            recent.append(exe)
        self._bound_exe_recent = recent

        seen_fav = set()
        favs = []
        for raw in self._bound_exe_favorites:
            exe = str(raw or "").strip()
            if not exe:
                continue
            key = os.path.normcase(exe)
            if key in seen_fav:
                continue
            if key not in seen:
                continue
            seen_fav.add(key)
            favs.append(exe)
        self._bound_exe_favorites = favs[:3]

        if self._bound_exe:
            cur = str(self._bound_exe).strip()
            self._bound_exe = cur
            if cur:
                cur_key = os.path.normcase(cur)
                if cur_key not in seen:
                    self._bound_exe_recent.insert(0, cur)

    def _ordered_bound_apps(self) -> List[str]:
        self._normalize_bound_app_lists()
        fav_keys = {os.path.normcase(x) for x in self._bound_exe_favorites}
        rest = [x for x in self._bound_exe_recent if os.path.normcase(x) not in fav_keys]
        return list(self._bound_exe_favorites) + rest

    def _is_bound_favorite(self, exe: str) -> bool:
        key = os.path.normcase(str(exe or ""))
        return any(os.path.normcase(x) == key for x in self._bound_exe_favorites)

    def _bound_context_for_record(self, rec: Optional[Record] = None, notify_missing: bool = False) -> Dict[str, Any]:
        rec = rec if rec is not None else self._current_record()

        prefix = "Приложение"
        mode = "global"
        display_exe = str(self._bound_exe or "").strip()

        if rec and bool(getattr(rec, "bind_to_process", False)):
            bound_exe = str(getattr(rec, "bound_exe", "") or "").strip()
            override = str(getattr(rec, "bound_exe_override", "") or "").strip()
            if override:
                prefix = "Временное приложение"
                mode = "record_temp"
                display_exe = override
            elif bound_exe:
                prefix = "Привязанное приложение"
                mode = "record_bind"
                display_exe = bound_exe
            else:
                prefix = "Привязанное приложение"
                mode = "record_missing"
                display_exe = ""
                if notify_missing:
                    self._set_status("Для записи не задано привязанное приложение.", level="error")

        enabled = bool(self._bound_exe_enabled)
        effective_exe = display_exe if enabled else ""
        return {
            "prefix": prefix,
            "mode": mode,
            "display_exe": display_exe,
            "effective_exe": effective_exe,
            "enabled": enabled,
        }

    def _remember_bound_exe_in_history(self, exe: str):
        exe = (exe or "").strip()
        if not exe:
            return
        key = os.path.normcase(exe)
        self._bound_exe_recent = [x for x in self._bound_exe_recent if os.path.normcase(str(x)) != key]
        self._bound_exe_recent.insert(0, exe)
        self._normalize_bound_app_lists()

    def _remember_bound_exe(self, exe: str):
        exe = (exe or "").strip()
        if not exe:
            return

        self._bound_exe = exe
        self._bound_exe_enabled = True
        self._remember_bound_exe_in_history(exe)

    def _set_temporary_bound_exe_for_current_record(self, exe: str) -> bool:
        rec = self._current_record()
        if not rec or not bool(getattr(rec, "bind_to_process", False)):
            return False

        exe = (exe or "").strip()
        if not exe:
            return False

        if not str(getattr(rec, "bound_exe", "") or "").strip():
            rec.bound_exe = str(self._bound_exe or "").strip() or exe
        rec.bound_exe_override = exe
        self._bound_exe_enabled = True
        self._remember_bound_exe_in_history(exe)
        self._refresh_bound_process_caption(rec)
        return True

    def _clear_current_record_temporary_bound_exe(self):
        rec = self._current_record()
        if not rec:
            self._set_status("Сначала выберите запись.", level="error")
            return
        if not bool(getattr(rec, "bind_to_process", False)):
            self._set_status("У выбранной записи нет привязки к процессу.")
            return
        if not str(getattr(rec, "bound_exe_override", "") or "").strip():
            self._set_status("Временное приложение не выбрано.")
            return

        rec.bound_exe_override = ""
        self._save()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=str(getattr(rec, "bound_exe", "") or "").strip())
        self._refresh_bound_process_caption(rec)
        self._set_status("Временное приложение для записи убрано.")

    def _bind_current_app_to_record(self):
        rec = self._current_record()
        if not rec:
            self._set_status("Сначала выберите запись.", level="error")
            return

        ctx = self._bound_context_for_record(rec=rec, notify_missing=False)
        exe = str(ctx.get("display_exe", "") or "").strip()
        if not exe:
            exe = str(self._bound_exe or "").strip()
        if not exe:
            self._set_status("Сначала выберите приложение.", level="error")
            return

        rec.bind_to_process = True
        rec.bound_exe = exe
        rec.bound_exe_override = ""
        self._remember_bound_exe_in_history(exe)
        self._save()
        self._refresh_repeat_ui()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=exe)
        self._set_status(f"Приложение привязано к записи: {os.path.basename(exe)}")

    def _detach_current_record_binding(self):
        rec = self._current_record()
        if not rec:
            self._set_status("Сначала выберите запись.", level="error")
            return

        had_binding = bool(
            getattr(rec, "bind_to_process", False)
            or str(getattr(rec, "bound_exe", "") or "").strip()
            or str(getattr(rec, "bound_exe_override", "") or "").strip()
        )
        rec.bind_to_process = False
        rec.bound_exe = ""
        rec.bound_exe_override = ""
        self._save()
        self._refresh_repeat_ui()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=self._bound_exe)
        if had_binding:
            self._set_status("Привязка записи к процессу снята.")
        else:
            self._set_status("У выбранной записи нет привязки к процессу.")

    def _select_bound_exe_from_history(self, exe: str):
        exe = (exe or "").strip()
        if not exe:
            return

        rec = self._current_record()
        if rec and bool(getattr(rec, "bind_to_process", False)):
            self._set_temporary_bound_exe_for_current_record(exe)
            self._save()
            self._save_settings()
            self._refresh_global_buttons()
            self._refresh_app_settings_list(select_exe=exe)
            self._set_status(f"Временное приложение для записи: {os.path.basename(exe)}")
            return

        self._remember_bound_exe(exe)
        self._save_settings()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=exe)
        self._set_status(f"Приложение выбрано: {os.path.basename(exe)}")

    def _toggle_bound_exe_enabled(self):
        ctx = self._bound_context_for_record()
        target = str(ctx.get("display_exe", "") or "").strip()
        if not target:
            self._set_status("Сначала задайте приложение.", level="error")
            return
        self._bound_exe_enabled = not self._bound_exe_enabled
        self._save_settings()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=target)
        state = "включено" if self._bound_exe_enabled else "выключено"
        prefix = str(ctx.get("prefix", "Приложение"))
        self._set_status(f"{prefix} {state}: {os.path.basename(target)}")

    def _app_settings_selected_exe(self) -> str:
        if not hasattr(self, "lw_app_history"):
            return ""
        it = self.lw_app_history.currentItem()
        if not it:
            return ""
        return str(it.data(Qt.UserRole) or "").strip()

    def _refresh_app_settings_list(self, select_exe: Optional[str] = None):
        if not hasattr(self, "lw_app_history"):
            return

        ordered = self._ordered_bound_apps()
        ctx = self._bound_context_for_record()
        active_exe = str(ctx.get("display_exe", "") or "").strip()

        self.lw_app_history.blockSignals(True)
        self.lw_app_history.clear()
        for exe in ordered:
            base = os.path.basename(exe) or exe
            fav = self._is_bound_favorite(exe)
            cur = bool(active_exe) and os.path.normcase(exe) == os.path.normcase(active_exe)
            text = f"{'★ ' if fav else ''}{base}{' (текущее)' if cur else ''}"
            it = QListWidgetItem(text)
            it.setData(Qt.UserRole, exe)
            it.setToolTip(exe)
            self.lw_app_history.addItem(it)
        self.lw_app_history.blockSignals(False)

        target = (select_exe or active_exe).strip()
        if target:
            target_key = os.path.normcase(target)
            for i in range(self.lw_app_history.count()):
                it = self.lw_app_history.item(i)
                if os.path.normcase(str(it.data(Qt.UserRole) or "")) == target_key:
                    self.lw_app_history.setCurrentRow(i)
                    break
        elif self.lw_app_history.count() > 0:
            self.lw_app_history.setCurrentRow(0)

        self._on_app_settings_selection_changed()

    def _on_app_settings_selection_changed(self):
        if not hasattr(self, "btn_app_toggle_fav"):
            return
        exe = self._app_settings_selected_exe()
        has = bool(exe)
        is_fav = self._is_bound_favorite(exe) if has else False

        self.btn_app_toggle_fav.setEnabled(has)
        self.btn_app_toggle_fav.setText("Убрать из избранного" if is_fav else "Добавить в избранное")

        can_move = False
        can_up = False
        can_down = False
        if is_fav:
            keys = [os.path.normcase(x) for x in self._bound_exe_favorites]
            idx = keys.index(os.path.normcase(exe))
            can_move = True
            can_up = idx > 0
            can_down = idx < (len(keys) - 1)

        self.btn_app_fav_up.setEnabled(can_move and can_up)
        self.btn_app_fav_down.setEnabled(can_move and can_down)

    def _toggle_app_favorite(self):
        exe = self._app_settings_selected_exe()
        if not exe:
            return
        key = os.path.normcase(exe)
        fav_keys = [os.path.normcase(x) for x in self._bound_exe_favorites]

        if key in fav_keys:
            idx = fav_keys.index(key)
            self._bound_exe_favorites.pop(idx)
            msg = "Убрано из избранного."
        else:
            if len(self._bound_exe_favorites) >= 3:
                self._set_status("В избранном можно хранить максимум 3 приложения.", level="error")
                return
            self._bound_exe_favorites.append(exe)
            msg = "Добавлено в избранное."

        self._normalize_bound_app_lists()
        self._save_settings()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=exe)
        self._set_status(msg)

    def _move_app_favorite(self, direction: int):
        exe = self._app_settings_selected_exe()
        if not exe:
            return
        fav_keys = [os.path.normcase(x) for x in self._bound_exe_favorites]
        key = os.path.normcase(exe)
        if key not in fav_keys:
            return
        i = fav_keys.index(key)
        j = i + int(direction)
        if j < 0 or j >= len(self._bound_exe_favorites):
            return
        self._bound_exe_favorites[i], self._bound_exe_favorites[j] = self._bound_exe_favorites[j], self._bound_exe_favorites[i]
        self._save_settings()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=exe)

    def _clear_non_favorite_apps(self):
        self._normalize_bound_app_lists()
        fav = list(self._bound_exe_favorites)
        self._bound_exe_recent = [x for x in fav]

        if self._bound_exe:
            cur_key = os.path.normcase(self._bound_exe)
            fav_keys = {os.path.normcase(x) for x in fav}
            if cur_key not in fav_keys:
                self._bound_exe = fav[0] if fav else ""
                if not self._bound_exe:
                    self._bound_exe_enabled = False

        self._save_settings()
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=self._bound_exe)
        self._set_status("Список очищен: оставлены только избранные приложения.")

    def _refresh_bind_app_menu(self, force: bool = False):
        if not hasattr(self, "menu_bind_app"):
            return

        if self.menu_bind_app.isVisible() and not force:
            return

        ordered = self._ordered_bound_apps()
        ctx = self._bound_context_for_record()
        active_exe = str(ctx.get("display_exe", "") or "").strip()
        bind_source_exe = active_exe or str(self._bound_exe or "").strip()
        rec = self._current_record()

        snapshot = (
            tuple(ordered[:4]),
            tuple(self._bound_exe_favorites),
            active_exe,
            str(ctx.get("mode", "")),
            bool(self._bound_exe_enabled),
            bool(rec and getattr(rec, "bind_to_process", False)),
            str(getattr(rec, "bound_exe_override", "") if rec else ""),
        )
        if not force and snapshot == getattr(self, "_bind_app_menu_snapshot", None):
            return
        self._bind_app_menu_snapshot = snapshot

        self.menu_bind_app.clear()

        act_set = self.menu_bind_app.addAction("Выбрать приложение")
        act_bind_record = self.menu_bind_app.addAction("Привязать приложение к записи")
        act_detach = self.menu_bind_app.addAction("Отвязать от записи")
        act_toggle = self.menu_bind_app.addAction("Включить/выключить")

        act_set.triggered.connect(self._bind_base_by_click)
        act_bind_record.triggered.connect(self._bind_current_app_to_record)
        act_detach.triggered.connect(self._detach_current_record_binding)
        act_toggle.triggered.connect(self._toggle_bound_exe_enabled)
        act_bind_record.setEnabled(rec is not None and bool(bind_source_exe))
        act_detach.setEnabled(rec is not None)

        if ordered:
            self.menu_bind_app.addSeparator()

        for exe in ordered[:4]:
            base = os.path.basename(exe) or exe
            cur = bool(active_exe) and os.path.normcase(exe) == os.path.normcase(active_exe)
            fav = self._is_bound_favorite(exe)
            suffix = " (текущее)" if cur else ""
            text = f"{'★ ' if fav else ''}{base}{suffix}"
            act = self.menu_bind_app.addAction(text)
            act.setToolTip(exe)
            act.triggered.connect(lambda _=False, x=exe: self._select_bound_exe_from_history(x))

        if str(ctx.get("mode", "")) == "record_temp":
            self.menu_bind_app.addSeparator()
            act_clear_temp = self.menu_bind_app.addAction("Убрать выбор временного приложения")
            act_clear_temp.triggered.connect(self._clear_current_record_temporary_bound_exe)

    def _bind_base_by_click(self):
        if not _is_windows():
            self._set_status("Привязка приложения работает только на Windows.", level="error")
            return

        # Нельзя ловить клик "под модальным окном", поэтому сначала показываем инструкцию.
        accepted = self._show_dark_info_dialog(
            "Выбрать приложение",
            "После нажатия «Выбрать приложение» интерфейс будет временно скрыт.\n"
            "Кликните левой кнопкой мыши по окну нужного приложения.\n"
            "Нажмите Esc для отмены выбора и возврата в приложение.",
            accept_text="Выбрать приложение",
            reject_text="Отменить действие",
        )
        if not accepted:
            self._set_status("Выбор приложения отменён.")
            return

        try:
            from pynput import keyboard, mouse
        except Exception as ex:
            QMessageBox.warning(self, "Ошибка", f"Нет pynput: {ex}")
            return

        windows_to_track = [self]
        for attr in ("repeat_dialog", "measure_dialog", "app_settings_dialog"):
            w = getattr(self, attr, None)
            if w is not None:
                windows_to_track.append(w)

        visible_before = {w: bool(w.isVisible()) for w in windows_to_track}
        for w, was_visible in visible_before.items():
            if was_visible:
                w.hide()
        QApplication.processEvents()

        result = {"exe": "", "hwnd": 0, "cancelled": False}
        loop = QEventLoop()

        def _finish_pick():
            try:
                if loop.isRunning():
                    loop.quit()
            except Exception:
                pass

        def on_click(x, y, button, pressed):
            if not pressed:
                return
            if str(button).endswith("left"):
                hwnd = _win_hwnd_from_point(int(x), int(y))
                pid = _win_pid_from_hwnd(hwnd)
                if pid == os.getpid():
                    return
                exe = _win_exe_from_pid(pid)
                if exe:
                    result["exe"] = exe
                    result["hwnd"] = hwnd
                    _finish_pick()
                    return False

        def on_key_press(key):
            try:
                if key == keyboard.Key.esc:
                    result["cancelled"] = True
                    _finish_pick()
                    return False
            except Exception:
                pass

        ms_listener = mouse.Listener(on_click=on_click)
        kb_listener = keyboard.Listener(on_press=on_key_press)
        ms_listener.daemon = True
        kb_listener.daemon = True
        ms_listener.start()
        kb_listener.start()

        try:
            loop.exec()
        finally:
            for l in (ms_listener, kb_listener):
                try:
                    l.stop()
                except Exception:
                    pass

            # Возвращаем окно приложения (и ранее открытые доп. окна) после выбора/отмены.
            for w, was_visible in visible_before.items():
                if was_visible:
                    try:
                        w.show()
                    except Exception:
                        pass
            if visible_before.get(self, False):
                self.raise_()
                self.activateWindow()
            for attr in ("repeat_dialog", "measure_dialog", "app_settings_dialog"):
                w = getattr(self, attr, None)
                if w is not None and visible_before.get(w, False):
                    self._schedule_dark_titlebar(w)
            self._schedule_dark_titlebar(self)

        if result["cancelled"]:
            self._set_status("Выбор приложения отменён (Esc).")
            return

        exe = (result["exe"] or "").strip()
        if not exe:
            self._set_status("Не удалось определить приложение. Выберите видимое окно приложения.", level="error")
            return

        if self._set_temporary_bound_exe_for_current_record(exe):
            self._save()
            self._save_settings()
            self._refresh_global_buttons()
            self._refresh_app_settings_list(select_exe=exe)
            self._set_status(f"Временное приложение для записи: {os.path.basename(exe)}")
            return

        self._remember_bound_exe(exe)
        self._save_settings()

        # обновим UI
        self._refresh_global_buttons()
        self._refresh_app_settings_list(select_exe=exe)

        # опционально: обновим базовую область во всех записях (чтобы в списке было "красиво")
        base = resolve_bound_base_rect_dip(exe)
        if base:
            for rec in self.records:
                if rec.actions and isinstance(rec.actions[0], dict) and rec.actions[0].get("type") == "base_area":
                    ba = BaseAreaAction.from_dict(rec.actions[0])
                    ba.x1, ba.y1, ba.x2, ba.y2 = base.left(), base.top(), base.right(), base.bottom()
                    rec.actions[0] = ba.to_dict()
            self._save()

        self._set_status(f"Приложение привязано: {os.path.basename(exe)}")

    def _stop_word_set(self):
        # выбираем область + слово + индекс, сохраняем как WordAreaAction-подобный dict
        base = None
        ctx = self._bound_context_for_record()
        effective_exe = str(ctx.get("effective_exe", "") or "").strip()
        if effective_exe:
            base = resolve_bound_base_rect_dip(effective_exe)
        overlay = AreaSelectOverlay(initial_global=None)
        rect = overlay.show_and_block()
        if not rect or not rect.isValid():
            return

        word, ok = self._input_text_dark("Стоп-слово", "Какое слово искать?")
        if not ok:
            return
        word = str(word or "")
        if not word.strip():
            return

        idx, ok = self._input_int_dark(
            "Стоп-слово", "Какое по счёту совпадение? (1 = первое)", 1, 1, 9999, 1
        )
        if not ok:
            return

        sec_default = int(config.STOP_WORD_POLL_SEC)
        sec, ok = self._input_int_dark(
            "Стоп-слово",
            "Проверять каждые N секунд:",
            sec_default,
            1,
            9999,
            1,
        )
        if not ok:
            return

        wa = WordAreaAction(
            word=word,
            index=max(1, int(idx)),
            click=False,
            button="left",
            ocr_lang=self._default_ocr_lang(),
        )
        wa.count = 0
        wa.trigger = dict(DEFAULT_TRIGGER)

        if base and base.isValid():
            rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)
            wa.coord = "rel"
            wa.rx1, wa.ry1, wa.rx2, wa.ry2 = rx1, ry1, rx2, ry2
        else:
            wa.coord = "abs"
            wa.x1, wa.y1, wa.x2, wa.y2 = rect.left(), rect.top(), rect.right(), rect.bottom()

        self._stop_word_cfg = wa.to_dict()
        self._stop_word_cfg["interval_sec"] = int(sec)

        # по умолчанию включаем
        self._stop_word_enabled_event.set()

        self._save_settings()
        self._refresh_global_buttons()
        self._set_status("Стоп-слово задано и включено.")

    def _stop_word_toggle(self):
        if not self._stop_word_cfg:
            self._set_status("Стоп-слово ещё не задано.")
            return
        if self._stop_word_enabled_event.is_set():
            self._stop_word_enabled_event.clear()
            self._set_status("Стоп-слово: выключено.")
        else:
            self._stop_word_enabled_event.set()
            self._set_status("Стоп-слово: включено.")

        self._save_settings()
        self._refresh_global_buttons()

    def _stop_word_clear(self):
        self._stop_word_cfg = None
        self._stop_word_enabled_event.clear()
        self._save_settings()
        self._refresh_global_buttons()
        self._set_status("Стоп-слово убрано.")

    def _apply_dark_titlebar_widget(self, widget: QWidget):
        try:
            if not sys.platform.startswith("win"):
                return

            if widget is None:
                return

            widget.setAttribute(Qt.WA_NativeWindow, True)
            hwnd = int(widget.winId())
            if not hwnd:
                return

            for attr in (20, 19):
                value = ctypes.c_int(1)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
                )

            # Явно задаём цвета non-client части (Win10/11), чтобы не откатывалось в белый.
            caption_color = ctypes.c_uint(0x0015110F)  # #0f1115 (COLORREF BGR)
            text_color = ctypes.c_uint(0x00F7EEE9)     # #e9eef7 (COLORREF BGR)
            border_color = ctypes.c_uint(0x00362A24)   # #242a36 (COLORREF BGR)
            for attr, value in ((35, caption_color), (36, text_color), (34, border_color)):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
                )

            SWP_NOSIZE = 0x0001
            SWP_NOMOVE = 0x0002
            SWP_NOZORDER = 0x0004
            SWP_FRAMECHANGED = 0x0020
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
            )

            # NEW: принудительная перерисовка рамки/заголовка
            RDW_INVALIDATE = 0x0001
            RDW_FRAME = 0x0400
            RDW_UPDATENOW = 0x0100
            ctypes.windll.user32.RedrawWindow(
                hwnd, None, None, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW
            )
        except Exception:
            pass

    def _schedule_dark_titlebar(self, widget: QWidget):
        if not sys.platform.startswith("win") or widget is None:
            return
        for delay_ms in (0, 120, 260):
            QTimer.singleShot(delay_ms, lambda w=widget: self._apply_dark_titlebar_widget(w))

    def _apply_dark_titlebar_windows(self):
        self._schedule_dark_titlebar(self)
        if hasattr(self, "repeat_dialog"):
            self._schedule_dark_titlebar(self.repeat_dialog)
        if hasattr(self, "measure_dialog"):
            self._schedule_dark_titlebar(self.measure_dialog)
        if hasattr(self, "app_settings_dialog"):
            self._schedule_dark_titlebar(self.app_settings_dialog)

    def _apply_default_splitter(self):
        # 62% / 38% от текущей ширины окна — “красиво по умолчанию”
        w = max(800, self.width())
        left = int(w * 0.35)
        right = max(300, w - left)
        self.main_splitter.setSizes([left, right])

    def _top_tool_buttons(self):
        return (
            self.btn_record_menu,
            self.btn_current_record,
            self.btn_repeat_settings,
            self.btn_bind_base,
            self.btn_measure_window,
            self.btn_stop_word,
            self.btn_app_settings,
        )

    def _lock_top_buttons_geometry(self):
        buttons = self._top_tool_buttons()
        if not buttons:
            return

        text_h = max((b.fontMetrics().height() for b in buttons), default=0)
        h = max(26, text_h + 8)
        for b in buttons:
            b.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            b.setMinimumHeight(h)
            b.setMaximumHeight(h)
            if b is self.btn_repeat_settings:
                gear_w = max(h, b.fontMetrics().horizontalAdvance("⚙") + 16)
                b.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                b.setMinimumWidth(gear_w)
                b.setMaximumWidth(gear_w)
            else:
                b.setMinimumWidth(0)
                b.setMaximumWidth(16777215)

    def _shrink_top_buttons(self):
        self._lock_top_buttons_geometry()

    def _set_current_record(self, idx: int):
        if not (0 <= idx < len(self.records)):
            self.current_index = -1
            self._last_record_index = -1
            self.btn_current_record.setText("—")
            self.btn_current_record.setEnabled(False)

            self._refresh_record_menu()
            self._refresh_record_list_widget()

            self._refresh_actions()
            self._refresh_repeat_ui()
            self._refresh_global_buttons()
            if not getattr(self, "_suspend_settings_autosave", False):
                self._save_settings()
            return

        self.current_index = idx
        self._last_record_index = idx

        # синхронизируем левый список (если он используется)
        if hasattr(self, "record_list"):
            self.record_list.blockSignals(True)
            self.record_list.setCurrentRow(idx)
            self.record_list.blockSignals(False)

        self.btn_current_record.setEnabled(True)
        self.btn_current_record.setText(self.records[idx].name)

        self._refresh_record_menu()
        self._refresh_record_list_widget()

        self._refresh_actions()
        self._refresh_repeat_ui()
        self._bound_context_for_record(notify_missing=True)
        self._refresh_global_buttons()
        if not getattr(self, "_suspend_settings_autosave", False):
            self._save_settings()

    def _ensure_base_and_migrate(self, rec: Record):

        # если нет ABS областей/слов (которые надо конвертить) — не трогаем запись
        need_migrate_abs = False
        for ad in rec.actions:
            if not isinstance(ad, dict):
                continue
            if ad.get("type") not in ("area", "area_word"):
                continue
            # уже rel -> не считается
            if any(k in ad for k in ("rx1", "ry1", "rx2", "ry2")):
                continue
            # есть ABS зона -> нужна миграция
            need_migrate_abs = True
            break

        if not need_migrate_abs:
            return  # <-- ВАЖНО: больше не вставляем base_area автоматически

        # 1) если базовой области нет — создаём по bounding box старых abs-координат
        if not (rec.actions and isinstance(rec.actions[0], dict) and rec.actions[0].get("type") == "base_area"):
            minx = miny = 10 ** 9
            maxx = maxy = -10 ** 9
            found = False

            for ad in rec.actions:
                if not isinstance(ad, dict):
                    continue
                if ad.get("type") not in ("area", "area_word"):
                    continue
                if any(k in ad for k in ("rx1", "ry1", "rx2", "ry2")):
                    continue  # уже rel, bbox не нужен

                x1 = int(ad.get("x1", 0));
                y1 = int(ad.get("y1", 0))
                x2 = int(ad.get("x2", 0));
                y2 = int(ad.get("y2", 0))
                minx = min(minx, x1, x2);
                miny = min(miny, y1, y2)
                maxx = max(maxx, x1, x2);
                maxy = max(maxy, y1, y2)
                found = True

            if not found:
                vg = virtual_geometry()
                base_rect = vg
            else:
                base_rect = QRect(QPoint(minx, miny), QPoint(maxx, maxy)).normalized()

            base = BaseAreaAction(
                x1=base_rect.left(), y1=base_rect.top(),
                x2=base_rect.right(), y2=base_rect.bottom(),
                click=True, trigger=dict(DEFAULT_TRIGGER),
            )
            rec.actions.insert(0, base.to_dict())

        # 2) конвертируем старые abs area/word в rel относительно base
        base_rect = BaseAreaAction.from_dict(rec.actions[0]).rect()

        for i in range(1, len(rec.actions)):
            ad = rec.actions[i]
            if not isinstance(ad, dict):
                continue
            t = ad.get("type")

            if t == "area":
                a = AreaAction.from_dict(ad)
                if a.coord != "rel":
                    rg = a.rect_abs()
                    rec.actions[i] = AreaAction.from_global(
                        rg, base_rect, click=a.click, trigger=a.trigger, multiplier=a.multiplier, delay=a.delay
                    ).to_dict()

            elif t == "area_word":
                w = WordAreaAction.from_dict(ad)
                if w.coord != "rel":
                    rg = w.search_rect_abs()
                    rx1, ry1, rx2, ry2 = rect_to_rel(base_rect, rg)
                    w.coord = "rel"
                    w.rx1, w.ry1, w.rx2, w.ry2 = rx1, ry1, rx2, ry2
                    rec.actions[i] = w.to_dict()

    def _get_base_area(self, rec: Record) -> QRect:
        # 1) если привязан exe — используем реальную клиентскую область окна
        ctx = self._bound_context_for_record(rec=rec, notify_missing=False)
        effective_exe = str(ctx.get("effective_exe", "") or "").strip()
        if effective_exe:
            r = resolve_bound_base_rect_dip(effective_exe)
            if r and r.isValid():
                return r
        # 2) fallback: то что лежит в записи
        if rec.actions and isinstance(rec.actions[0], dict) and rec.actions[0].get("type") == "base_area":
            return BaseAreaAction.from_dict(rec.actions[0]).rect()

        return virtual_geometry()

    def _make_unique_record_name(self, base: str) -> str:
        base = (base or "").strip() or "Запись"
        existing = {r.name for r in self.records}

        if base not in existing:
            return base

        # сначала попробуем " (import)"
        candidate = f"{base} (import)"
        if candidate not in existing:
            return candidate

        # дальше " (import 2)", " (import 3)"...
        i = 2
        while True:
            candidate = f"{base} (import {i})"
            if candidate not in existing:
                return candidate
            i += 1

    def _make_export_name(self) -> str:
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return f"foxBot_{stamp}.json"

    def _write_json_to(self, path: Path, records_to_write: List[Record]):
        payload = {"records": [r.to_dict() for r in records_to_write]}
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def save_dialog(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Сохранение", "Выберите запись для сохранения.")
            return

        start_dir = getattr(self, "_last_io_dir", str(Path.home()))
        folder = QFileDialog.getExistingDirectory(self, "Куда сохранить файл", start_dir)
        if not folder:
            return

        folder_path = Path(folder)
        base = self._make_export_name()
        out = folder_path / base

        i = 1
        while out.exists():
            out = folder_path / base.replace(".json", f"_{i}.json")
            i += 1

        try:
            # сохраняем ТОЛЬКО текущую запись
            self._write_json_to(out, [rec])
            self._last_io_dir = str(folder_path)
            self._set_status(f"Сохранено (1 запись): {out}")
        except Exception as ex:
            QMessageBox.warning(self, "Ошибка сохранения", str(ex))

    def load_dialog(self):
        self.stop_playback()

        start_dir = getattr(self, "_last_io_dir", str(Path.home()))
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Загрузить файл",
            start_dir,
            "JSON (*.json);;Все файлы (*)"
        )
        if not filename:
            return

        p = Path(filename)

        try:
            payload = json.loads(p.read_text(encoding="utf-8"))

            loaded_records: List[Record] = []
            # поддержка текущего формата {"records":[...]}
            for rd in (payload.get("records", []) or []):
                loaded_records.append(Record.from_dict(rd))

            # на всякий случай: если вдруг файл содержит одну запись как {"record": {...}}
            if not loaded_records and isinstance(payload.get("record"), dict):
                loaded_records.append(Record.from_dict(payload["record"]))

        except Exception as ex:
            QMessageBox.warning(self, "Ошибка загрузки", f"Не удалось загрузить файл:\n{ex}")
            return

        if not loaded_records:
            QMessageBox.information(self, "Загрузка", "В файле не найдено записей.")
            return

        # добавляем как ДОПОЛНИТЕЛЬНЫЕ (не затираем существующие)e
        old_len = len(self.records)
        for r in loaded_records:
            self._ensure_base_and_migrate(r)  # <-- ВОТ ЭТО НУЖНО ДОБАВИТЬ
            r.name = self._make_unique_record_name(r.name)
            self.records.append(r)

        self._last_io_dir = str(p.parent)
        self._refresh_record_list_widget()

        # выделим первую импортированную
        self._set_current_record(old_len)


        # сохраняем объединённое состояние в автосейв
        self._save()

        self._set_status(f"Импортировано записей: {len(loaded_records)} из {p}")

    def pause_playback(self):
        if self.player and self.player.isRunning():
            self.player.pause("Пауза пользователем")

    def resume_playback(self):
        if self.player and self.player.isRunning():
            self.player.resume()

    def _set_pause_controls(self, playing: bool, paused: bool):
        self.btn_play.setEnabled(not playing)
        self.btn_stop.setEnabled(playing)
        self.btn_pause.setEnabled(playing and not paused)
        self.btn_resume.setEnabled(playing and paused)

    def _set_row_bg(self, row: int, color: Optional[QColor]):
        if row < 0 or row >= self.actions_table.rowCount():
            return
        for c in range(self.actions_table.columnCount()):
            it = self.actions_table.item(row, c)
            if not it:
                continue
            it.setBackground(QBrush(color) if color else QBrush())

    def _set_row_font_bold(self, row: int, bold: bool):
        if row < 0 or row >= self.actions_table.rowCount():
            return
        want = bool(bold)
        for c in range(self.actions_table.columnCount()):
            it = self.actions_table.item(row, c)
            if not it:
                continue
            f = it.font()
            if f.bold() == want:
                continue
            f.setBold(want)
            it.setFont(f)

    def _get_anchor_index(self, rec: Optional[Record]) -> int:
        if not rec:
            return -1
        idx = getattr(rec, "_anchor_index", None)
        if not isinstance(idx, int):
            idx = -1
        if not rec.actions:
            idx = -1
        elif idx < 0 or idx >= len(rec.actions):
            idx = 0
        rec._anchor_index = idx
        return idx

    def _set_anchor_index(self, rec: Optional[Record], idx: int):
        if not rec:
            return
        if not rec.actions:
            rec._anchor_index = -1
            return
        if idx < 0 or idx >= len(rec.actions):
            idx = 0
        rec._anchor_index = idx

    def _is_anchor_row(self, row: int) -> bool:
        rec = self._current_record()
        return row >= 0 and rec is not None and row == self._get_anchor_index(rec)

    def _action_row_label(self, row: int, rec: Optional[Record] = None) -> str:
        rec = rec or self._current_record()
        base = str(row + 1)
        if rec and row == self._get_anchor_index(rec):
            return f"A {base}"
        return base

    def _update_action_row_number(self, row: int):
        if row < 0 or row >= self.actions_table.rowCount():
            return
        text = self._action_row_label(row)
        it = self.actions_table.item(row, 0)
        if it is None:
            self.actions_table.setItem(row, 0, QTableWidgetItem(text))
        else:
            it.setText(text)

    def _apply_row_visual(self, row: int):
        if row < 0 or row >= self.actions_table.rowCount():
            return
        if row in self._error_rows:
            self._set_row_bg(row, QColor(180, 60, 60, 150))  # красный
        elif self._highlighted_row == row:
            self._set_row_bg(row, QColor(80, 110, 180, 120))  # синий
        elif self._is_anchor_row(row):
            self._set_row_bg(row, QColor(60, 100, 80, 120))
        else:
            self._set_row_bg(row, None)
        self._set_row_font_bold(row, self._params_row == row)

    def _set_params_row(self, row: Optional[int]):
        prev = self._params_row if isinstance(self._params_row, int) else None
        if isinstance(row, int) and 0 <= row < self.actions_table.rowCount():
            new_row: Optional[int] = row
        else:
            new_row = None
        if prev == new_row:
            return
        self._params_row = new_row
        if prev is not None:
            self._apply_row_visual(prev)
        if new_row is not None:
            self._apply_row_visual(new_row)

    def _clear_all_row_colors(self):
        for r in range(self.actions_table.rowCount()):
            self._apply_row_visual(r)

    def _on_player_paused(self, is_paused: bool, reason: str):
        self._is_paused = bool(is_paused)
        self._set_pause_controls(playing=True, paused=self._is_paused)
        if self._is_paused:
            self._set_status(f"Пауза: {reason}" if reason else "Пауза")

    def _on_player_action_error(self, row: int, msg: str):
        self._error_rows.add(row)
        self._apply_row_visual(row)
        self._set_status(msg, level="error")

    def _on_player_action_ok(self, row: int):
        if row in self._error_rows:
            self._error_rows.discard(row)
            self._apply_row_visual(row)

    def add_wait_action(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Сначала выберите или создайте запись.")
            return

        sec, ok = self._input_double_dark("Ожидание", "Сколько секунд ждать?", 1.0, 0.0, 9999.0, 3)
        if not ok:
            return

        wa = WaitAction(delay=Delay("fixed", float(sec), float(sec)))
        if self._is_fail_actions_focus_active():
            self._append_fail_action(wa)
            return

        rec.actions.append(wa.to_dict())
        new_row = len(rec.actions) - 1
        self._refresh_actions(select_row=new_row)
        self._save()

    def add_wait_event_action(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Сначала выберите или создайте запись.")
            return

        overlay = AreaSelectOverlay(initial_global=self._last_area_global(rec))
        rect = overlay.show_and_block()
        if not rect or not rect.isValid():
            return

        base = self._get_base_area(rec)
        text, ok = self._input_text_dark("Ожидание: Событие", "Текст для ожидания:", text="")
        if not ok:
            self._set_status("Отмена создания действия")
            return

        text = str(text or "")
        if not text.strip():
            self._set_status("Отмена создания действия")
            return

        new_a = WaitEventAction.from_global(
            rect,
            base,
            expected_text=text,
            ocr_lang=self._default_ocr_lang(),
            poll=1.0,
        )

        if self._is_fail_actions_focus_active():
            self._append_fail_action(new_a)
            return

        rec.actions.append(new_a.to_dict())
        new_row = len(rec.actions) - 1
        self._refresh_actions(select_row=new_row)
        self._save()

    @staticmethod
    def _delay_text(delay: Delay) -> str:
        if delay.mode == "range":
            return f"{delay.a:.3f}–{delay.b:.3f}с"
        return f"{delay.a:.3f}с"

    @staticmethod
    def _area_text_from_rel(rx1: float, ry1: float, rx2: float, ry2: float) -> str:
        try:
            w = abs(float(rx2) - float(rx1))
            h = abs(float(ry2) - float(ry1))
        except Exception:
            w, h = 0.0, 0.0
        return f"{(max(0.0, w) * max(0.0, h) * 100.0):.2f}%"

    @staticmethod
    def _area_text_from_rect(rect: QRect) -> str:
        area = max(0, int(rect.width())) * max(0, int(rect.height()))
        return f"{area} px²"

    def _action_area_text(self, a: Action) -> str:
        if isinstance(a, AreaAction):
            if a.coord == "rel":
                return self._area_text_from_rel(a.rx1, a.ry1, a.rx2, a.ry2)
            return self._area_text_from_rect(a.rect_abs())
        if isinstance(a, WordAreaAction):
            if a.coord == "rel":
                return self._area_text_from_rel(a.rx1, a.ry1, a.rx2, a.ry2)
            return self._area_text_from_rect(a.search_rect_abs())
        if isinstance(a, WaitEventAction):
            if a.coord == "rel":
                return self._area_text_from_rel(a.rx1, a.ry1, a.rx2, a.ry2)
            return self._area_text_from_rect(a.rect_abs())
        return ""

    def _action_type_text(self, a: Action) -> str:
        if isinstance(a, BaseAreaAction):
            return "Опорная область"
        if isinstance(a, WaitAction):
            return "Ожидание"
        if isinstance(a, WaitEventAction):
            return "Ожидание"
        if isinstance(a, WordAreaAction):
            return "Область"
        if isinstance(a, AreaAction):
            return "Область"
        if isinstance(a, KeyAction):
            return "Нажатие"
        return action_to_display(a)[0]

    def _action_desc_text(self, a: Action) -> str:
        if isinstance(a, WaitAction):
            return ""
        if isinstance(a, WaitEventAction):
            text = str(getattr(a, "expected_text", "") or "")
            return f"\"{text}\" ({self._action_area_text(a)})"

        if isinstance(a, AreaAction):
            area_txt = self._action_area_text(a)
            if bool(getattr(a, "click", False)):
                trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
                return f"{spec_to_pretty(trig)} · {area_txt}"
            return area_txt

        if isinstance(a, WordAreaAction):
            word = str(getattr(a, "word", "") or "")
            try:
                idx = max(1, int(getattr(a, "index", 1)))
            except Exception:
                idx = 1
            try:
                count = int(getattr(a, "count", 0))
            except Exception:
                count = 0
            idx_txt = f"{idx}/{count}" if count > 0 else str(idx)
            return f"\"{word}\" {idx_txt} ({self._action_area_text(a)})"

        if isinstance(a, KeyAction):
            mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
            if mode == "long":
                seq: List[str] = []
                for item in self._key_long_actions_from_key_action(a):
                    trig = normalize_trigger(item.get("trigger", DEFAULT_TRIGGER))
                    seq.append(spec_to_pretty(trig))
                if seq:
                    return " -> ".join(seq)
                return ""

            trig = normalize_trigger(
                {
                    "kind": getattr(a, "kind", "keys"),
                    "keys": list(getattr(a, "keys", []) or []),
                    "mouse_button": getattr(a, "mouse_button", None),
                }
            )
            return spec_to_pretty(trig)

        return action_to_display(a)[1]

    def _action_params_text(self, a: Action) -> str:
        if isinstance(a, KeyAction):
            return self._key_action_params_text(a)
        if isinstance(a, WaitAction):
            return self._delay_text(a.delay)
        if isinstance(a, WaitEventAction):
            return f"{max(0.1, a.poll):.2f}"
        if isinstance(a, AreaAction):
            parts: List[str] = [f"x{max(1, int(getattr(a, 'multiplier', 1)))}"]
            delay = getattr(a, "delay", None)
            if isinstance(delay, Delay):
                parts.append(self._delay_text(delay))
            return "; ".join(parts)
        if isinstance(a, WordAreaAction):
            parts: List[str] = []
            mult = max(1, int(getattr(a, "multiplier", 1)))
            parts.append(f"x{mult}")
            delay = getattr(a, "delay", None)
            if isinstance(delay, Delay):
                parts.append(self._delay_text(delay))
            if bool(getattr(a, "click", False)):
                trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
                parts.append(spec_to_pretty(trig))
            return "; ".join(parts)
        return ""

    def _update_actions_table_row(self, row: int, a: Action):
        if row < 0 or row >= self.actions_table.rowCount():
            return

        self._update_action_row_number(row)

        t = self._action_type_text(a)
        desc = self._action_desc_text(a)
        params = self._action_params_text(a)

        for col, text in [(1, t), (2, desc), (3, params)]:
            it = self.actions_table.item(row, col)
            if it is None:
                self.actions_table.setItem(row, col, QTableWidgetItem(text))
            else:
                it.setText(text)

    def _get_selected_area_action(self) -> Optional[tuple[int, Union[AreaAction, WordAreaAction]]]:
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return None
        a = action_from_dict(rec.actions[row])
        if not isinstance(a, (AreaAction, WordAreaAction)):
            return None
        return (row, a)

    def _sync_area_panel(self, a: Union[AreaAction, WordAreaAction]):
        trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))

        self.cb_area_click.blockSignals(True)
        self.cb_area_click.setChecked(bool(getattr(a, "click", False)))
        self.cb_area_click.blockSignals(False)

        if hasattr(self, "btn_area_pick"):
            self.btn_area_pick.setText(spec_to_pretty(trig))

        if hasattr(self, "sp_area_multiplier"):
            self.sp_area_multiplier.blockSignals(True)
            self.sp_area_multiplier.setValue(max(1, int(getattr(a, "multiplier", 1))))
            self.sp_area_multiplier.blockSignals(False)

        if hasattr(self, "cb_area_random_delay"):
            self.cb_area_random_delay.blockSignals(True)
            self.cb_area_random_delay.setChecked(getattr(a, "delay", Delay()).mode == "range")
            self.cb_area_random_delay.blockSignals(False)

        if hasattr(self, "sp_area_delay_a"):
            self.sp_area_delay_a.blockSignals(True)
            self.sp_area_delay_a.setValue(float(getattr(a, "delay", Delay()).a))
            self.sp_area_delay_a.blockSignals(False)

        if hasattr(self, "sp_area_delay_b"):
            self.sp_area_delay_b.blockSignals(True)
            self.sp_area_delay_b.setValue(float(getattr(a, "delay", Delay()).b))
            self.sp_area_delay_b.blockSignals(False)

        if hasattr(self, "_set_area_delay_b_row_visible"):
            self._set_area_delay_b_row_visible(getattr(a, "delay", Delay()).mode == "range")

        is_word = isinstance(a, WordAreaAction)

        # слово/номер доступны только для "по слову"
        if hasattr(self, "le_area_word"):
            self.le_area_word.blockSignals(True)
            self.le_area_word.setEnabled(is_word)
            self.le_area_word.setText(a.word if is_word else "")
            self.le_area_word.blockSignals(False)

        if hasattr(self, "area_ocr_lang_group"):
            self._set_ocr_lang_selector(
                self.area_ocr_lang_group,
                getattr(a, "ocr_lang", self._default_ocr_lang()) if is_word else self._default_ocr_lang(),
            )

        search_infinite = bool(getattr(a, "search_infinite", True)) if is_word else True
        try:
            search_max_tries = max(1, int(getattr(a, "search_max_tries", 100))) if is_word else 100
        except Exception:
            search_max_tries = 100
        search_on_fail = str(getattr(a, "search_on_fail", "retry")) if is_word else "retry"
        if search_on_fail not in ("retry", "error", "action"):
            search_on_fail = "retry"
        fail_post_mode = str(getattr(a, "on_fail_post_mode", "none")) if is_word else "none"
        if fail_post_mode not in ("none", "stop", "repeat"):
            fail_post_mode = "none"

        if hasattr(self, "cb_area_search_infinite"):
            self.cb_area_search_infinite.blockSignals(True)
            self.cb_area_search_infinite.setEnabled(is_word)
            self.cb_area_search_infinite.setChecked(search_infinite)
            self.cb_area_search_infinite.blockSignals(False)
        if hasattr(self, "sp_area_search_max_tries"):
            self.sp_area_search_max_tries.blockSignals(True)
            self.sp_area_search_max_tries.setValue(search_max_tries)
            self.sp_area_search_max_tries.blockSignals(False)
        if hasattr(self, "area_search_on_fail_row"):
            self.area_search_on_fail_row.setEnabled(is_word)
        if hasattr(self, "area_search_on_fail_group"):
            self._set_single_choice_selector(self.area_search_on_fail_group, search_on_fail, default_code="retry")
        self._set_area_search_max_tries_enabled(is_word and (not search_infinite))

        if hasattr(self, "sp_area_index"):
            self.sp_area_index.blockSignals(True)
            self.sp_area_index.setEnabled(is_word)
            self.sp_area_index.setValue(max(1, int(getattr(a, "index", 1))) if is_word else 1)
            self.sp_area_index.blockSignals(False)

        if hasattr(self, "sp_area_count"):
            self.sp_area_count.blockSignals(True)
            self.sp_area_count.setEnabled(is_word)
            try:
                count_val = int(getattr(a, "count", 1)) if is_word else 1
            except Exception:
                count_val = 1
            if count_val < 1:
                count_val = 1
            self.sp_area_count.setValue(count_val)
            self.sp_area_count.blockSignals(False)

        if hasattr(self, "lbl_area_index"):
            self.lbl_area_index.setEnabled(is_word)

        if hasattr(self, "lbl_area_count"):
            self.lbl_area_count.setEnabled(is_word)

        if hasattr(self, "lbl_area_word"):
            self.lbl_area_word.setEnabled(is_word)

        if hasattr(self, "lbl_area_ocr_lang"):
            self.lbl_area_ocr_lang.setEnabled(is_word)
        if hasattr(self, "lbl_area_search_opts"):
            self.lbl_area_search_opts.setEnabled(is_word)
        if hasattr(self, "lbl_area_search_on_fail"):
            self.lbl_area_search_on_fail.setEnabled(is_word)

        if hasattr(self, "lbl_area_index"):
            self.lbl_area_index.setVisible(is_word)
        if hasattr(self, "area_word_row"):
            self.area_word_row.setVisible(is_word)
        if hasattr(self, "lbl_area_ocr_lang"):
            self.lbl_area_ocr_lang.setVisible(is_word)
        if hasattr(self, "area_ocr_lang_row"):
            self.area_ocr_lang_row.setVisible(is_word)
        if hasattr(self, "lbl_area_search_opts"):
            self.lbl_area_search_opts.setVisible(is_word)
        if hasattr(self, "area_search_opts_row"):
            self.area_search_opts_row.setVisible(is_word)
        if hasattr(self, "lbl_area_search_on_fail"):
            self.lbl_area_search_on_fail.setVisible(is_word)
        if hasattr(self, "area_search_on_fail_row"):
            self.area_search_on_fail_row.setVisible(is_word)

        fail_actions_visible = bool(is_word and search_on_fail == "action")
        if hasattr(self, "fail_actions_box"):
            self.fail_actions_box.setVisible(fail_actions_visible)
            self.fail_actions_box.setEnabled(is_word)
        if hasattr(self, "cb_focus_fail_actions"):
            self.cb_focus_fail_actions.setEnabled(fail_actions_visible)
            if not fail_actions_visible:
                self._set_fail_actions_focus_checked(False)
        if hasattr(self, "_set_fail_actions_post_mode"):
            self._set_fail_actions_post_mode(fail_post_mode)
        if hasattr(self, "cb_fail_actions_stop"):
            self.cb_fail_actions_stop.setEnabled(fail_actions_visible)
        if hasattr(self, "cb_fail_actions_repeat"):
            self.cb_fail_actions_repeat.setEnabled(fail_actions_visible)

        if fail_actions_visible:
            self._refresh_fail_actions()
        else:
            if hasattr(self, "fail_actions_table"):
                self.fail_actions_table.blockSignals(True)
                self.fail_actions_table.setRowCount(0)
                self.fail_actions_table.blockSignals(False)

    def _apply_area_params(self):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        sel = self._get_selected_area_action()
        if not rec or not sel:
            return
        row, a = sel

        a.click = bool(self.cb_area_click.isChecked())
        a.trigger = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
        if hasattr(self, "sp_area_multiplier"):
            a.multiplier = max(1, int(self.sp_area_multiplier.value()))
        if hasattr(self, "cb_area_random_delay") and hasattr(self, "sp_area_delay_a") and hasattr(self, "sp_area_delay_b"):
            if self.cb_area_random_delay.isChecked():
                a.delay.mode = "range"
                a.delay.a = float(self.sp_area_delay_a.value())
                a.delay.b = float(self.sp_area_delay_b.value())
            else:
                a.delay.mode = "fixed"
                a.delay.a = float(self.sp_area_delay_a.value())
                a.delay.b = float(self.sp_area_delay_a.value())
            self._set_area_delay_b_row_visible(self.cb_area_random_delay.isChecked())

        if isinstance(a, WordAreaAction):
            a.word = str(self.le_area_word.text() or "")
            a.index = max(1, int(self.sp_area_index.value()))
            if hasattr(self, "sp_area_count"):
                a.count = max(1, int(self.sp_area_count.value()))
            if hasattr(self, "area_ocr_lang_group"):
                a.ocr_lang = self._get_ocr_lang_selector(self.area_ocr_lang_group)
            if hasattr(self, "cb_area_search_infinite"):
                a.search_infinite = bool(self.cb_area_search_infinite.isChecked())
            else:
                a.search_infinite = True
            if hasattr(self, "sp_area_search_max_tries"):
                a.search_max_tries = max(1, int(self.sp_area_search_max_tries.value()))
            else:
                a.search_max_tries = 100
            if hasattr(self, "area_search_on_fail_group"):
                a.search_on_fail = self._get_single_choice_selector(self.area_search_on_fail_group, default_code="retry")
            else:
                a.search_on_fail = "retry"
            if hasattr(self, "_get_fail_actions_post_mode"):
                a.on_fail_post_mode = self._get_fail_actions_post_mode()
            else:
                a.on_fail_post_mode = "none"
            self._set_area_search_max_tries_enabled(not bool(a.search_infinite))

        if isinstance(a, WordAreaAction) and a.trigger["kind"] == "mouse" and a.trigger["mouse_button"]:
            a.button = a.trigger["mouse_button"]

        rec.actions[row] = a.to_dict()
        self._update_actions_table_row(row, a)  # обновляем только одну строку, без пересоздания таблицы
        if isinstance(a, WordAreaAction):
            self._sync_area_panel(a)
        self._save()

    def _pick_area_trigger(self):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        sel = self._get_selected_area_action()
        if not rec or not sel:
            return
        row, a = sel

        # покажем дырку по текущей зоне
        base = self._get_base_area(rec)
        if isinstance(a, AreaAction):
            area_global = a.rect_global(base)
        else:
            area_global = a.search_rect_global(base)

        # initial для оверлея
        trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
        initial = KeyAction(kind=trig["kind"], keys=list(trig["keys"]), mouse_button=trig["mouse_button"])

        overlay = KeyCaptureOverlay(area_global=area_global, initial=initial)
        spec = overlay.show_and_block()
        if spec is None:
            return

        a.trigger = normalize_trigger(spec)
        # удобно: если ты выбрал кнопку/комбо — логично включить boolean
        a.click = True

        if isinstance(a, WordAreaAction) and a.trigger["kind"] == "mouse" and a.trigger["mouse_button"]:
            a.button = a.trigger["mouse_button"]

        rec.actions[row] = a.to_dict()
        self._refresh_actions()
        self.actions_table.selectRow(row)
        self._sync_area_panel(a)
        self._save()

    def _reset_area_trigger(self):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        sel = self._get_selected_area_action()
        if not rec or not sel:
            return
        row, a = sel
        a.trigger = dict(DEFAULT_TRIGGER)
        if isinstance(a, WordAreaAction):
            a.button = "left"
        rec.actions[row] = a.to_dict()
        self._refresh_actions()
        self.actions_table.selectRow(row)
        self._sync_area_panel(a)
        self._save()

    def _set_area_params_enabled(self, en: bool):
        self.area_params.setEnabled(en)

    def _get_selected_wait_event_action(self) -> Optional[tuple[int, WaitEventAction]]:
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return None
        a = action_from_dict(rec.actions[row])
        if not isinstance(a, WaitEventAction):
            return None
        return (row, a)

    def _sync_wait_event_panel(self, a: WaitEventAction):
        self.le_wait_event_text.blockSignals(True)
        self.le_wait_event_text.setText(str(getattr(a, "expected_text", "")))
        self.le_wait_event_text.blockSignals(False)

        self.sp_wait_event_poll.blockSignals(True)
        self.sp_wait_event_poll.setValue(max(0.1, float(getattr(a, "poll", 1.0))))
        self.sp_wait_event_poll.blockSignals(False)

        if hasattr(self, "wait_event_ocr_lang_group"):
            self._set_ocr_lang_selector(
                self.wait_event_ocr_lang_group,
                getattr(a, "ocr_lang", self._default_ocr_lang()),
            )

    def _apply_wait_event_params(self):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        sel = self._get_selected_wait_event_action()
        if not rec or not sel:
            return
        row, a = sel

        a.expected_text = str(self.le_wait_event_text.text() or "")
        a.poll = max(0.1, float(self.sp_wait_event_poll.value()))
        if hasattr(self, "wait_event_ocr_lang_group"):
            a.ocr_lang = self._get_ocr_lang_selector(self.wait_event_ocr_lang_group)
        else:
            a.ocr_lang = self._default_ocr_lang()

        rec.actions[row] = a.to_dict()
        self._update_actions_table_row(row, a)
        self._save()

    def _set_wait_event_params_enabled(self, en: bool):
        self.wait_event_params.setEnabled(en)

    def add_area_action_by_word(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Сначала выберите или создайте запись.")
            return

        overlay = AreaSelectOverlay(initial_global=self._last_area_global(rec))
        rect = overlay.show_and_block()
        if not rect or not rect.isValid():
            return

        word, ok = self._input_text_dark("Область: Текст", "Какой текст искать в выделенной зоне?")
        if not ok:
            self._set_status("Отмена создания действия")
            return
        word = str(word or "")
        if not word.strip():
            self._set_status("Отмена создания действия")
            return

        idx, ok = self._input_int_dark(
            "Номер совпадения",
            "Какое по счёту совпадение брать? (1 = первое)",
            1,
            1,
            9999,
            1,
        )
        if not ok:
            self._set_status("Отмена создания действия")
            return

        base = self._get_base_area(rec)
        rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)

        aa = WordAreaAction(
            word=word, index=idx,
            click=True,
            button="left",
            ocr_lang=self._default_ocr_lang(),
        )
        aa.coord = "rel"
        aa.rx1, aa.ry1, aa.rx2, aa.ry2 = rx1, ry1, rx2, ry2
        aa.trigger = dict(DEFAULT_TRIGGER)

        if self._is_fail_actions_focus_active():
            self._append_fail_action(aa)
            return

        rec.actions.append(aa.to_dict())

        new_row = len(rec.actions) - 1
        self._refresh_actions(select_row=new_row)
        self._save()

    def _selected_action_rows(self) -> List[int]:
        sm = self.actions_table.selectionModel()
        if not sm:
            return []
        rows = sorted({idx.row() for idx in sm.selectedRows()})
        return rows

    def _selected_fail_action_rows(self) -> List[int]:
        if not hasattr(self, "fail_actions_table"):
            return []
        sm = self.fail_actions_table.selectionModel()
        if not sm:
            return []
        return sorted({idx.row() for idx in sm.selectedRows()})

    def _selected_fail_action_row(self) -> Optional[int]:
        rows = self._selected_fail_action_rows()
        return rows[0] if rows else None

    def _get_fail_actions_owner(self) -> Optional[tuple[Record, int, WordAreaAction]]:
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return None
        if row < 0 or row >= len(rec.actions):
            return None
        try:
            a = action_from_dict(rec.actions[row])
        except Exception:
            return None
        if not isinstance(a, WordAreaAction):
            return None
        return rec, row, a

    def _is_fail_actions_mode_available(self) -> bool:
        owner = self._get_fail_actions_owner()
        if not owner:
            return False
        _rec, _row, a = owner
        return str(getattr(a, "search_on_fail", "retry") or "retry") == "action"

    def _is_fail_actions_focus_active(self) -> bool:
        if not hasattr(self, "cb_focus_fail_actions"):
            return False
        return bool(self.cb_focus_fail_actions.isChecked() and self._is_fail_actions_mode_available())

    def _set_fail_actions_focus_checked(self, checked: bool):
        if not hasattr(self, "cb_focus_fail_actions"):
            return
        self.cb_focus_fail_actions.blockSignals(True)
        self.cb_focus_fail_actions.setChecked(bool(checked))
        self.cb_focus_fail_actions.blockSignals(False)

    def _set_fail_actions_post_mode(self, mode: str):
        norm = str(mode or "none")
        if norm not in ("none", "stop", "repeat"):
            norm = "none"
        stop_checked = norm == "stop"
        repeat_checked = norm == "repeat"
        if hasattr(self, "cb_fail_actions_stop"):
            self.cb_fail_actions_stop.blockSignals(True)
            self.cb_fail_actions_stop.setChecked(stop_checked)
            self.cb_fail_actions_stop.blockSignals(False)
        if hasattr(self, "cb_fail_actions_repeat"):
            self.cb_fail_actions_repeat.blockSignals(True)
            self.cb_fail_actions_repeat.setChecked(repeat_checked)
            self.cb_fail_actions_repeat.blockSignals(False)

    def _get_fail_actions_post_mode(self) -> str:
        stop_checked = bool(getattr(self, "cb_fail_actions_stop", None) and self.cb_fail_actions_stop.isChecked())
        repeat_checked = bool(getattr(self, "cb_fail_actions_repeat", None) and self.cb_fail_actions_repeat.isChecked())
        if stop_checked and not repeat_checked:
            return "stop"
        if repeat_checked and not stop_checked:
            return "repeat"
        return "none"

    def _on_fail_actions_stop_toggled(self, checked: bool):
        if bool(checked) and hasattr(self, "cb_fail_actions_repeat") and self.cb_fail_actions_repeat.isChecked():
            self.cb_fail_actions_repeat.blockSignals(True)
            self.cb_fail_actions_repeat.setChecked(False)
            self.cb_fail_actions_repeat.blockSignals(False)
        self._apply_area_params()

    def _on_fail_actions_repeat_toggled(self, checked: bool):
        if bool(checked) and hasattr(self, "cb_fail_actions_stop") and self.cb_fail_actions_stop.isChecked():
            self.cb_fail_actions_stop.blockSignals(True)
            self.cb_fail_actions_stop.setChecked(False)
            self.cb_fail_actions_stop.blockSignals(False)
        self._apply_area_params()

    def _on_fail_actions_focus_toggled(self, _checked: bool):
        if not self._is_fail_actions_mode_available():
            self._set_fail_actions_focus_checked(False)
        self._on_action_selected()

    def _save_fail_actions_owner(self, rec: Record, owner_row: int, owner_action: WordAreaAction):
        rec.actions[owner_row] = owner_action.to_dict()
        self._update_actions_table_row(owner_row, owner_action)
        self._save()

    def _refresh_fail_actions(self, select_row: Optional[int] = None):
        if not hasattr(self, "fail_actions_table"):
            return

        owner = self._get_fail_actions_owner()
        if not owner:
            self.fail_actions_table.blockSignals(True)
            self.fail_actions_table.setRowCount(0)
            self.fail_actions_table.blockSignals(False)
            self._active_fail_actions_owner_row = None
            return

        _rec, owner_row, owner_action = owner
        if str(getattr(owner_action, "search_on_fail", "retry") or "retry") != "action":
            self.fail_actions_table.blockSignals(True)
            self.fail_actions_table.setRowCount(0)
            self.fail_actions_table.blockSignals(False)
            self._active_fail_actions_owner_row = owner_row
            return

        actions = list(getattr(owner_action, "on_fail_actions", []) or [])
        prev_row = self._selected_fail_action_row()

        self.fail_actions_table.blockSignals(True)
        self.fail_actions_table.setRowCount(0)

        for i, ad in enumerate(actions):
            try:
                a = action_from_dict(ad)
            except Exception:
                continue
            t = self._action_type_text(a)
            desc = self._action_desc_text(a)
            params = self._action_params_text(a)

            self.fail_actions_table.insertRow(i)
            self.fail_actions_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.fail_actions_table.setItem(i, 1, QTableWidgetItem(t))
            self.fail_actions_table.setItem(i, 2, QTableWidgetItem(desc))
            self.fail_actions_table.setItem(i, 3, QTableWidgetItem(params))

        self.fail_actions_table.blockSignals(False)

        row = select_row if select_row is not None else prev_row
        if row is None:
            row = 0 if self.fail_actions_table.rowCount() > 0 else -1
        if 0 <= row < self.fail_actions_table.rowCount():
            self.fail_actions_table.selectRow(row)
            it = self.fail_actions_table.item(row, 0)
            if it:
                self.fail_actions_table.scrollToItem(it, QAbstractItemView.PositionAtCenter)
        else:
            self.fail_actions_table.clearSelection()

        self._active_fail_actions_owner_row = owner_row

    def _append_fail_action(self, action_obj: Action):
        owner = self._get_fail_actions_owner()
        if not owner:
            return
        rec, owner_row, owner_action = owner
        owner_action.on_fail_actions.append(action_obj.to_dict())
        self._save_fail_actions_owner(rec, owner_row, owner_action)
        self._refresh_fail_actions(select_row=len(owner_action.on_fail_actions) - 1)

    def _normalize_long_delay_cfg(self, raw: Any, *, default_sec: float) -> Dict[str, Any]:
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

    def _delay_cfg_to_text(self, cfg: Dict[str, Any]) -> str:
        mode = "range" if str(cfg.get("mode", "fixed")).lower() == "range" else "fixed"
        try:
            a = max(0.0, float(cfg.get("a", 0.0)))
        except Exception:
            a = 0.0
        try:
            b = max(0.0, float(cfg.get("b", a)))
        except Exception:
            b = a
        if mode == "range":
            return f"{a:.3f}–{b:.3f}с"
        return f"{a:.3f}с"

    def _default_key_long_action_item(self, trigger: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        return {
            "trigger": normalize_trigger(trigger or DEFAULT_TRIGGER),
            "hold": self._normalize_long_delay_cfg(None, default_sec=0.2),
            "activate_mode": "after_prev",
            "start_delay": self._normalize_long_delay_cfg(None, default_sec=0.0),
        }

    def _normalize_key_long_action_item(self, raw: Any) -> Optional[Dict[str, Any]]:
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

        item = self._default_key_long_action_item(trigger_raw)
        item["hold"] = self._normalize_long_delay_cfg(hold_raw, default_sec=0.2)
        item["activate_mode"] = "from_start" if str(activate_mode_raw).strip().lower() == "from_start" else "after_prev"
        item["start_delay"] = self._normalize_long_delay_cfg(start_delay_raw, default_sec=0.0)
        return item

    def _key_long_actions_from_key_action(self, a: KeyAction) -> List[Dict[str, Any]]:
        raw_cfg = getattr(a, "long_press", {})
        if not isinstance(raw_cfg, dict):
            return []
        raw_actions = raw_cfg.get("actions", [])
        if not isinstance(raw_actions, list):
            return []

        out: List[Dict[str, Any]] = []
        for raw in raw_actions:
            item = self._normalize_key_long_action_item(raw)
            if item is not None:
                out.append(item)
        return out

    def _set_key_long_actions_for_key_action(self, a: KeyAction, actions: List[Dict[str, Any]]):
        cfg = copy.deepcopy(getattr(a, "long_press", {})) if isinstance(getattr(a, "long_press", {}), dict) else {}
        out: List[Dict[str, Any]] = []
        for raw in (actions or []):
            item = self._normalize_key_long_action_item(raw)
            if item is not None:
                out.append(item)
        cfg["actions"] = out
        a.long_press = cfg

    def _selected_key_long_action_rows(self) -> List[int]:
        if not hasattr(self, "key_long_actions_table"):
            return []
        sm = self.key_long_actions_table.selectionModel()
        if not sm:
            return []
        return sorted({idx.row() for idx in sm.selectedRows()})

    def _selected_key_long_action_row(self) -> Optional[int]:
        rows = self._selected_key_long_action_rows()
        return rows[0] if rows else None

    def _set_key_long_hold_b_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "lbl_key_long_hold_b"):
            self.lbl_key_long_hold_b.setVisible(v)
        if hasattr(self, "sp_key_long_hold_b"):
            self.sp_key_long_hold_b.setVisible(v)
            self.sp_key_long_hold_b.setEnabled(v)

    def _set_key_long_start_b_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "lbl_key_long_start_b"):
            self.lbl_key_long_start_b.setVisible(v)
        if hasattr(self, "sp_key_long_start_b"):
            self.sp_key_long_start_b.setVisible(v)
            self.sp_key_long_start_b.setEnabled(v)

    def _set_key_long_start_delay_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "lbl_key_long_start_delay"):
            self.lbl_key_long_start_delay.setVisible(v)
        if hasattr(self, "key_long_start_delay_row"):
            self.key_long_start_delay_row.setVisible(v)
        self._set_key_long_start_b_visible(v and hasattr(self, "cb_key_long_start_range") and self.cb_key_long_start_range.isChecked())

    def _set_key_long_params_enabled(self, enabled: bool):
        en = bool(enabled)
        if hasattr(self, "cb_key_long_hold_range"):
            self.cb_key_long_hold_range.setEnabled(en)
        if hasattr(self, "sp_key_long_hold_a"):
            self.sp_key_long_hold_a.setEnabled(en)
        if hasattr(self, "sp_key_long_hold_b"):
            self.sp_key_long_hold_b.setEnabled(en and self.cb_key_long_hold_range.isChecked())
        if hasattr(self, "key_long_activation_slider"):
            self.key_long_activation_slider.setEnabled(en)
        if hasattr(self, "cb_key_long_start_range"):
            self.cb_key_long_start_range.setEnabled(en)
        if hasattr(self, "sp_key_long_start_a"):
            self.sp_key_long_start_a.setEnabled(en)
        if hasattr(self, "sp_key_long_start_b"):
            can_start_b = (
                en
                and hasattr(self, "key_long_activation_slider")
                and self.key_long_activation_slider.mode() == "from_start"
                and hasattr(self, "cb_key_long_start_range")
                and self.cb_key_long_start_range.isChecked()
            )
            self.sp_key_long_start_b.setEnabled(can_start_b)

    def _clear_key_long_params_panel(self):
        if not hasattr(self, "cb_key_long_hold_range"):
            return
        self._key_long_params_loading = True
        if hasattr(self, "cb_key_long_hold_range"):
            self.cb_key_long_hold_range.setChecked(False)
        if hasattr(self, "sp_key_long_hold_a"):
            self.sp_key_long_hold_a.setValue(0.2)
        if hasattr(self, "sp_key_long_hold_b"):
            self.sp_key_long_hold_b.setValue(0.2)
        if hasattr(self, "key_long_activation_slider"):
            self.key_long_activation_slider.set_mode("after_prev", animate=False, emit_signal=False)
        if hasattr(self, "cb_key_long_start_range"):
            self.cb_key_long_start_range.setChecked(False)
        if hasattr(self, "sp_key_long_start_a"):
            self.sp_key_long_start_a.setValue(0.0)
        if hasattr(self, "sp_key_long_start_b"):
            self.sp_key_long_start_b.setValue(0.0)
        self._key_long_params_loading = False
        self._set_key_long_hold_b_visible(False)
        self._set_key_long_start_delay_visible(False)

    def _sync_selected_key_long_action_params(self):
        owner = self._get_key_long_actions_owner()
        row = self._selected_key_long_action_row()
        if not owner or row is None:
            self._clear_key_long_params_panel()
            return

        _rec, _owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        if row < 0 or row >= len(actions):
            self._clear_key_long_params_panel()
            return

        item = actions[row]
        hold_cfg = self._normalize_long_delay_cfg(item.get("hold"), default_sec=0.2)
        activate_mode = "from_start" if str(item.get("activate_mode", "")).lower() == "from_start" else "after_prev"
        start_cfg = self._normalize_long_delay_cfg(item.get("start_delay"), default_sec=0.0)

        self._key_long_params_loading = True
        self.cb_key_long_hold_range.setChecked(hold_cfg.get("mode") == "range")
        self.sp_key_long_hold_a.setValue(float(hold_cfg.get("a", 0.2)))
        self.sp_key_long_hold_b.setValue(float(hold_cfg.get("b", hold_cfg.get("a", 0.2))))
        self.key_long_activation_slider.set_mode(activate_mode, animate=False, emit_signal=False)
        self.cb_key_long_start_range.setChecked(start_cfg.get("mode") == "range")
        self.sp_key_long_start_a.setValue(float(start_cfg.get("a", 0.0)))
        self.sp_key_long_start_b.setValue(float(start_cfg.get("b", start_cfg.get("a", 0.0))))
        self._key_long_params_loading = False

        self._set_key_long_hold_b_visible(self.cb_key_long_hold_range.isChecked())
        self._set_key_long_start_delay_visible(activate_mode == "from_start")

    def _sync_key_long_actions_buttons_state(self):
        owner = self._get_key_long_actions_owner()
        can_work = bool(owner) and (not getattr(self, "_playing", False))
        selected_row = self._selected_key_long_action_row()
        has_selection = selected_row is not None
        total_rows = self.key_long_actions_table.rowCount() if hasattr(self, "key_long_actions_table") else 0
        can_move_up = bool(can_work and has_selection and selected_row is not None and selected_row > 0)
        can_move_down = bool(
            can_work and has_selection and selected_row is not None and selected_row < (total_rows - 1)
        )

        if hasattr(self, "btn_key_long_actions_add"):
            self.btn_key_long_actions_add.setEnabled(can_move_up)
        if hasattr(self, "btn_key_long_actions_del"):
            self.btn_key_long_actions_del.setEnabled(can_move_down)

        if hasattr(self, "btn_key_long_bottom_add_key"):
            self.btn_key_long_bottom_add_key.setEnabled(can_work)
        if hasattr(self, "btn_key_long_bottom_edit"):
            self.btn_key_long_bottom_edit.setEnabled(can_work and has_selection)
        if hasattr(self, "btn_key_long_bottom_delete"):
            self.btn_key_long_bottom_delete.setEnabled(can_work and has_selection)

        self._set_key_long_params_enabled(can_work and has_selection)
        if can_work and has_selection:
            self._sync_selected_key_long_action_params()
        else:
            self._clear_key_long_params_panel()

    def _get_key_long_actions_owner(self) -> Optional[tuple[Record, int, KeyAction]]:
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return None
        if row < 0 or row >= len(rec.actions):
            return None
        try:
            a = action_from_dict(rec.actions[row])
        except Exception:
            return None
        if not isinstance(a, KeyAction):
            return None
        mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
        if mode != "long":
            return None
        return rec, row, a

    def _save_key_long_actions_owner(self, rec: Record, owner_row: int, owner_action: KeyAction):
        rec.actions[owner_row] = owner_action.to_dict()
        self._update_actions_table_row(owner_row, owner_action)
        self._save()

    def _commit_key_long_actions(
        self,
        rec: Record,
        owner_row: int,
        owner_action: KeyAction,
        actions: List[Dict[str, Any]],
        *,
        select_row: Optional[int] = None,
    ):
        self._set_key_long_actions_for_key_action(owner_action, actions)
        self._save_key_long_actions_owner(rec, owner_row, owner_action)
        self._refresh_key_long_actions(select_row=select_row)

    def _refresh_key_long_actions(self, select_row: Optional[int] = None):
        if not hasattr(self, "key_long_actions_table"):
            return

        owner = self._get_key_long_actions_owner()

        if not owner:
            self.key_long_actions_table.blockSignals(True)
            self.key_long_actions_table.setRowCount(0)
            self.key_long_actions_table.blockSignals(False)
            self._active_key_long_actions_owner_row = None
            self._sync_key_long_actions_buttons_state()
            return

        _rec, owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        prev_row = self._selected_key_long_action_row()

        self.key_long_actions_table.blockSignals(True)
        self.key_long_actions_table.setRowCount(0)

        for i, item in enumerate(actions):
            trigger = normalize_trigger(item.get("trigger", DEFAULT_TRIGGER))
            hold_cfg = self._normalize_long_delay_cfg(item.get("hold"), default_sec=0.2)
            start_cfg = self._normalize_long_delay_cfg(item.get("start_delay"), default_sec=0.0)
            activate_mode = "from_start" if str(item.get("activate_mode", "")).lower() == "from_start" else "after_prev"

            hold_txt = self._delay_cfg_to_text(hold_cfg)
            if activate_mode == "from_start":
                activate_txt = f"От старта: {self._delay_cfg_to_text(start_cfg)}"
            else:
                activate_txt = "После предыдущего"

            self.key_long_actions_table.insertRow(i)
            self.key_long_actions_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.key_long_actions_table.setItem(i, 1, QTableWidgetItem(spec_to_pretty(trigger)))
            self.key_long_actions_table.setItem(i, 2, QTableWidgetItem(hold_txt))
            self.key_long_actions_table.setItem(i, 3, QTableWidgetItem(activate_txt))

        self.key_long_actions_table.blockSignals(False)

        row = select_row if select_row is not None else prev_row
        if row is None:
            row = 0 if self.key_long_actions_table.rowCount() > 0 else -1
        if 0 <= row < self.key_long_actions_table.rowCount():
            self.key_long_actions_table.selectRow(row)
            it = self.key_long_actions_table.item(row, 0)
            if it:
                self.key_long_actions_table.scrollToItem(it, QAbstractItemView.PositionAtCenter)
        else:
            self.key_long_actions_table.clearSelection()

        self._active_key_long_actions_owner_row = owner_row
        self._sync_key_long_actions_buttons_state()

    def _append_key_long_action(self, action_item: Dict[str, Any]):
        owner = self._get_key_long_actions_owner()
        if not owner:
            return
        rec, owner_row, owner_action = owner
        normalized = self._normalize_key_long_action_item(action_item)
        if normalized is None:
            return
        actions = self._key_long_actions_from_key_action(owner_action)
        actions.append(normalized)
        self._commit_key_long_actions(rec, owner_row, owner_action, actions, select_row=len(actions) - 1)

    def _apply_selected_key_long_action_params(self):
        if getattr(self, "_key_long_params_loading", False):
            return

        owner = self._get_key_long_actions_owner()
        row = self._selected_key_long_action_row()
        if not owner or row is None:
            return

        rec, owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        if row < 0 or row >= len(actions):
            return

        item = dict(actions[row])
        item["trigger"] = normalize_trigger(item.get("trigger", DEFAULT_TRIGGER))

        hold_mode = "range" if self.cb_key_long_hold_range.isChecked() else "fixed"
        hold_a = max(0.0, float(self.sp_key_long_hold_a.value()))
        hold_b = max(0.0, float(self.sp_key_long_hold_b.value())) if hold_mode == "range" else hold_a
        item["hold"] = {"mode": hold_mode, "a": hold_a, "b": hold_b}

        activate_mode = self.key_long_activation_slider.mode()
        item["activate_mode"] = "from_start" if activate_mode == "from_start" else "after_prev"

        start_mode = "range" if self.cb_key_long_start_range.isChecked() else "fixed"
        start_a = max(0.0, float(self.sp_key_long_start_a.value()))
        start_b = max(0.0, float(self.sp_key_long_start_b.value())) if start_mode == "range" else start_a
        item["start_delay"] = {"mode": start_mode, "a": start_a, "b": start_b}

        actions[row] = item
        self._set_key_long_hold_b_visible(hold_mode == "range")
        self._set_key_long_start_delay_visible(item["activate_mode"] == "from_start")
        self._commit_key_long_actions(rec, owner_row, owner_action, actions, select_row=row)

    def _on_key_long_hold_range_toggled(self, checked: bool):
        self._set_key_long_hold_b_visible(bool(checked))
        self._apply_selected_key_long_action_params()

    def _on_key_long_start_range_toggled(self, checked: bool):
        mode = self.key_long_activation_slider.mode() if hasattr(self, "key_long_activation_slider") else "after_prev"
        self._set_key_long_start_b_visible(bool(checked) and mode == "from_start")
        self._apply_selected_key_long_action_params()

    def _on_key_long_activation_mode_changed(self, mode: str):
        mode_norm = "from_start" if str(mode).strip().lower() == "from_start" else "after_prev"
        self._set_key_long_start_delay_visible(mode_norm == "from_start")
        self._apply_selected_key_long_action_params()

    def _capture_key_long_trigger(self, initial_trigger: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
        owner = self._get_key_long_actions_owner()
        if not owner:
            return None

        rec, owner_row, _owner_action = owner
        area = self._last_area_global(rec, up_to=owner_row)

        initial_action = None
        if isinstance(initial_trigger, dict):
            trig = normalize_trigger(initial_trigger)
            initial_action = KeyAction(kind=trig.get("kind", "keys"))
            initial_action.keys = list(trig.get("keys", []) or [])
            initial_action.mouse_button = trig.get("mouse_button", None)

        overlay = KeyCaptureOverlay(area_global=area, initial=initial_action)
        spec = overlay.show_and_block()
        if spec is None:
            return None
        return normalize_trigger(spec)

    def _move_selected_key_long_action(self, direction: int):
        owner = self._get_key_long_actions_owner()
        row = self._selected_key_long_action_row()
        if not owner or row is None:
            return

        rec, owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        j = row + int(direction)
        if j < 0 or j >= len(actions):
            return
        actions[row], actions[j] = actions[j], actions[row]
        self._commit_key_long_actions(rec, owner_row, owner_action, actions, select_row=j)

    def _delete_selected_key_long_action(self):
        owner = self._get_key_long_actions_owner()
        rows = self._selected_key_long_action_rows()
        if not owner or not rows:
            return

        rec, owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        target_row = rows[0]
        for r in reversed(rows):
            if 0 <= r < len(actions):
                actions.pop(r)
        if target_row >= len(actions):
            target_row = len(actions) - 1
        self._commit_key_long_actions(
            rec,
            owner_row,
            owner_action,
            actions,
            select_row=target_row if target_row >= 0 else None,
        )

    def _edit_selected_key_long_action(self):
        owner = self._get_key_long_actions_owner()
        row = self._selected_key_long_action_row()
        if not owner or row is None:
            return

        rec, owner_row, owner_action = owner
        actions = self._key_long_actions_from_key_action(owner_action)
        if row < 0 or row >= len(actions):
            return

        trigger = self._capture_key_long_trigger(actions[row].get("trigger"))
        if trigger is None:
            return
        actions[row]["trigger"] = trigger
        self._commit_key_long_actions(rec, owner_row, owner_action, actions, select_row=row)

    def _add_key_long_key_action(self):
        owner = self._get_key_long_actions_owner()
        if not owner:
            return
        rec, owner_row, owner_action = owner
        trigger = self._capture_key_long_trigger(None)
        if trigger is None:
            return
        actions = self._key_long_actions_from_key_action(owner_action)
        actions.append(self._default_key_long_action_item(trigger))
        self._commit_key_long_actions(rec, owner_row, owner_action, actions, select_row=len(actions) - 1)

    def _duplicate_selected_actions(self):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        if not rec:
            return

        rows = self._selected_action_rows()
        if not rows:
            return

        copied = [copy.deepcopy(rec.actions[r]) for r in rows]

        insert_at = len(rec.actions)
        rec.actions.extend(copied)

        self._refresh_actions()

        # выделим вставленные строки
        self.actions_table.clearSelection()
        if copied:
            r0 = insert_at
            r1 = insert_at + len(copied) - 1
            c1 = self.actions_table.columnCount() - 1
            rng = QTableWidgetSelectionRange(r0, 0, r1, c1)
            self.actions_table.setRangeSelected(rng, True)

            it = self.actions_table.item(r0, 0)
            if it:
                self.actions_table.scrollToItem(it, QAbstractItemView.PositionAtCenter)

        self._save()

    def _open_repeat_dialog(self):
        if getattr(self, "_playing", False):
            return
        self._refresh_repeat_ui()
        self._apply_dark_titlebar_widget(self.repeat_dialog)
        self.repeat_dialog.show()
        self.repeat_dialog.raise_()
        self.repeat_dialog.activateWindow()
        self._schedule_dark_titlebar(self.repeat_dialog)

    def _open_measure_dialog(self):
        if getattr(self, "_playing", False):
            return
        self._apply_dark_titlebar_widget(self.measure_dialog)
        self.measure_dialog.show()
        self.measure_dialog.raise_()
        self.measure_dialog.activateWindow()
        self._schedule_dark_titlebar(self.measure_dialog)

    def _open_app_settings_dialog(self):
        if hasattr(self, "app_lang_group"):
            self._set_single_choice_selector(self.app_lang_group, self._ui_language, default_code=i18n.DEFAULT_LANGUAGE)
        self._refresh_app_settings_list()
        self._on_app_settings_selection_changed()
        self._apply_dark_titlebar_widget(self.app_settings_dialog)
        self.app_settings_dialog.show()
        self.app_settings_dialog.raise_()
        self.app_settings_dialog.activateWindow()
        self._schedule_dark_titlebar(self.app_settings_dialog)

    def _open_project_github(self):
        QDesktopServices.openUrl(QUrl("https://github.com/true-meowmeow/Atari"))

    def _on_measure_dialog_closed(self, _result: int):
        if self.meter.is_running():
            self.meter.stop()

    def _toggle_measure(self):
        # во время проигрывания лучше не мешать
        if self.player and self.player.isRunning():
            QMessageBox.information(self, "Замер", "Остановите проигрывание перед замером.")
            return

        if self.meter.is_running():
            self.meter.stop()
        else:
            self._clear_measure()
            self.meter.start()
            self.btn_measure_toggle.setText("■ Стоп замера (F8)")

    def _on_meter_status(self, text: str):
        self.lbl_measure.setText(text)

    def _on_meter_stopped(self):
        self.btn_measure_toggle.setText("▶ Старт замера")

    def _on_meter_interval(self, dt_sec: float, key_name: str, idx: int):
        ms = dt_sec * 1000.0

        self.measure_table.insertRow(0)
        self.measure_table.setItem(0, 0, QTableWidgetItem(str(idx)))
        self.measure_table.setItem(0, 1, QTableWidgetItem(f"{ms:.1f}"))
        self.measure_table.setItem(0, 2, QTableWidgetItem(key_name))
        self.measure_table.scrollToTop()

        # оставим только последние 50 строк, чтобы таблица не разрасталась
        while self.measure_table.rowCount() > 50:
            self.measure_table.removeRow(self.measure_table.rowCount() - 1)

        # среднее
        vals = self.meter.intervals
        if vals:
            avg = sum(vals) / len(vals)
            self._measure_avg = avg
            self.lbl_measure.setText(
                f"Последнее: {ms:.1f} мс | Среднее: {avg * 1000.0:.1f} мс (N={len(vals)})"
            )

    def _clear_measure(self):
        self.measure_table.setRowCount(0)
        self._measure_avg = None
        self.lbl_measure.setText("Замер: выключен")

    def _apply_measure_avg_to_delay(self):
        if not self._measure_avg:
            QMessageBox.information(self, "Замер", "Нет данных замера (сначала нажмите пару клавиш).")
            return

        # применяем к выбранному действию (KeyAction) через спинбоксы
        if not self.key_params.isEnabled():
            QMessageBox.information(self, "Замер", "Выберите действие типа «Нажатие» в списке.")
            return

        v = float(self._measure_avg)
        self.sp_delay_a.setValue(v)
        # если режим фиксированный — синхронизируем B
        if not self.cb_random_delay.isChecked():
            self.sp_delay_b.setValue(v)

        # _apply_key_params сработает автоматически от valueChanged
        self._set_status(f"Задержка A выставлена: {v:.3f}с")

    def _copy_measure_avg(self):
        if not self._measure_avg:
            QMessageBox.information(self, "Замер", "Нет среднего значения для копирования.")
            return
        txt = f"{self._measure_avg:.3f}"
        QGuiApplication.clipboard().setText(txt)
        self._set_status(f"Скопировано: {txt} сек")

    def _on_player_action_row(self, row: int):
        self._highlight_action_row(row)

    def _highlight_action_row(self, row: int):
        prev = self._highlighted_row

        # снять прошлую подсветку (но если там ошибка — оставить красным)
        self._highlighted_row = None
        if prev is not None:
            self._apply_row_visual(prev)

        # поставить новую
        if row < 0 or row >= self.actions_table.rowCount():
            return

        self._highlighted_row = row
        self._apply_row_visual(row)

        it0 = self.actions_table.item(row, 0)
        if it0:
            self.actions_table.scrollToItem(it0, QAbstractItemView.PositionAtCenter)

    def _is_our_process_foreground(self) -> bool:
        """Чтобы F4 не конфликтовала с оверлеями."""
        if not _is_windows():
            return False
        try:
            fg = _hwnd_int(_user32.GetForegroundWindow())
            pid = _win_pid_from_hwnd(fg)
            return pid == os.getpid()
        except Exception:
            return False

    def _on_global_f4(self):
        if config.CAPTURE_OVERLAY_ACTIVE:
            return
        now = time.time()
        if now - self._last_f4_time < 0.25:
            return
        self._last_f4_time = now

        if self._is_our_process_foreground():
            return

        if self.player and self.player.isRunning():
            return
        self.play_current()

    def _on_global_f6(self):
        if config.CAPTURE_OVERLAY_ACTIVE:
            return
        now = time.time()
        if now - self._last_f6_time < 0.25:
            return
        self._last_f6_time = now
        if self.player and self.player.isRunning():
            self.resume_playback()

    def _on_global_f7(self):
        if config.CAPTURE_OVERLAY_ACTIVE:
            return
        now = time.time()
        if now - self._last_f7_time < 0.25:
            return
        self._last_f7_time = now
        if self.player and self.player.isRunning():
            self.stop_playback()

    def _on_global_esc(self):
        if config.CAPTURE_OVERLAY_ACTIVE:
            return
        # лёгкий антидребезг
        now = time.time()
        if now - self._last_esc_time < 0.2:
            return
        self._last_esc_time = now

        if self.player and self.player.isRunning():
            self.player.stop("Остановлено: ESC")

    def __init__(self):
        super().__init__()
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        self._settings_payload_cache: Dict[str, Any] = self._read_settings_payload()
        self._ui_language: str = i18n.normalize_language(self._settings_payload_cache.get("language"))
        self._ui_ready_for_language_change: bool = False
        i18n.set_language(self._ui_language)

        self.setWindowTitle("Atari")
        self.resize(1180, 720)

        self.records: List[Record] = []
        self.current_index: int = -1
        self._last_record_index: int = -1
        self._suspend_settings_autosave: bool = True
        self.player: Optional[MacroPlayer] = None

        # --- NEW: глобальная привязка базы к exe (не зависит от записи) ---
        self._bound_exe: str = ""
        self._bound_exe_enabled: bool = True
        self._bound_exe_recent: List[str] = []
        self._bound_exe_favorites: List[str] = []
        self._stop_word_cfg: Optional[dict] = None

        # NEW: enable можно переключать даже во время проигрывания
        self._stop_word_enabled_event = threading.Event()  # set() => enabled

        self._last_esc_time = 0.0
        self.hotkeys = GlobalHotkeyListener()
        self.hotkeys.esc_pressed.connect(self._on_global_esc)

        self.hotkeys.f4_pressed.connect(self._on_global_f4)

        # антидребезг
        self._last_f4_time = 0.0
        self._highlighted_row = None
        self._error_rows = set()
        self._params_row = None
        self._is_paused = False
        self._active_fail_actions_owner_row: Optional[int] = None
        self._active_key_long_actions_owner_row: Optional[int] = None

        self._build_ui()
        self._apply_style()
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self)

        QTimer.singleShot(0, self._apply_default_splitter)

        QTimer.singleShot(0, self._shrink_top_buttons)  # <-- вот это
        self._load()

        self._load_settings()
        if self.records:
            idx = self._last_record_index
            if idx < 0 or idx >= len(self.records):
                idx = 0
            self._set_current_record(idx)
        self._refresh_global_buttons()

        self._global_poll = QTimer(self)
        self._global_poll.timeout.connect(self._refresh_global_buttons)
        self._global_poll.start(1000)

        self.hotkeys.start()
        self._playing = False
        self.meter = IntervalMeter()
        self.meter.interval.connect(self._on_meter_interval)
        self.meter.status.connect(self._on_meter_status)
        self.meter.stopped.connect(self._on_meter_stopped)

        self._measure_avg = None  # float seconds
        self._suspend_settings_autosave = False
        self._ui_ready_for_language_change = True
        self._apply_ui_language(self._ui_language, persist=False, force_retranslate=True)

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ---- TOP BAR: Запись + текущая запись ----
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)

        self.btn_record_menu = QToolButton()
        self.btn_record_menu.setObjectName("top_tool_btn")
        self.btn_record_menu.setText("Запись")
        self.btn_record_menu.setPopupMode(QToolButton.InstantPopup)

        self.menu_record = QMenu(self)

        act_new = self.menu_record.addAction("Создать (Новую)")
        act_ren = self.menu_record.addAction("Переименовать (Текущую)")  # <-- НОВОЕ
        act_del = self.menu_record.addAction("Удалить (Текущую)")
        self.menu_record.addSeparator()
        act_save = self.menu_record.addAction("Сохранить…")
        act_load = self.menu_record.addAction("Загрузить…")

        act_new.triggered.connect(self.create_record)
        act_ren.triggered.connect(self.rename_record)  # <-- НОВОЕ
        act_del.triggered.connect(self.delete_record)
        act_save.triggered.connect(self.save_dialog)
        act_load.triggered.connect(self.load_dialog)

        self.btn_record_menu.setMenu(self.menu_record)

        self.btn_repeat_settings = QToolButton()
        self.btn_repeat_settings.setObjectName("top_tool_btn")
        self.btn_repeat_settings.setText("⚙")
        self.btn_repeat_settings.setProperty("state", "ok")
        self.btn_repeat_settings.clicked.connect(self._open_repeat_dialog)

        self.btn_current_record = QToolButton()
        self.btn_current_record.setObjectName("top_tool_btn")
        self.btn_current_record.setText("—")
        self.btn_current_record.setPopupMode(QToolButton.InstantPopup)

        self.menu_record_select = QMenu(self)
        self.btn_current_record.setMenu(self.menu_record_select)

        top_bar.addWidget(self.btn_record_menu)
        top_bar.addWidget(self.btn_current_record)
        top_bar.addWidget(self.btn_repeat_settings)
        top_bar.addStretch(1)

        # --- NEW: блок управления процессом и отдельными окнами ---
        self.btn_bind_base = QToolButton()
        self.btn_bind_base.setObjectName("top_tool_btn")
        self.btn_bind_base.setText("🎯 Приложение: не выбрано")
        self.btn_bind_base.setPopupMode(QToolButton.InstantPopup)
        self.menu_bind_app = QMenu(self)
        self.menu_bind_app.aboutToShow.connect(lambda: self._refresh_bind_app_menu(force=True))
        self.btn_bind_base.setMenu(self.menu_bind_app)

        self.btn_measure_window = QToolButton()
        self.btn_measure_window.setObjectName("top_tool_btn")
        self.btn_measure_window.setText("⏱ Замер интервалов")
        self.btn_measure_window.setProperty("state", "ok")
        self.btn_measure_window.clicked.connect(self._open_measure_dialog)

        self.btn_stop_word = QToolButton()
        self.btn_stop_word.setObjectName("top_tool_btn")
        self.btn_stop_word.setText("🛑 Стоп-слово: не задано")
        self.btn_stop_word.setPopupMode(QToolButton.InstantPopup)

        self.menu_stop_word = QMenu(self)
        act_sw_set = self.menu_stop_word.addAction("Задать/изменить…")
        act_sw_toggle = self.menu_stop_word.addAction("Включить/выключить")
        act_sw_clear = self.menu_stop_word.addAction("Убрать")

        act_sw_set.triggered.connect(self._stop_word_set)
        act_sw_toggle.triggered.connect(self._stop_word_toggle)
        act_sw_clear.triggered.connect(self._stop_word_clear)

        self.btn_stop_word.setMenu(self.menu_stop_word)

        self.btn_app_settings = QToolButton()
        self.btn_app_settings.setObjectName("top_tool_btn")
        self.btn_app_settings.setText("⚙ Настройки приложения")
        self.btn_app_settings.setProperty("state", "ok")
        self.btn_app_settings.clicked.connect(self._open_app_settings_dialog)

        top_bar.addWidget(self.btn_bind_base)
        top_bar.addWidget(self.btn_measure_window)
        top_bar.addWidget(self.btn_stop_word)
        top_bar.addWidget(self.btn_app_settings)

        layout.addLayout(top_bar)

        # Главный горизонтальный сплиттер: слева (записи+действия), справа (параметры)
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(6)  # чуть удобнее “барьер”

        self.records_panel = QWidget()
        rp = QVBoxLayout(self.records_panel)
        rp.setContentsMargins(0, 0, 0, 0)
        rp.setSpacing(8)

        self.lbl_records_title = QLabel("Записи")
        self.lbl_records_title.setFont(QFont("Segoe UI", 13, QFont.DemiBold))

        self.record_list = QListWidget()
        self.record_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.record_list.currentRowChanged.connect(self._on_record_selected)

        row_w = QWidget()
        row = QHBoxLayout(row_w)
        row.setContentsMargins(0, 0, 0, 0)
        self.btn_add_record = QPushButton("Создать")
        self.btn_del_record = QPushButton("Удалить")
        self.btn_ren_record = QPushButton("Переименовать")
        row.addWidget(self.btn_add_record)
        row.addWidget(self.btn_del_record)
        row.addWidget(self.btn_ren_record)

        row2_w = QWidget()
        row2 = QHBoxLayout(row2_w)
        row2.setContentsMargins(0, 0, 0, 0)
        self.btn_rec_up = QPushButton("▲")
        self.btn_rec_down = QPushButton("▼")
        row2.addWidget(self.btn_rec_up)
        row2.addWidget(self.btn_rec_down)
        row2.addStretch(1)



        # ✅ скрываем “дублирующую” панель записей
        self.records_panel.setVisible(False)

        self.btn_add_record.clicked.connect(self.create_record)
        self.btn_del_record.clicked.connect(self.delete_record)
        self.btn_ren_record.clicked.connect(self.rename_record)
        self.btn_rec_up.clicked.connect(lambda: self.move_record(-1))
        self.btn_rec_down.clicked.connect(lambda: self.move_record(+1))


        # --- (B) Действия (низ)
        actions_panel = QWidget()
        actions_l = QVBoxLayout(actions_panel)
        actions_l.setSpacing(10)

        # --- Records UI (вставляем реально в левую колонку) ---
        rp.addWidget(self.lbl_records_title)
        rp.addWidget(self.record_list)
        rp.addWidget(row_w)
        rp.addWidget(row2_w)

        actions_l.addWidget(self.records_panel)

        actions_panel.setMinimumWidth(0)
        self.record_list.setMinimumWidth(0)

        # ... позже, когда таблица уже создана:
        self.actions_table = QTableWidget(0, 4)
        self.actions_table.setMinimumWidth(0)  # <-- теперь ок

        self.actions_table.setHorizontalHeaderLabels(["#", "Действие", "Что делает", "Настройки"])
        self.actions_table.verticalHeader().setVisible(False)
        self.actions_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.actions_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.actions_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.actions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.actions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.actions_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.actions_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.actions_table.itemSelectionChanged.connect(self._on_action_selected)
        self.actions_table.cellDoubleClicked.connect(self._on_action_double_clicked)
        actions_l.addWidget(self.actions_table, 1)

        self._dup_shortcut = QShortcut(QKeySequence.Copy, self.actions_table)
        self._dup_shortcut.activated.connect(self._duplicate_selected_actions)

        self._del_shortcut = QShortcut(QKeySequence.Delete, self.actions_table)
        self._del_shortcut.activated.connect(self.delete_selected_action)

        # Одна строка: добавление действий + изменение порядка
        actions_controls_row = QHBoxLayout()

        self.btn_add_key = QPushButton("Нажатие")
        self.btn_add_area = QPushButton("Область")
        self.btn_add_wait = QPushButton("Ожидание")
        self.btn_act_up = QPushButton("Действие ▲")
        self.btn_act_down = QPushButton("Действие ▼")

        actions_controls_row.addWidget(self.btn_act_up)
        actions_controls_row.addWidget(self.btn_act_down)
        actions_controls_row.addWidget(self.btn_add_key)
        actions_controls_row.addWidget(self.btn_add_area)
        actions_controls_row.addWidget(self.btn_add_wait)
        actions_l.addLayout(actions_controls_row)

        # connections
        self.btn_add_area.clicked.connect(self.add_area_action)
        self.btn_add_key.clicked.connect(self.add_key_action)
        self.btn_add_wait.clicked.connect(self.add_wait_action)

        self.btn_act_up.clicked.connect(lambda: self.move_action(-1))
        self.btn_act_down.clicked.connect(lambda: self.move_action(+1))

        self.main_splitter.addWidget(actions_panel)


        # -------- ПРАВАЯ КОЛОНКА: параметры/повтор/проигрывание/замер (в скролле)
        settings_container = QWidget()
        settings_container.setObjectName("settings_container")  # <-- добавь
        settings_l = QVBoxLayout(settings_container)
        settings_l.setSpacing(10)

        self.action_params_section = QWidget()
        self.action_params_l = QVBoxLayout(self.action_params_section)
        self.action_params_l.setContentsMargins(0, 0, 0, 0)
        self.action_params_l.setSpacing(10)

        # --- сверху справа: Изменить / Удалить (50% / 50%) ---
        self.action_edit_row_w = QWidget()
        edit_row = QHBoxLayout(self.action_edit_row_w)
        edit_row.setContentsMargins(0, 0, 0, 0)

        self.btn_edit_action = QPushButton("Изменить")
        self.btn_del_action = QPushButton("Удалить")

        self.btn_edit_action.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_del_action.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_edit_action.clicked.connect(self.edit_selected_action)
        self.btn_del_action.clicked.connect(self.delete_selected_action)

        edit_row.addWidget(self.btn_edit_action, 1)
        edit_row.addWidget(self.btn_del_action, 1)

        self.action_params_l.addWidget(self.action_edit_row_w)
        self._set_action_edit_controls_visible(False)

        # Key-action parameters panel
        self.key_params = QGroupBox("")
        self.key_params.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.key_params_l = QVBoxLayout(self.key_params)
        self.key_params_l.setContentsMargins(0, 0, 0, 0)
        self.key_params_l.setSpacing(8)

        self.key_form_host = QWidget()
        self.key_form_host.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        form = QFormLayout(self.key_form_host)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignLeft)

        self.sp_multiplier = QSpinBox()
        self.sp_multiplier.setRange(1, 10000)
        self.sp_multiplier.setButtonSymbols(QSpinBox.NoButtons)
        self.sp_multiplier.setAlignment(Qt.AlignCenter)
        self.sp_multiplier.setMinimumWidth(90)
        self.sp_multiplier.setMaximumWidth(120)
        self.sp_multiplier.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.sp_multiplier.valueChanged.connect(self._apply_key_params)

        self.btn_multiplier_minus = QToolButton()
        self.btn_multiplier_minus.setObjectName("stepper_btn")
        self.btn_multiplier_minus.setText("-")
        self.btn_multiplier_minus.setCursor(Qt.PointingHandCursor)
        self.btn_multiplier_minus.setAutoRepeat(True)
        self.btn_multiplier_minus.setAutoRepeatDelay(220)
        self.btn_multiplier_minus.setAutoRepeatInterval(70)
        self.btn_multiplier_minus.clicked.connect(lambda _=False: self._step_spinbox(self.sp_multiplier, -1))

        self.btn_multiplier_plus = QToolButton()
        self.btn_multiplier_plus.setObjectName("stepper_btn")
        self.btn_multiplier_plus.setText("+")
        self.btn_multiplier_plus.setCursor(Qt.PointingHandCursor)
        self.btn_multiplier_plus.setAutoRepeat(True)
        self.btn_multiplier_plus.setAutoRepeatDelay(220)
        self.btn_multiplier_plus.setAutoRepeatInterval(70)
        self.btn_multiplier_plus.clicked.connect(lambda _=False: self._step_spinbox(self.sp_multiplier, +1))

        self.key_multiplier_row = QWidget()
        key_multiplier_l = QHBoxLayout(self.key_multiplier_row)
        key_multiplier_l.setContentsMargins(0, 0, 0, 0)
        key_multiplier_l.setSpacing(6)
        key_multiplier_l.addWidget(self.sp_multiplier, 1)
        key_multiplier_l.addWidget(self.btn_multiplier_minus)
        key_multiplier_l.addWidget(self.btn_multiplier_plus)
        key_multiplier_l.addStretch(1)

        self.cb_random_delay = ChipCheckBox("Диапазон")
        self.cb_random_delay.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.cb_random_delay.toggled.connect(self._on_random_delay_toggled)

        self.lbl_delay_a = QLabel("A")
        self.lbl_delay_b = QLabel("B")

        self.sp_delay_a = QDoubleSpinBox()
        self.sp_delay_a.setDecimals(3)
        self.sp_delay_a.setRange(0.0, 9999.0)
        self.sp_delay_a.setSingleStep(0.3)
        self.sp_delay_a.setValue(0.1)
        self.sp_delay_a.valueChanged.connect(self._apply_key_params)

        self.sp_delay_b = QDoubleSpinBox()
        self.sp_delay_b.setDecimals(3)
        self.sp_delay_b.setRange(0.0, 9999.0)
        self.sp_delay_b.setSingleStep(0.3)
        self.sp_delay_b.setValue(0.1)
        self.sp_delay_b.valueChanged.connect(self._apply_key_params)

        self.key_timing_row = QWidget()
        key_timing_l = QHBoxLayout(self.key_timing_row)
        key_timing_l.setContentsMargins(0, 0, 0, 0)
        key_timing_l.setSpacing(8)
        self.sp_delay_a.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.sp_delay_b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        key_timing_l.addWidget(self.cb_random_delay)
        key_timing_l.addWidget(self.lbl_delay_a)
        key_timing_l.addWidget(self.sp_delay_a, 2)
        key_timing_l.addWidget(self.lbl_delay_b)
        key_timing_l.addWidget(self.sp_delay_b, 2)

        self.lbl_key_multiplier = QLabel("Количество нажатий")
        self.lbl_key_timing = QLabel("Задержка до выполнения")
        form.addRow(self.lbl_key_multiplier, self.key_multiplier_row)
        form.addRow(self.lbl_key_timing, self.key_timing_row)
        self.key_long_actions_panel = QWidget()
        self.key_long_actions_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        key_long_actions_panel_l = QVBoxLayout(self.key_long_actions_panel)
        key_long_actions_panel_l.setContentsMargins(0, 0, 0, 0)
        key_long_actions_panel_l.setSpacing(8)

        self.key_long_actions_box = QGroupBox("Дополнительные действия")
        self.key_long_actions_box.setTitle("")
        self.key_long_actions_box.setObjectName("fail_actions_plain")
        self.key_long_actions_box.setFlat(True)
        self.key_long_actions_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        key_long_actions_l = QVBoxLayout(self.key_long_actions_box)
        key_long_actions_l.setContentsMargins(0, 0, 0, 0)
        key_long_actions_l.setSpacing(6)

        self.key_long_actions_top_row = QWidget()
        key_long_top_l = QHBoxLayout(self.key_long_actions_top_row)
        key_long_top_l.setContentsMargins(0, 0, 0, 0)
        key_long_top_l.setSpacing(8)

        self.key_long_actions_table = QTableWidget(0, 4)
        self.key_long_actions_table.setHorizontalHeaderLabels(["#", "Триггер", "Удержание", "Запуск"])
        self.key_long_actions_table.verticalHeader().setVisible(False)
        self.key_long_actions_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.key_long_actions_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.key_long_actions_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.key_long_actions_table.horizontalHeader().setVisible(False)
        self.key_long_actions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.key_long_actions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.key_long_actions_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.key_long_actions_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.key_long_actions_table.setMinimumHeight(180)
        self.key_long_actions_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.key_long_actions_table.cellDoubleClicked.connect(lambda _row, _col: self._edit_selected_key_long_action())
        self.key_long_actions_table.itemSelectionChanged.connect(self._sync_key_long_actions_buttons_state)
        key_long_top_l.addWidget(self.key_long_actions_table, 1)

        self.key_long_actions_buttons_col = QWidget()
        self.key_long_actions_buttons_col.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        key_long_buttons_l = QVBoxLayout(self.key_long_actions_buttons_col)
        key_long_buttons_l.setContentsMargins(0, 0, 0, 0)
        key_long_buttons_l.setSpacing(8)

        self.btn_key_long_actions_add = QToolButton()
        self.btn_key_long_actions_add.setText("▲")
        self.btn_key_long_actions_add.setCursor(Qt.PointingHandCursor)
        self.btn_key_long_actions_add.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.btn_key_long_actions_add.setMinimumWidth(42)
        self.btn_key_long_actions_add.setMaximumWidth(42)
        self.btn_key_long_actions_add.clicked.connect(lambda _=False: self._move_selected_key_long_action(-1))

        self.btn_key_long_actions_del = QToolButton()
        self.btn_key_long_actions_del.setText("▼")
        self.btn_key_long_actions_del.setCursor(Qt.PointingHandCursor)
        self.btn_key_long_actions_del.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.btn_key_long_actions_del.setMinimumWidth(42)
        self.btn_key_long_actions_del.setMaximumWidth(42)
        self.btn_key_long_actions_del.clicked.connect(lambda _=False: self._move_selected_key_long_action(+1))

        key_long_buttons_l.addWidget(self.btn_key_long_actions_add, 1)
        key_long_buttons_l.addWidget(self.btn_key_long_actions_del, 1)
        key_long_top_l.addWidget(self.key_long_actions_buttons_col, 0)

        self.key_long_actions_bottom_row = QWidget()
        key_long_bottom_l = QHBoxLayout(self.key_long_actions_bottom_row)
        key_long_bottom_l.setContentsMargins(0, 0, 0, 0)
        key_long_bottom_l.setSpacing(8)

        self.btn_key_long_bottom_add_key = QPushButton("Нажатие")
        self.btn_key_long_bottom_add_key.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_key_long_bottom_add_key.clicked.connect(self._add_key_long_key_action)

        self.btn_key_long_bottom_edit = QPushButton("Изменение")
        self.btn_key_long_bottom_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_key_long_bottom_edit.clicked.connect(self._edit_selected_key_long_action)

        self.btn_key_long_bottom_delete = QPushButton("Удаление")
        self.btn_key_long_bottom_delete.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_key_long_bottom_delete.clicked.connect(self._delete_selected_key_long_action)

        key_long_bottom_l.addWidget(self.btn_key_long_bottom_add_key, 1)
        key_long_bottom_l.addWidget(self.btn_key_long_bottom_edit, 1)
        key_long_bottom_l.addWidget(self.btn_key_long_bottom_delete, 1)

        self.key_long_params_host = QWidget()
        key_long_params_form = QFormLayout(self.key_long_params_host)
        key_long_params_form.setContentsMargins(0, 0, 0, 0)
        key_long_params_form.setLabelAlignment(Qt.AlignLeft)
        key_long_params_form.setSpacing(6)

        self.cb_key_long_hold_range = ChipCheckBox("Диапазон")
        self.cb_key_long_hold_range.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.cb_key_long_hold_range.toggled.connect(self._on_key_long_hold_range_toggled)
        self.lbl_key_long_hold_a = QLabel("A")
        self.lbl_key_long_hold_b = QLabel("B")
        self.sp_key_long_hold_a = QDoubleSpinBox()
        self.sp_key_long_hold_a.setDecimals(3)
        self.sp_key_long_hold_a.setRange(0.0, 9999.0)
        self.sp_key_long_hold_a.setSingleStep(0.3)
        self.sp_key_long_hold_a.setValue(0.2)
        self.sp_key_long_hold_a.valueChanged.connect(self._apply_selected_key_long_action_params)
        self.sp_key_long_hold_b = QDoubleSpinBox()
        self.sp_key_long_hold_b.setDecimals(3)
        self.sp_key_long_hold_b.setRange(0.0, 9999.0)
        self.sp_key_long_hold_b.setSingleStep(0.3)
        self.sp_key_long_hold_b.setValue(0.2)
        self.sp_key_long_hold_b.valueChanged.connect(self._apply_selected_key_long_action_params)
        self.key_long_hold_row = QWidget()
        key_long_hold_l = QHBoxLayout(self.key_long_hold_row)
        key_long_hold_l.setContentsMargins(0, 0, 0, 0)
        key_long_hold_l.setSpacing(8)
        key_long_hold_l.addWidget(self.cb_key_long_hold_range)
        key_long_hold_l.addWidget(self.lbl_key_long_hold_a)
        key_long_hold_l.addWidget(self.sp_key_long_hold_a, 2)
        key_long_hold_l.addWidget(self.lbl_key_long_hold_b)
        key_long_hold_l.addWidget(self.sp_key_long_hold_b, 2)
        key_long_params_form.addRow("Удержание", self.key_long_hold_row)

        self.key_long_activation_slider = LongActivationSlider()
        self.key_long_activation_slider.set_mode("after_prev", animate=False, emit_signal=False)
        self.key_long_activation_slider.modeChanged.connect(self._on_key_long_activation_mode_changed)
        key_long_params_form.addRow("Когда запускать", self.key_long_activation_slider)

        self.lbl_key_long_start_delay = QLabel("Смещение от старта")
        self.cb_key_long_start_range = ChipCheckBox("Диапазон")
        self.cb_key_long_start_range.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.cb_key_long_start_range.toggled.connect(self._on_key_long_start_range_toggled)
        self.lbl_key_long_start_a = QLabel("A")
        self.lbl_key_long_start_b = QLabel("B")
        self.sp_key_long_start_a = QDoubleSpinBox()
        self.sp_key_long_start_a.setDecimals(3)
        self.sp_key_long_start_a.setRange(0.0, 9999.0)
        self.sp_key_long_start_a.setSingleStep(0.3)
        self.sp_key_long_start_a.setValue(0.0)
        self.sp_key_long_start_a.valueChanged.connect(self._apply_selected_key_long_action_params)
        self.sp_key_long_start_b = QDoubleSpinBox()
        self.sp_key_long_start_b.setDecimals(3)
        self.sp_key_long_start_b.setRange(0.0, 9999.0)
        self.sp_key_long_start_b.setSingleStep(0.3)
        self.sp_key_long_start_b.setValue(0.0)
        self.sp_key_long_start_b.valueChanged.connect(self._apply_selected_key_long_action_params)
        self.key_long_start_delay_row = QWidget()
        key_long_start_l = QHBoxLayout(self.key_long_start_delay_row)
        key_long_start_l.setContentsMargins(0, 0, 0, 0)
        key_long_start_l.setSpacing(8)
        key_long_start_l.addWidget(self.cb_key_long_start_range)
        key_long_start_l.addWidget(self.lbl_key_long_start_a)
        key_long_start_l.addWidget(self.sp_key_long_start_a, 2)
        key_long_start_l.addWidget(self.lbl_key_long_start_b)
        key_long_start_l.addWidget(self.sp_key_long_start_b, 2)
        key_long_params_form.addRow(self.lbl_key_long_start_delay, self.key_long_start_delay_row)

        key_long_actions_l.addWidget(self.key_long_actions_top_row, 1)
        key_long_actions_l.addWidget(self.key_long_actions_bottom_row, 0)
        key_long_actions_l.addWidget(self.key_long_params_host, 0)
        key_long_actions_panel_l.addWidget(self.key_long_actions_box, 1)

        self._key_long_params_loading = False
        self._set_key_long_hold_b_visible(False)
        self._set_key_long_start_delay_visible(False)
        self._set_key_long_params_enabled(False)

        self.key_long_actions_panel.setVisible(False)
        form.addRow(self.key_long_actions_panel)

        self.key_params_l.addWidget(self.key_form_host, 1)
        self._set_delay_b_row_visible(False)

        self.action_params_l.addWidget(self.key_params)

        # Area-action parameters panel
        self.area_params = QGroupBox("")
        aform = QFormLayout(self.area_params)

        self.cb_area_click = ChipCheckBox("Действие в области")
        self.cb_area_click.toggled.connect(self._apply_area_params)

        self.btn_area_pick = QPushButton("ЛКМ")
        self.btn_area_reset = QPushButton("Сбросить")
        self.btn_area_pick.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_area_pick.clicked.connect(self._pick_area_trigger)
        self.btn_area_reset.clicked.connect(self._reset_area_trigger)

        self.area_action_state_row = QWidget()
        area_action_state_l = QHBoxLayout(self.area_action_state_row)
        area_action_state_l.setContentsMargins(0, 0, 0, 0)
        area_action_state_l.setSpacing(8)
        area_action_state_l.addWidget(self.cb_area_click)
        area_action_state_l.addWidget(self.btn_area_pick, 1)
        area_action_state_l.addWidget(self.btn_area_reset)
        area_action_state_l.addStretch(1)

        aform.addRow("Действие", self.area_action_state_row)

        self.sp_area_multiplier = QSpinBox()
        self.sp_area_multiplier.setRange(1, 10000)
        self.sp_area_multiplier.setButtonSymbols(QSpinBox.NoButtons)
        self.sp_area_multiplier.setAlignment(Qt.AlignCenter)
        self.sp_area_multiplier.setMinimumWidth(90)
        self.sp_area_multiplier.setMaximumWidth(120)
        self.sp_area_multiplier.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.sp_area_multiplier.valueChanged.connect(self._apply_area_params)

        self.btn_area_multiplier_minus = QToolButton()
        self.btn_area_multiplier_minus.setObjectName("stepper_btn")
        self.btn_area_multiplier_minus.setText("-")
        self.btn_area_multiplier_minus.setCursor(Qt.PointingHandCursor)
        self.btn_area_multiplier_minus.setAutoRepeat(True)
        self.btn_area_multiplier_minus.setAutoRepeatDelay(220)
        self.btn_area_multiplier_minus.setAutoRepeatInterval(70)
        self.btn_area_multiplier_minus.clicked.connect(lambda _=False: self._step_spinbox(self.sp_area_multiplier, -1))

        self.btn_area_multiplier_plus = QToolButton()
        self.btn_area_multiplier_plus.setObjectName("stepper_btn")
        self.btn_area_multiplier_plus.setText("+")
        self.btn_area_multiplier_plus.setCursor(Qt.PointingHandCursor)
        self.btn_area_multiplier_plus.setAutoRepeat(True)
        self.btn_area_multiplier_plus.setAutoRepeatDelay(220)
        self.btn_area_multiplier_plus.setAutoRepeatInterval(70)
        self.btn_area_multiplier_plus.clicked.connect(lambda _=False: self._step_spinbox(self.sp_area_multiplier, +1))

        self.area_multiplier_row = QWidget()
        area_multiplier_l = QHBoxLayout(self.area_multiplier_row)
        area_multiplier_l.setContentsMargins(0, 0, 0, 0)
        area_multiplier_l.setSpacing(6)
        area_multiplier_l.addWidget(self.sp_area_multiplier, 1)
        area_multiplier_l.addWidget(self.btn_area_multiplier_minus)
        area_multiplier_l.addWidget(self.btn_area_multiplier_plus)
        area_multiplier_l.addStretch(1)

        self.cb_area_random_delay = ChipCheckBox("Диапазон")
        self.cb_area_random_delay.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.cb_area_random_delay.toggled.connect(self._on_area_random_delay_toggled)

        self.lbl_area_delay_a = QLabel("A")
        self.lbl_area_delay_b = QLabel("B")

        self.sp_area_delay_a = QDoubleSpinBox()
        self.sp_area_delay_a.setDecimals(3)
        self.sp_area_delay_a.setRange(0.0, 9999.0)
        self.sp_area_delay_a.setSingleStep(0.3)
        self.sp_area_delay_a.setValue(0.1)
        self.sp_area_delay_a.valueChanged.connect(self._apply_area_params)

        self.sp_area_delay_b = QDoubleSpinBox()
        self.sp_area_delay_b.setDecimals(3)
        self.sp_area_delay_b.setRange(0.0, 9999.0)
        self.sp_area_delay_b.setSingleStep(0.3)
        self.sp_area_delay_b.setValue(0.1)
        self.sp_area_delay_b.valueChanged.connect(self._apply_area_params)

        self.area_timing_row = QWidget()
        area_timing_l = QHBoxLayout(self.area_timing_row)
        area_timing_l.setContentsMargins(0, 0, 0, 0)
        area_timing_l.setSpacing(8)
        self.sp_area_delay_a.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.sp_area_delay_b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        area_timing_l.addWidget(self.cb_area_random_delay)
        area_timing_l.addWidget(self.lbl_area_delay_a)
        area_timing_l.addWidget(self.sp_area_delay_a, 2)
        area_timing_l.addWidget(self.lbl_area_delay_b)
        area_timing_l.addWidget(self.sp_area_delay_b, 2)

        self.lbl_area_multiplier = QLabel("Количество нажатий")
        self.lbl_area_timing = QLabel("Задержка до выполнения")
        aform.addRow(self.lbl_area_multiplier, self.area_multiplier_row)
        aform.addRow(self.lbl_area_timing, self.area_timing_row)
        self._set_area_delay_b_row_visible(False)

        self.le_area_word = QLineEdit()
        self.le_area_word.setPlaceholderText("")
        self.le_area_word.textEdited.connect(self._apply_area_params)

        self.lbl_area_index = QLabel("Номер")
        self.lbl_area_word = QLabel("Текст")
        self.lbl_area_count = QLabel("Кол-во")
        self.sp_area_index = QSpinBox()
        self.sp_area_index.setRange(1, 9999)
        self.sp_area_index.valueChanged.connect(self._apply_area_params)
        self.sp_area_index.setMinimumWidth(70)
        self.sp_area_index.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.sp_area_count = QSpinBox()
        self.sp_area_count.setRange(1, 9999)
        self.sp_area_count.valueChanged.connect(self._apply_area_params)
        self.sp_area_count.setMinimumWidth(70)
        self.sp_area_count.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.area_word_row = QWidget()
        area_word_l = QHBoxLayout(self.area_word_row)
        area_word_l.setContentsMargins(0, 0, 0, 0)
        area_word_l.setSpacing(8)
        area_word_l.addWidget(self.sp_area_index)
        area_word_l.addWidget(self.lbl_area_count)
        area_word_l.addWidget(self.sp_area_count)
        area_word_l.addWidget(self.lbl_area_word)
        area_word_l.addWidget(self.le_area_word, 1)

        aform.addRow(self.lbl_area_index, self.area_word_row)

        self.lbl_area_ocr_lang = QLabel("Язык OCR")
        self.area_ocr_lang_row, self.area_ocr_lang_group = self._build_ocr_lang_selector(self._apply_area_params)
        aform.addRow(self.lbl_area_ocr_lang, self.area_ocr_lang_row)

        self.cb_area_search_infinite = ChipCheckBox("Искать бесконечно")
        self.cb_area_search_infinite.setChecked(True)
        self.cb_area_search_infinite.toggled.connect(self._on_area_search_infinite_toggled)

        self.lbl_area_search_max_tries = QLabel("Макс. попыток")
        self.sp_area_search_max_tries = QSpinBox()
        self.sp_area_search_max_tries.setRange(1, 1_000_000)
        self.sp_area_search_max_tries.setValue(100)
        self.sp_area_search_max_tries.setMinimumWidth(90)
        self.sp_area_search_max_tries.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.sp_area_search_max_tries.valueChanged.connect(self._apply_area_params)

        self.area_search_opts_row = QWidget()
        area_search_opts_l = QHBoxLayout(self.area_search_opts_row)
        area_search_opts_l.setContentsMargins(0, 0, 0, 0)
        area_search_opts_l.setSpacing(8)
        area_search_opts_l.addWidget(self.cb_area_search_infinite)
        area_search_opts_l.addWidget(self.lbl_area_search_max_tries)
        area_search_opts_l.addWidget(self.sp_area_search_max_tries)
        area_search_opts_l.addStretch(1)

        self.lbl_area_search_opts = QLabel("Поиск текста")
        aform.addRow(self.lbl_area_search_opts, self.area_search_opts_row)

        self.lbl_area_search_on_fail = QLabel("Если не найдено")
        self.area_search_on_fail_row, self.area_search_on_fail_group = self._build_single_choice_selector(
            [("retry", "Повторить поиск"), ("error", "Вывести ошибку"), ("action", "Действие")],
            self._on_area_search_on_fail_changed,
            default_code="retry",
        )
        aform.addRow(self.lbl_area_search_on_fail, self.area_search_on_fail_row)

        self.fail_actions_box = QGroupBox("Действия при «Если не найдено»")
        self.fail_actions_box.setTitle("")
        self.fail_actions_box.setObjectName("fail_actions_plain")
        self.fail_actions_box.setFlat(True)
        self.fail_actions_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        fail_l = QVBoxLayout(self.fail_actions_box)
        fail_l.setContentsMargins(0, 0, 0, 0)
        fail_l.setSpacing(6)

        self.fail_actions_table = QTableWidget(0, 4)
        self.fail_actions_table.setHorizontalHeaderLabels(["#", "Действие", "Что делает", "Настройки"])
        self.fail_actions_table.verticalHeader().setVisible(False)
        self.fail_actions_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.fail_actions_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.fail_actions_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.fail_actions_table.horizontalHeader().setVisible(False)
        self.fail_actions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.fail_actions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.fail_actions_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.fail_actions_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.fail_actions_table.setMinimumHeight(180)
        self.fail_actions_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        fail_l.addWidget(self.fail_actions_table, 1)

        self.fail_actions_controls_row = QWidget()
        fail_controls_l = QHBoxLayout(self.fail_actions_controls_row)
        fail_controls_l.setContentsMargins(0, 0, 0, 0)
        fail_controls_l.setSpacing(8)

        self.cb_focus_fail_actions = ChipCheckBox("Установить внимание")
        self.cb_focus_fail_actions.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cb_focus_fail_actions.toggled.connect(self._on_fail_actions_focus_toggled)
        fail_controls_l.addWidget(self.cb_focus_fail_actions, 2)
        fail_controls_l.addSpacing(18)

        self.cb_fail_actions_stop = ChipCheckBox("Остановиться")
        self.cb_fail_actions_stop.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cb_fail_actions_stop.toggled.connect(self._on_fail_actions_stop_toggled)
        fail_controls_l.addWidget(self.cb_fail_actions_stop, 1)

        self.cb_fail_actions_repeat = ChipCheckBox("Повторить")
        self.cb_fail_actions_repeat.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cb_fail_actions_repeat.toggled.connect(self._on_fail_actions_repeat_toggled)
        fail_controls_l.addWidget(self.cb_fail_actions_repeat, 1)

        fail_l.addWidget(self.fail_actions_controls_row)

        aform.addRow(self.fail_actions_box)
        self.fail_actions_box.setVisible(False)
        self._set_fail_actions_focus_checked(False)
        self._set_fail_actions_post_mode("none")
        self._set_area_search_max_tries_enabled(False)

        self.action_params_l.addWidget(self.area_params)
        self._set_area_params_enabled(False)

        # Wait-event parameters panel
        self.wait_event_params = QGroupBox("")
        we_form = QFormLayout(self.wait_event_params)

        self.lbl_wait_event_text = QLabel("Текст")
        self.le_wait_event_text = QLineEdit()
        self.le_wait_event_text.setPlaceholderText("Что должно совпасть")
        self.le_wait_event_text.textEdited.connect(self._apply_wait_event_params)
        we_form.addRow(self.lbl_wait_event_text, self.le_wait_event_text)

        self.sp_wait_event_poll = QDoubleSpinBox()
        self.sp_wait_event_poll.setDecimals(2)
        self.sp_wait_event_poll.setRange(0.1, 9999.0)
        self.sp_wait_event_poll.setSingleStep(0.3)
        self.sp_wait_event_poll.setValue(1.0)
        self.sp_wait_event_poll.valueChanged.connect(self._apply_wait_event_params)
        we_form.addRow("Период опроса (сек)", self.sp_wait_event_poll)

        self.wait_event_ocr_lang_row, self.wait_event_ocr_lang_group = self._build_ocr_lang_selector(
            self._apply_wait_event_params
        )
        we_form.addRow("Язык OCR", self.wait_event_ocr_lang_row)

        self.action_params_l.addWidget(self.wait_event_params)
        self._set_wait_event_params_enabled(False)

        self.action_key_mode_spacer = QWidget()
        self.action_key_mode_spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.action_params_l.addWidget(self.action_key_mode_spacer, 1)

        self.key_press_mode_slider = PressModeSlider()
        self.key_press_mode_slider.set_mode("normal", animate=False, emit_signal=False)
        self.key_press_mode_slider.modeChanged.connect(self._on_key_press_mode_changed)
        self.action_params_l.addWidget(self.key_press_mode_slider, 0)
        self._set_key_mode_widgets_visible(False)

        self.wait_mode_slider = WaitModeSlider()
        self.wait_mode_slider.set_mode("time", animate=False, emit_signal=False)
        self.wait_mode_slider.modeChanged.connect(self._on_wait_mode_changed)
        self.action_params_l.addWidget(self.wait_mode_slider, 0)
        self._set_wait_mode_widgets_visible(False)

        self.area_mode_slider = AreaModeSlider()
        self.area_mode_slider.set_mode("screen", animate=False, emit_signal=False)
        self.area_mode_slider.modeChanged.connect(self._on_area_mode_changed)
        self.action_params_l.addWidget(self.area_mode_slider, 0)
        self._set_area_mode_widgets_visible(False)

        settings_l.addWidget(self.action_params_section, 1)

        # Record settings (в отдельном окне)
        self.repeat_box = QGroupBox("Настройки записи")
        rform = QFormLayout(self.repeat_box)

        self.cb_repeat = ChipCheckBox("Повторять запись циклично")
        self.cb_repeat.toggled.connect(self._apply_repeat)

        self.sp_repeat_count = QSpinBox()
        self.sp_repeat_count.setRange(0, 1_000_000)
        self.sp_repeat_count.setToolTip("0 = бесконечно")
        self.sp_repeat_count.valueChanged.connect(self._apply_repeat)

        self.cb_repeat_random = ChipCheckBox("Диапазон")
        self.cb_repeat_random.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.cb_repeat_random.toggled.connect(self._apply_repeat)

        self.sp_repeat_a = QDoubleSpinBox()
        self.sp_repeat_a.setDecimals(3)
        self.sp_repeat_a.setRange(0.0, 9999.0)
        self.sp_repeat_a.setSingleStep(0.05)
        self.sp_repeat_a.valueChanged.connect(self._apply_repeat)

        self.sp_repeat_b = QDoubleSpinBox()
        self.sp_repeat_b.setDecimals(3)
        self.sp_repeat_b.setRange(0.0, 9999.0)
        self.sp_repeat_b.setSingleStep(0.05)
        self.sp_repeat_b.valueChanged.connect(self._apply_repeat)

        rform.addRow("", self.cb_repeat)
        rform.addRow("Количество проигрываний:", self.sp_repeat_count)
        self.lbl_repeat_delay_a = QLabel("A")
        self.lbl_repeat_b = QLabel("B")
        self.repeat_delay_row = QWidget()
        repeat_delay_row_l = QHBoxLayout(self.repeat_delay_row)
        repeat_delay_row_l.setContentsMargins(0, 0, 0, 0)
        repeat_delay_row_l.setSpacing(8)
        self.sp_repeat_a.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.sp_repeat_b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        repeat_delay_row_l.addWidget(self.cb_repeat_random)
        repeat_delay_row_l.addWidget(self.lbl_repeat_delay_a)
        repeat_delay_row_l.addWidget(self.sp_repeat_a, 2)
        repeat_delay_row_l.addWidget(self.lbl_repeat_b)
        repeat_delay_row_l.addWidget(self.sp_repeat_b, 2)

        self.lbl_repeat_timing = QLabel("Пауза")
        rform.addRow(self.lbl_repeat_timing, self.repeat_delay_row)
        self._set_repeat_b_row_visible(False)

        self.cb_move_mouse = ChipCheckBox("Перемещение мыши")
        self.cb_move_mouse.setChecked(True)
        self.cb_move_mouse.toggled.connect(self._apply_repeat)
        rform.addRow("", self.cb_move_mouse)

        self.bind_process_row = QWidget()
        bind_process_row_l = QHBoxLayout(self.bind_process_row)
        bind_process_row_l.setContentsMargins(0, 0, 0, 0)
        bind_process_row_l.setSpacing(0)

        self.cb_bind_record_process = ChipCheckBox("Привязать запись к процессу")
        self.cb_bind_record_process.toggled.connect(self._apply_repeat)
        self.lbl_bound_process_name = QLabel("Процесс: —")
        self.lbl_bound_process_name.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        bind_process_row_l.addWidget(self.cb_bind_record_process, 0)
        bind_process_row_l.addSpacing(18)
        bind_process_row_l.addWidget(self.lbl_bound_process_name, 1)
        rform.addRow("", self.bind_process_row)
        self.repeat_dialog = self.DarkTitleDialog(self, self)
        self.repeat_dialog.setWindowTitle("Настройка записи")
        self.repeat_dialog.setModal(False)
        self.repeat_dialog.setMinimumSize(520, 280)
        self.repeat_dialog.setAttribute(Qt.WA_NativeWindow, True)
        self.repeat_dialog.winId()
        repeat_dialog_l = QVBoxLayout(self.repeat_dialog)
        repeat_dialog_l.setContentsMargins(12, 12, 12, 12)
        repeat_dialog_l.addWidget(self.repeat_box)

        # ---- Measure box (в отдельном окне) ----
        self.measure_box = QGroupBox("Замер интервалов")
        self.measure_box.setObjectName("measure_box")

        mvl = QVBoxLayout(self.measure_box)
        mvl.setContentsMargins(10, 10, 10, 10)
        mvl.setSpacing(8)

        self.btn_measure_toggle = QPushButton("▶ Старт замера")
        self.btn_measure_toggle.clicked.connect(self._toggle_measure)

        self.lbl_measure = QLabel("Замер: выключен")

        btns_w = QWidget()
        btns = QHBoxLayout(btns_w)
        btns.setContentsMargins(0, 0, 0, 0)

        self.btn_measure_apply = QPushButton("Применить среднее → задержка A")
        self.btn_measure_apply.clicked.connect(self._apply_measure_avg_to_delay)

        self.btn_measure_copy = QPushButton("Копировать среднее")
        self.btn_measure_copy.clicked.connect(self._copy_measure_avg)

        self.btn_measure_clear = QPushButton("Очистить")
        self.btn_measure_clear.clicked.connect(self._clear_measure)

        btns.addWidget(self.btn_measure_apply)
        btns.addWidget(self.btn_measure_copy)
        btns.addWidget(self.btn_measure_clear)

        self.measure_table = QTableWidget(0, 3)
        self.measure_table.setHorizontalHeaderLabels(["#", "мс", "Кнопка"])
        self.measure_table.verticalHeader().setVisible(False)
        self.measure_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.measure_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.measure_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.measure_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.measure_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.measure_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.measure_table.setMinimumHeight(340)
        self.measure_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        mvl.addWidget(self.btn_measure_toggle)
        mvl.addWidget(self.lbl_measure)
        mvl.addWidget(btns_w)
        mvl.addWidget(self.measure_table, 1)

        self.measure_dialog = self.DarkTitleDialog(self, self, on_close=self._on_measure_dialog_closed)
        self.measure_dialog.setWindowTitle("Замер интервалов")
        self.measure_dialog.setModal(False)
        self.measure_dialog.setMinimumSize(720, 560)
        self.measure_dialog.setAttribute(Qt.WA_NativeWindow, True)
        self.measure_dialog.winId()
        measure_dialog_l = QVBoxLayout(self.measure_dialog)
        measure_dialog_l.setContentsMargins(12, 12, 12, 12)
        measure_dialog_l.addWidget(self.measure_box)

        self.app_settings_dialog = self.DarkTitleDialog(self, self)
        self.app_settings_dialog.setWindowTitle("Настройки приложения")
        self.app_settings_dialog.setModal(False)
        self.app_settings_dialog.setMinimumSize(520, 240)
        self.app_settings_dialog.setAttribute(Qt.WA_NativeWindow, True)
        self.app_settings_dialog.winId()

        app_settings_l = QVBoxLayout(self.app_settings_dialog)
        app_settings_l.setContentsMargins(12, 12, 12, 12)

        self.lbl_app_language = QLabel("Язык приложения")
        self.app_lang_row, self.app_lang_group = self._build_single_choice_selector(
            i18n.language_choices(),
            self._on_app_language_changed,
            default_code=self._ui_language,
        )
        app_settings_l.addWidget(self.lbl_app_language)
        app_settings_l.addWidget(self.app_lang_row)

        self.btn_project_github = QPushButton("GitHub: Atari")
        self.btn_project_github.setCursor(Qt.PointingHandCursor)
        self.btn_project_github.clicked.connect(self._open_project_github)
        app_settings_l.addWidget(self.btn_project_github)

        self.lw_app_history = QListWidget()
        self.lw_app_history.setSelectionMode(QAbstractItemView.SingleSelection)
        self.lw_app_history.itemSelectionChanged.connect(self._on_app_settings_selection_changed)
        app_settings_l.addWidget(self.lw_app_history, 1)

        app_btn_row = QHBoxLayout()
        self.btn_app_toggle_fav = QPushButton("Добавить в избранное")
        self.btn_app_fav_up = QPushButton("Избранное ▲")
        self.btn_app_fav_down = QPushButton("Избранное ▼")
        self.btn_app_clear_nonfav = QPushButton("Очистить не избранные")

        self.btn_app_toggle_fav.clicked.connect(self._toggle_app_favorite)
        self.btn_app_fav_up.clicked.connect(lambda: self._move_app_favorite(-1))
        self.btn_app_fav_down.clicked.connect(lambda: self._move_app_favorite(+1))
        self.btn_app_clear_nonfav.clicked.connect(self._clear_non_favorite_apps)

        app_btn_row.addWidget(self.btn_app_toggle_fav)
        app_btn_row.addWidget(self.btn_app_fav_up)
        app_btn_row.addWidget(self.btn_app_fav_down)
        app_btn_row.addStretch(1)
        app_btn_row.addWidget(self.btn_app_clear_nonfav)
        app_settings_l.addLayout(app_btn_row)

        # Playback bar (как отдельный виджет, чтобы нормально жил в скролле)
        play_bar_w = QWidget()
        play_bar = QHBoxLayout(play_bar_w)
        play_bar.setContentsMargins(0, 0, 0, 0)

        self.btn_play = QPushButton("▶ Проиграть (F4)")
        self.btn_pause = QPushButton("⏸ Пауза")
        self.btn_resume = QPushButton("▶ Продолжить")
        self.btn_stop = QPushButton("■ Стоп")

        self.lbl_status = ClickableLabel("Готово")
        self.lbl_status.setObjectName("status_label")
        self.lbl_status.setProperty("level", "info")
        self.lbl_status.setMinimumHeight(28)
        self.lbl_status.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.lbl_status.setCursor(Qt.PointingHandCursor)
        self.lbl_status.clicked.connect(self._clear_status_text)

        play_bar.addWidget(self.btn_play)
        play_bar.addWidget(self.btn_pause)
        play_bar.addWidget(self.btn_resume)
        play_bar.addWidget(self.btn_stop)
        play_bar.addWidget(self.lbl_status, 1)

        settings_l.addWidget(play_bar_w)

        self.btn_play.clicked.connect(self.play_current)
        self.btn_stop.clicked.connect(self.stop_playback)
        self.btn_pause.clicked.connect(self.pause_playback)
        self.btn_resume.clicked.connect(self.resume_playback)

        self._set_pause_controls(playing=False, paused=False)

        actions_panel.setMinimumWidth(520)  # левый столбец
        settings_container.setMinimumWidth(420)  # правый столбец

        self.main_splitter.addWidget(settings_container)

        # Пропорции по ширине
        self.main_splitter.setStretchFactor(0, 2)  # слева шире
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setSizes([720, 460])

        self._set_key_params_enabled(False)
        self._set_area_params_enabled(False)
        self._set_wait_event_params_enabled(False)
        self._set_action_param_panels_visible(key=False, area=False, wait_event=False)

        layout.addWidget(self.main_splitter, 1)

    def _apply_style(self):
        self.setStyleSheet("""
            /* 1) Чтобы окно (центральный QWidget тоже) не было белым */
            QMainWindow, QWidget { background: #0f1115; }

            QLabel { color: #e9eef7; }

            /* 2) Убираем белую подсветку выделения в списке и таблице */
            QListWidget::item:selected {
                background: #25324a;
                color: #e9eef7;
            }
            QListWidget::item:selected:!active {
                background: #1f293d;
                color: #e9eef7;
            }
            QTableWidget::item:selected {
                background: #2c3f60;
                color: #f4f8ff;
            }
            QTableWidget::item:selected:!active {
                background: #253650;
                color: #edf3fb;
            }
            QListWidget:focus, QTableWidget:focus { outline: 0; }

            /* (необязательно, но обычно приятнее) меню тоже в тёмном стиле */
            QMenuBar { background: #0f1115; color: #e9eef7; }
            QMenuBar::item:selected { background: #1b2332; }
            QMenu { background: #141821; color: #e9eef7; border: 1px solid #242a36; }
            QMenu::item:selected { background: #25324a; }

            QListWidget, QTableWidget {
                background: #141821; color: #e9eef7;
                border: 1px solid #242a36; border-radius: 10px;
                padding: 6px;
            }

            QScrollArea, QScrollArea::viewport {
                background: #0f1115;
                border: none;
            }
            QWidget#settings_container {
                background: #0f1115;
            }
            QLabel#status_label {
                background: #141821;
                border: 1px solid #242a36;
                border-radius: 10px;
                padding: 6px 10px;
            }
            
            QLabel#status_label[level="error"] {
                background: rgba(180, 60, 60, 170);
                border: 1px solid rgba(220, 110, 110, 220);
            }


            QLineEdit {
                background: #141821; color: #e9eef7;
                border: 1px solid #2a3447; border-radius: 10px;
                padding: 6px;
            }

            QHeaderView::section {
                background: #141821; color: #a9b3c7; border: none;
                padding: 6px 8px;
            }

            QPushButton {
                background: #1b2332; color: #e9eef7;
                border: 1px solid #2a3447;
                padding: 8px 10px; border-radius: 10px;
            }
            QPushButton:focus { outline: none; }
            QPushButton:hover { background: #222c40; }
            QPushButton:pressed { background: #161d2a; }


            QToolButton {
                background: #1b2332; color: #e9eef7;
                border: 1px solid #2a3447;
                padding: 8px 10px; border-radius: 10px;
            }
            QToolButton:focus { outline: none; }
            QToolButton:hover { background: #222c40; }
            QToolButton:pressed { background: #161d2a; }
            QToolButton::menu-indicator { image: none; } /* убираем стандартную стрелку */

            QToolButton#lang_chip {
                padding: 6px 12px;
                min-height: 22px;
                border-radius: 8px;
                background: #151c2a;
                border: 1px solid #2a3447;
                color: #c8d3e8;
            }
            QToolButton#lang_chip:hover {
                background: #1d293c;
                border: 1px solid #395271;
            }
            QToolButton#lang_chip:checked {
                background: #284162;
                border: 1px solid #4d78ad;
                color: #f4f8ff;
                font-weight: 600;
            }
            QToolButton#lang_chip:focus {
                outline: none;
            }
            QToolButton#stepper_btn {
                min-width: 68px;
                max-width: 68px;
                padding: 6px 0px;
                font-weight: 700;
            }

            QGroupBox {
                border: 1px solid #242a36; border-radius: 12px;
                margin-top: 10px; padding: 10px;
                color: #e9eef7;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 10px; padding: 0 6px;
                color: #a9b3c7;
            }
            QGroupBox#fail_actions_plain {
                border: none;
                margin-top: 0px;
                padding: 0px;
                background: transparent;
            }
            QGroupBox#fail_actions_plain::title {
                subcontrol-origin: margin;
                left: 0px;
                margin: 0px;
                padding: 0px;
                color: transparent;
            }

            QSpinBox, QDoubleSpinBox {
                background: #141821; color: #e9eef7;
                border: 1px solid #2a3447; border-radius: 10px;
                padding: 6px;
            }

            QCheckBox {
                color: #c8d3e8;
                background: #151c2a;
                border: 1px solid #2a3447;
                border-radius: 8px;
                padding: 6px 12px;
                min-height: 22px;
                spacing: 0px;
            }
            QCheckBox:hover {
                background: #1d293c;
                border: 1px solid #395271;
            }
            QCheckBox:checked {
                background: #284162;
                border: 1px solid #4d78ad;
                color: #f4f8ff;
            }
            QCheckBox:focus {
                outline: none;
            }
            QCheckBox:disabled {
                background: #11161f;
                border: 1px solid #2a3447;
                color: #697892;
            }
            QCheckBox::indicator {
                width: 0px;
                height: 0px;
                margin: 0px;
            }
            QCheckBox::indicator:unchecked,
            QCheckBox::indicator:checked {
                image: none;
            }
            
            /* ---- Sliders (QSlider): серые, без "сетчатого" фона ---- */
            QSlider::groove:horizontal {
                height: 8px;
                background: #2a2f36;          /* базовая дорожка */
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::sub-page:horizontal {
                background: #7a828c;          /* прогресс (светлее) */
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::add-page:horizontal {
                background: #1f2329;          /* НЕ прогресс (темнее и без сетки) */
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                width: 16px;
                margin: -5px 0;               /* чтобы ручка была по центру дорожки */
                background: #9aa3ad;          /* ручка серым */
                border: 1px solid #2a3447;
                border-radius: 8px;
            }
            QSlider::handle:horizontal:hover {
                background: #b0b9c3;
            }
            
            /* Вертикальные, если вдруг используешь */
            QSlider::groove:vertical {
                width: 8px;
                background: #2a2f36;
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::sub-page:vertical {
                background: #7a828c;
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::add-page:vertical {
                background: #1f2329;
                border: 1px solid #242a36;
                border-radius: 4px;
            }
            QSlider::handle:vertical {
                height: 16px;
                margin: 0 -5px;
                background: #9aa3ad;
                border: 1px solid #2a3447;
                border-radius: 8px;
            }
            QSlider::handle:vertical:hover {
                background: #b0b9c3;
            }
            QToolButton#top_tool_btn {
                padding: 3px 10px;      /* было 8px 10px -> из-за этого резало текст */
                min-height: 24px;       /* чтобы стиль не пытался “ужать” до нуля */
            }
            
            /* ===== Scrollbars (ползунки прокрутки) ===== */
            QScrollBar:vertical {
                background: #1f2329;      /* дорожка (темнее) */
                width: 12px;
                margin: 0px;
                border: none;
            }
            QScrollBar::handle:vertical {
                background: #7a828c;      /* ползунок (серый, светлее дорожки) */
                min-height: 28px;
                border-radius: 6px;
                border: 1px solid #242a36;
            }
            QScrollBar::handle:vertical:hover { background: #9099a3; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: #1f2329;      /* “не прогресс” — просто тёмно-серым, без сетки */
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;              /* убрать кнопки-стрелки */
                background: none;
                border: none;
            }
            
            QScrollBar:horizontal {
                background: #1f2329;
                height: 12px;
                margin: 0px;
                border: none;
            }
            QScrollBar::handle:horizontal {
                background: #7a828c;
                min-width: 28px;
                border-radius: 6px;
                border: 1px solid #242a36;
            }
            QScrollBar::handle:horizontal:hover { background: #9099a3; }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: #1f2329;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
                background: none;
                border: none;
            }
            
            /* ===== Splitter handle (та самая "сетчатая" ручка между панелями) ===== */
            QSplitter::handle {
                background: #1f2329;      /* темно-серый вместо "сетчатого" */
                border: none;
            }
            QSplitter::handle:hover {
                background: #2a2f36;
            }
            QToolButton[state="missing"] { color: #d67c7c; }  /* красноватый */
            QToolButton[state="ok"] { color: #e9eef7; }       /* обычный */
            QGroupBox::indicator { width: 0px; height: 0px; }


        """)
        # Диалоги отдельные от main window, поэтому дублируем стиль явно.
        if hasattr(self, "repeat_dialog"):
            self.repeat_dialog.setStyleSheet(self.styleSheet())
        if hasattr(self, "measure_dialog"):
            self.measure_dialog.setStyleSheet(self.styleSheet())
        if hasattr(self, "app_settings_dialog"):
            self.app_settings_dialog.setStyleSheet(self.styleSheet())

    # ---- Records ops ----
    def create_record(self):
        name, ok = self._input_text_dark("Создать запись", "Имя записи:")
        if not ok:
            return
        name = (name or "").strip() or "Новая запись"

        r = Record(name=name)
        r.actions = []  # <-- НЕТ base_area

        self.records.append(r)
        self._refresh_record_list()
        self._set_current_record(len(self.records) - 1)
        self._save()

    def delete_record(self):
        idx = self.current_index
        if idx < 0 or idx >= len(self.records):
            return

        if not self._show_dark_info_dialog(
            "Удалить",
            f"Удалить запись «{self.records[idx].name}»?",
            accept_text="Удалить",
            reject_text="Отмена",
            danger_accept=True,
        ):
            return
        self.stop_playback()
        self.records.pop(idx)
        self._refresh_record_list()
        self._set_current_record(min(idx, len(self.records) - 1))
        self._save()


    def rename_record(self):
        idx = self.current_index
        if idx < 0 or idx >= len(self.records):
            return

        name, ok = self._input_text_dark("Переименовать", "Новое имя:", text=self.records[idx].name)
        if not ok:
            return
        name = (name or "").strip()
        if not name:
            return
        self.records[idx].name = name
        self._refresh_record_list()
        self._set_current_record(idx)
        self._save()


    def move_record(self, direction: int):
        idx = self.current_index
        if idx < 0 or idx >= len(self.records):
            return

        j = idx + direction
        if j < 0 or j >= len(self.records):
            return
        self.records[idx], self.records[j] = self.records[j], self.records[idx]
        self._refresh_record_list()
        self._set_current_record(j)
        self._save()

    def _refresh_record_list(self):
        self._refresh_record_menu()
        self._refresh_record_list_widget()

    def _refresh_record_list_widget(self):
        """Левый QListWidget (если он у тебя используется/виден)."""
        if not hasattr(self, "record_list"):
            return

        self.record_list.blockSignals(True)
        self.record_list.clear()
        for r in self.records:
            self.record_list.addItem(QListWidgetItem(r.name))
        # синхронизируем выделение
        if 0 <= self.current_index < len(self.records):
            self.record_list.setCurrentRow(self.current_index)
        self.record_list.blockSignals(False)

    def _refresh_record_menu(self):
        """Меню второй кнопки (текущая запись)."""
        self.menu_record_select.clear()

        if not self.records:
            self.btn_current_record.setText("—")
            self.btn_current_record.setEnabled(False)
            return

        self.btn_current_record.setEnabled(True)

        group = QActionGroup(self.menu_record_select)
        group.setExclusive(True)

        for i, r in enumerate(self.records):
            act = QAction(r.name, self.menu_record_select)
            act.setCheckable(True)
            act.setChecked(i == self.current_index)
            act.triggered.connect(lambda _=False, idx=i: self._set_current_record(idx))
            group.addAction(act)
            self.menu_record_select.addAction(act)

    def _on_record_selected(self, row: int):
        # один источник правды — _set_current_record()
        self._set_current_record(row)

    # ---- Actions ops ----
    def add_area_action(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Сначала выберите или создайте запись.")
            return
        overlay = AreaSelectOverlay(initial_global=self._last_area_global(rec))
        rect = overlay.show_and_block()
        if rect and rect.isValid():
            base = self._get_base_area(rec)
            aa = AreaAction.from_global(rect, base, click=False, trigger=DEFAULT_TRIGGER, multiplier=1)
            if self._is_fail_actions_focus_active():
                self._append_fail_action(aa)
                return
            rec.actions.append(aa.to_dict())
            new_row = len(rec.actions) - 1
            self._refresh_actions(select_row=new_row)
            self._save()

    def add_key_action(self):
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Сначала выберите или создайте запись.")
            return
        area = self._last_area_global(rec)
        overlay = KeyCaptureOverlay(area_global=area, initial=None)
        spec = overlay.show_and_block()
        if spec is None:
            return
        ka = KeyAction(kind=spec.get("kind", "keys"))
        # include mods in keys list if spec says mouse: keep modifiers in keys for display, but action.kind decides behavior
        ka.keys = list(spec.get("keys", []))
        ka.mouse_button = spec.get("mouse_button", None)
        ka.multiplier = 1
        ka.delay = Delay("fixed", 0.1, 0.1)
        if self._is_fail_actions_focus_active():
            self._append_fail_action(ka)
            return
        rec.actions.append(ka.to_dict())
        new_row = len(rec.actions) - 1
        self._refresh_actions(select_row=new_row)
        self._save()

    def edit_selected_action(self):
        rec = self._current_record()
        if not rec:
            return

        if self._is_fail_actions_focus_active():
            owner = self._get_fail_actions_owner()
            row = self._selected_fail_action_row()
            if not owner or row is None:
                return
            owner_rec, owner_row, owner_action = owner
            if row < 0 or row >= len(owner_action.on_fail_actions):
                return

            a = action_from_dict(owner_action.on_fail_actions[row])

            if isinstance(a, AreaAction):
                base = self._get_base_area(rec)
                overlay = AreaSelectOverlay(initial_global=a.rect_global(base))
                rect = overlay.show_and_block()
                if rect and rect.isValid():
                    base = self._get_base_area(rec)
                    new_a = AreaAction.from_global(
                        rect, base, click=a.click,
                        trigger=getattr(a, "trigger", DEFAULT_TRIGGER),
                        multiplier=a.multiplier,
                        delay=a.delay,
                    )
                    owner_action.on_fail_actions[row] = new_a.to_dict()
                    self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
                    self._refresh_fail_actions(select_row=row)
                return

            if isinstance(a, WordAreaAction):
                base = self._get_base_area(rec)
                overlay = AreaSelectOverlay(initial_global=a.search_rect_global(base))
                rect = overlay.show_and_block()
                if rect and rect.isValid():
                    base = self._get_base_area(rec)
                    rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)
                    a.coord = "rel"
                    a.rx1, a.ry1, a.rx2, a.ry2 = rx1, ry1, rx2, ry2

                word, ok = self._input_text_dark("Область: Текст", "Текст для поиска:", text=a.word)
                if ok:
                    word = str(word or "")
                    if word.strip():
                        a.word = word

                idx, ok = self._input_int_dark(
                    "Номер совпадения",
                    "Какое по счёту совпадение брать? (1 = первое)",
                    int(getattr(a, "index", 1)),
                    1,
                    9999,
                    1,
                )
                if ok:
                    a.index = max(1, int(idx))

                owner_action.on_fail_actions[row] = a.to_dict()
                self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
                self._refresh_fail_actions(select_row=row)
                return

            if isinstance(a, WaitEventAction):
                base = self._get_base_area(rec)
                overlay = AreaSelectOverlay(initial_global=a.rect_global(base))
                rect = overlay.show_and_block()
                if rect and rect.isValid():
                    base = self._get_base_area(rec)
                    a = WaitEventAction.from_global(
                        rect, base, expected_text=a.expected_text, ocr_lang=a.ocr_lang, poll=a.poll,
                    )

                text, ok = self._input_text_dark("Ожидание: Событие", "Текст для ожидания:", text=a.expected_text)
                if not ok:
                    return
                text = str(text or "")
                if not text.strip():
                    self._set_status("Отмена изменения действия")
                    return
                a.expected_text = text

                owner_action.on_fail_actions[row] = a.to_dict()
                self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
                self._refresh_fail_actions(select_row=row)
                return

            if isinstance(a, KeyAction):
                area = self._last_area_global(rec)
                overlay = KeyCaptureOverlay(area_global=area, initial=a)
                spec = overlay.show_and_block()
                if spec is None:
                    return
                a.kind = spec.get("kind", a.kind)
                a.keys = list(spec.get("keys", a.keys))
                a.mouse_button = spec.get("mouse_button", a.mouse_button)
                owner_action.on_fail_actions[row] = a.to_dict()
                self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
                self._refresh_fail_actions(select_row=row)
                return

            return

        row = self._selected_action_row()
        if row is None:
            return

        a = action_from_dict(rec.actions[row])

        if isinstance(a, KeyAction):
            mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
            if mode == "long":
                return

        # --- AreaAction: правим прямоугольник ---
        if isinstance(a, AreaAction):
            base = self._get_base_area(rec)
            overlay = AreaSelectOverlay(initial_global=a.rect_global(base))
            rect = overlay.show_and_block()
            if rect and rect.isValid():
                base = self._get_base_area(rec)
                # сохраняем relative, но клики/триггер оставляем прежними
                new_a = AreaAction.from_global(
                    rect, base, click=a.click,
                    trigger=getattr(a, "trigger", DEFAULT_TRIGGER),
                    multiplier=a.multiplier,
                    delay=a.delay,
                )
                rec.actions[row] = new_a.to_dict()
                self._refresh_actions(select_row=row)
                self._sync_area_panel(new_a)
                self._save()
            return

        # --- WordAreaAction: правим и зону, и слово, и номер ---
        if isinstance(a, WordAreaAction):
            # зона поиска
            base = self._get_base_area(rec)
            overlay = AreaSelectOverlay(initial_global=a.search_rect_global(base))
            rect = overlay.show_and_block()
            if rect and rect.isValid():
                base = self._get_base_area(rec)
                rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)
                a.coord = "rel"
                a.rx1, a.ry1, a.rx2, a.ry2 = rx1, ry1, rx2, ry2

            # слово
            word, ok = self._input_text_dark("Область: Текст", "Текст для поиска:", text=a.word)
            if ok:
                word = str(word or "")
                if word.strip():
                    a.word = word

            # номер совпадения
            idx, ok = self._input_int_dark(
                "Номер совпадения",
                "Какое по счёту совпадение брать? (1 = первое)",
                int(getattr(a, "index", 1)),
                1,
                9999,
                1,
            )
            if ok:
                a.index = max(1, int(idx))

            rec.actions[row] = a.to_dict()
            self._refresh_actions(select_row=row)
            self._sync_area_panel(a)
            self._save()
            return

        if isinstance(a, WaitEventAction):
            base = self._get_base_area(rec)
            overlay = AreaSelectOverlay(initial_global=a.rect_global(base))
            rect = overlay.show_and_block()
            if rect and rect.isValid():
                base = self._get_base_area(rec)
                a = WaitEventAction.from_global(
                    rect, base, expected_text=a.expected_text, ocr_lang=a.ocr_lang, poll=a.poll,
                )

            text, ok = self._input_text_dark("Ожидание: Событие", "Текст для ожидания:", text=a.expected_text)
            if not ok:
                return
            text = str(text or "")
            if not text.strip():
                self._set_status("Отмена изменения действия")
                return
            a.expected_text = text

            rec.actions[row] = a.to_dict()
            self._refresh_actions(select_row=row)
            self._save()
            return

        # --- KeyAction: захват клавиши/мыши ---
        if isinstance(a, KeyAction):
            area = self._last_area_global(rec, up_to=row)
            overlay = KeyCaptureOverlay(area_global=area, initial=a)
            spec = overlay.show_and_block()
            if spec is None:
                return
            a.kind = spec.get("kind", a.kind)
            a.keys = list(spec.get("keys", a.keys))
            a.mouse_button = spec.get("mouse_button", a.mouse_button)
            rec.actions[row] = a.to_dict()
            self._refresh_actions(select_row=row)
            self._save()
            return

    def delete_selected_action(self):
        if getattr(self, "_playing", False):
            return

        rec = self._current_record()
        if not rec:
            return

        if self._is_fail_actions_focus_active():
            owner = self._get_fail_actions_owner()
            rows = self._selected_fail_action_rows()
            if not owner or not rows:
                return

            owner_rec, owner_row, owner_action = owner
            target_row = rows[0]
            for r in reversed(rows):
                if 0 <= r < len(owner_action.on_fail_actions):
                    owner_action.on_fail_actions.pop(r)

            self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
            if target_row >= len(owner_action.on_fail_actions):
                target_row = len(owner_action.on_fail_actions) - 1
            self._refresh_fail_actions(select_row=target_row if target_row >= 0 else None)
            return

        rows = self._selected_action_rows()
        if not rows:
            return

        # хотим выбрать "следующую" строку после удаления:
        # это будет индекс min(rows) (после удаления туда сдвинется то, что было ниже)
        target_row = rows[0]

        # не даём удалить базовую область (строка 0)
        if 0 in rows and rec.actions and isinstance(rec.actions[0], dict) and rec.actions[0].get("type") == "base_area":
            QMessageBox.information(self, "Нельзя удалить", "Опорная область обязательна и не удаляется.")
            rows = [r for r in rows if r != 0]
            if not rows:
                return

        anchor_idx = self._get_anchor_index(rec)
        if anchor_idx >= 0:
            if anchor_idx in rows:
                new_anchor = 0
            else:
                shift = sum(1 for r in rows if r < anchor_idx)
                new_anchor = anchor_idx - shift
        else:
            new_anchor = -1

        # удаляем с конца, чтобы индексы не съезжали
        for r in reversed(rows):
            if 0 <= r < len(rec.actions):
                rec.actions.pop(r)

        self._set_anchor_index(rec, new_anchor)
        self._refresh_actions()
        self._save()

        # восстановим выделение
        if not rec.actions:
            self.actions_table.clearSelection()
            self._set_key_params_enabled(False)
            self._set_area_params_enabled(False)
            self._set_wait_event_params_enabled(False)
            self._set_action_param_panels_visible(key=False, area=False, wait_event=False)
            return

        # если target_row уехал за конец — берем последний (это "выше", если удаляли последний)
        if target_row >= len(rec.actions):
            target_row = len(rec.actions) - 1

        self.actions_table.selectRow(target_row)
        it = self.actions_table.item(target_row, 0)
        if it:
            self.actions_table.scrollToItem(it, QAbstractItemView.PositionAtCenter)

    def move_action(self, direction: int):
        rec = self._current_record()
        if not rec:
            return

        if self._is_fail_actions_focus_active():
            owner = self._get_fail_actions_owner()
            row = self._selected_fail_action_row()
            if not owner or row is None:
                return
            owner_rec, owner_row, owner_action = owner
            actions = owner_action.on_fail_actions
            j = row + direction
            if j < 0 or j >= len(actions):
                return
            actions[row], actions[j] = actions[j], actions[row]
            self._save_fail_actions_owner(owner_rec, owner_row, owner_action)
            self._refresh_fail_actions(select_row=j)
            return

        row = self._selected_action_row()
        if row is None:
            return

        def is_base(i: int) -> bool:
            return (
                    0 <= i < len(rec.actions)
                    and isinstance(rec.actions[i], dict)
                    and rec.actions[i].get("type") == "base_area"
            )

        j = row + direction
        if j < 0 or j >= len(rec.actions):
            return

        # не двигаем base_area
        if is_base(row) or is_base(j):
            return

        anchor_idx = self._get_anchor_index(rec)
        if anchor_idx == row:
            new_anchor = j
        elif anchor_idx == j:
            new_anchor = row
        else:
            new_anchor = anchor_idx

        rec.actions[row], rec.actions[j] = rec.actions[j], rec.actions[row]
        self._set_anchor_index(rec, new_anchor)
        self._refresh_actions(select_row=j)
        self._save()

    def _refresh_actions(self, select_row: Optional[int] = None):
        rec = self._current_record()

        # запомним, что было выделено (если select_row не задан)
        prev_row = self._selected_action_row()

        self.actions_table.blockSignals(True)
        self.actions_table.setRowCount(0)
        self._set_key_params_enabled(False)
        self._set_area_params_enabled(False)
        self._set_wait_event_params_enabled(False)
        self._set_key_mode_widgets_visible(False)
        self._set_wait_mode_widgets_visible(False)
        self._set_area_mode_widgets_visible(False)
        self._set_action_edit_controls_visible(False)
        self._set_action_param_panels_visible(key=False, area=False, wait_event=False)
        self._set_params_row(None)

        if not rec:
            self.actions_table.blockSignals(False)
            return

        for i, ad in enumerate(rec.actions):
            a = action_from_dict(ad)
            t = self._action_type_text(a)
            desc = self._action_desc_text(a)
            params = self._action_params_text(a)

            self.actions_table.insertRow(i)
            self.actions_table.setItem(i, 0, QTableWidgetItem(self._action_row_label(i, rec)))
            self.actions_table.setItem(i, 1, QTableWidgetItem(t))
            self.actions_table.setItem(i, 2, QTableWidgetItem(desc))
            self.actions_table.setItem(i, 3, QTableWidgetItem(params))

        self.actions_table.blockSignals(False)

        # что выделять после обновления
        row = select_row if select_row is not None else prev_row
        if row is None:
            row = 0 if self.actions_table.rowCount() > 0 else -1

        if 0 <= row < self.actions_table.rowCount():
            self.actions_table.selectRow(row)
            it = self.actions_table.item(row, 0)
            if it:
                self.actions_table.scrollToItem(it, QAbstractItemView.PositionAtCenter)
        else:
            self.actions_table.clearSelection()

        for r in range(self.actions_table.rowCount()):
            self._apply_row_visual(r)

    def _selected_action_row(self) -> Optional[int]:
        sel = self.actions_table.selectionModel().selectedRows()
        if not sel:
            return None
        return sel[0].row()

    def _on_action_selected(self):
        if getattr(self, "_playing", False):
            return

        rec = self._current_record()
        rows = self._selected_action_rows()
        self._set_action_edit_controls_visible(bool(rec and rows))
        if not rec or len(rows) != 1:
            self._set_fail_actions_focus_checked(False)
            self._set_params_row(None)
            self._set_key_params_enabled(False)
            self._set_key_mode_widgets_visible(False)
            self._set_wait_mode_widgets_visible(False)
            self._set_area_mode_widgets_visible(False)
            self._set_area_params_enabled(False)
            self._set_wait_event_params_enabled(False)
            self._set_action_param_panels_visible(key=False, area=False, wait_event=False)
            return

        row = rows[0]
        self._set_params_row(row)
        a = action_from_dict(rec.actions[row])

        # --- KeyAction ---
        if isinstance(a, KeyAction):
            self._set_fail_actions_focus_checked(False)
            self._set_key_mode_widgets_visible(True)
            self._set_wait_mode_widgets_visible(False)
            self._set_area_mode_widgets_visible(False)
            self.sp_multiplier.setEnabled(True)
            self.key_params.setTitle("")

            self._set_area_params_enabled(False)
            self._set_wait_event_params_enabled(False)
            self._set_key_params_enabled(True)
            self._set_action_param_panels_visible(key=True, area=False, wait_event=False)

            self.sp_multiplier.blockSignals(True)
            self.cb_random_delay.blockSignals(True)
            self.sp_delay_a.blockSignals(True)
            self.sp_delay_b.blockSignals(True)

            self.sp_multiplier.setValue(max(1, a.multiplier))
            self.cb_random_delay.setChecked(a.delay.mode == "range")
            self.sp_delay_a.setValue(float(a.delay.a))
            self.sp_delay_b.setValue(float(a.delay.b))

            self.sp_multiplier.blockSignals(False)
            self.cb_random_delay.blockSignals(False)
            self.sp_delay_a.blockSignals(False)
            self.sp_delay_b.blockSignals(False)

            mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
            if hasattr(self, "key_press_mode_slider"):
                self.key_press_mode_slider.set_mode(mode, animate=False, emit_signal=False)
            self._apply_key_press_mode_layout(mode)
            return

        if isinstance(a, WaitAction):
            self._set_fail_actions_focus_checked(False)
            self._set_key_mode_widgets_visible(False)
            self._set_wait_mode_widgets_visible(True)
            self._set_area_mode_widgets_visible(False)
            self._set_area_params_enabled(False)
            self._set_wait_event_params_enabled(False)
            self._set_key_params_enabled(True)
            self._set_action_param_panels_visible(key=True, area=False, wait_event=False)
            if hasattr(self, "key_long_actions_panel"):
                self.key_long_actions_panel.setVisible(False)
            if hasattr(self, "wait_mode_slider"):
                self.wait_mode_slider.set_mode("time", animate=False, emit_signal=False)

            if hasattr(self, "lbl_key_multiplier"):
                self.lbl_key_multiplier.setVisible(False)
            if hasattr(self, "key_multiplier_row"):
                self.key_multiplier_row.setVisible(False)
            if hasattr(self, "lbl_key_timing"):
                self.lbl_key_timing.setVisible(True)
            if hasattr(self, "key_timing_row"):
                self.key_timing_row.setVisible(True)
            self.sp_multiplier.setEnabled(False)
            self.key_params.setTitle("")

            self.cb_random_delay.blockSignals(True)
            self.sp_delay_a.blockSignals(True)
            self.sp_delay_b.blockSignals(True)

            self.cb_random_delay.setChecked(a.delay.mode == "range")
            self.sp_delay_a.setValue(float(a.delay.a))
            self.sp_delay_b.setValue(float(a.delay.b))

            self.cb_random_delay.blockSignals(False)
            self.sp_delay_a.blockSignals(False)
            self.sp_delay_b.blockSignals(False)

            self._set_delay_b_row_visible(self.cb_random_delay.isChecked())
            return

        if isinstance(a, WaitEventAction):
            self._set_fail_actions_focus_checked(False)
            self._set_key_mode_widgets_visible(False)
            self._set_wait_mode_widgets_visible(True)
            self._set_area_mode_widgets_visible(False)
            self._set_area_params_enabled(False)
            self._set_key_params_enabled(False)
            self._set_wait_event_params_enabled(True)
            self._set_action_param_panels_visible(key=False, area=False, wait_event=True)
            if hasattr(self, "wait_mode_slider"):
                self.wait_mode_slider.set_mode("event", animate=False, emit_signal=False)
            self._sync_wait_event_panel(a)
            return

        # --- AreaAction / WordAreaAction ---
        if isinstance(a, (AreaAction, WordAreaAction)):
            self._set_key_mode_widgets_visible(False)
            self._set_wait_mode_widgets_visible(False)
            self._set_area_mode_widgets_visible(True)
            self._set_key_params_enabled(False)
            self._set_wait_event_params_enabled(False)
            self._set_area_params_enabled(True)
            self._set_action_param_panels_visible(key=False, area=True, wait_event=False)
            if hasattr(self, "area_mode_slider"):
                area_mode = "text" if isinstance(a, WordAreaAction) else "screen"
                self.area_mode_slider.set_mode(area_mode, animate=False, emit_signal=False)
            self._sync_area_panel(a)
            return

        # --- Other / unknown ---
        self._set_fail_actions_focus_checked(False)
        self._set_key_mode_widgets_visible(False)
        self._set_wait_mode_widgets_visible(False)
        self._set_area_mode_widgets_visible(False)
        self._set_key_params_enabled(False)
        self._set_area_params_enabled(False)
        self._set_wait_event_params_enabled(False)
        self._set_action_param_panels_visible(key=False, area=False, wait_event=False)

    def _on_action_double_clicked(self, row: int, _col: int):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        if not rec:
            return
        if row < 0 or row >= len(rec.actions):
            return

        prev = self._get_anchor_index(rec)
        if prev == row:
            return

        self._set_anchor_index(rec, row)
        self._update_action_row_number(prev)
        self._update_action_row_number(row)
        self._apply_row_visual(prev)
        self._apply_row_visual(row)

    def _set_action_param_panels_visible(self, key: bool, area: bool, wait_event: bool):
        self.key_params.setVisible(bool(key))
        self.area_params.setVisible(bool(area))
        self.wait_event_params.setVisible(bool(wait_event))

    def _set_action_edit_controls_visible(self, visible: bool):
        if hasattr(self, "action_edit_row_w"):
            self.action_edit_row_w.setVisible(bool(visible))

    def _normalize_key_press_mode(self, mode: Any) -> str:
        return "long" if str(mode or "").strip().lower() == "long" else "normal"

    def _normalize_wait_mode(self, mode: Any) -> str:
        return "event" if str(mode or "").strip().lower() == "event" else "time"

    def _normalize_area_mode(self, mode: Any) -> str:
        return "text" if str(mode or "").strip().lower() == "text" else "screen"

    def _set_key_mode_widgets_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "key_press_mode_slider"):
            self.key_press_mode_slider.setVisible(v)
        if hasattr(self, "action_key_mode_spacer"):
            self.action_key_mode_spacer.setVisible(True)
        if hasattr(self, "action_params_l") and hasattr(self, "key_params"):
            self.action_params_l.setStretchFactor(self.key_params, 0)
        if not v and hasattr(self, "key_long_actions_panel"):
            self.key_long_actions_panel.setVisible(False)
        if not v and hasattr(self, "key_long_actions_table"):
            self.key_long_actions_table.blockSignals(True)
            self.key_long_actions_table.setRowCount(0)
            self.key_long_actions_table.blockSignals(False)
        if not v and hasattr(self, "btn_key_long_actions_add"):
            self.btn_key_long_actions_add.setEnabled(False)
        if not v and hasattr(self, "btn_key_long_actions_del"):
            self.btn_key_long_actions_del.setEnabled(False)
        if not v and hasattr(self, "btn_key_long_bottom_add_key"):
            self.btn_key_long_bottom_add_key.setEnabled(False)
        if not v and hasattr(self, "btn_key_long_bottom_edit"):
            self.btn_key_long_bottom_edit.setEnabled(False)
        if not v and hasattr(self, "btn_key_long_bottom_delete"):
            self.btn_key_long_bottom_delete.setEnabled(False)
        if not v:
            self._set_key_long_params_enabled(False)
            self._clear_key_long_params_panel()

    def _set_wait_mode_widgets_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "wait_mode_slider"):
            self.wait_mode_slider.setVisible(v)
        if hasattr(self, "action_key_mode_spacer"):
            self.action_key_mode_spacer.setVisible(True)

    def _set_area_mode_widgets_visible(self, visible: bool):
        v = bool(visible)
        if hasattr(self, "area_mode_slider"):
            self.area_mode_slider.setVisible(v)
        if hasattr(self, "action_key_mode_spacer"):
            self.action_key_mode_spacer.setVisible(True)

    def _apply_key_press_mode_layout(self, mode: str):
        mode_norm = self._normalize_key_press_mode(mode)
        is_long = mode_norm == "long"
        normal_visible = not is_long

        if hasattr(self, "lbl_key_multiplier"):
            self.lbl_key_multiplier.setVisible(normal_visible)
        if hasattr(self, "key_multiplier_row"):
            self.key_multiplier_row.setVisible(normal_visible)
        if hasattr(self, "lbl_key_timing"):
            self.lbl_key_timing.setVisible(normal_visible)
        if hasattr(self, "key_timing_row"):
            self.key_timing_row.setVisible(normal_visible)
        if hasattr(self, "key_long_actions_panel"):
            self.key_long_actions_panel.setVisible(is_long)
        if hasattr(self, "action_key_mode_spacer"):
            self.action_key_mode_spacer.setVisible(not is_long)
        if hasattr(self, "action_params_l") and hasattr(self, "key_params"):
            self.action_params_l.setStretchFactor(self.key_params, 1 if is_long else 0)

        if normal_visible:
            self._set_delay_b_row_visible(self.cb_random_delay.isChecked())
            if hasattr(self, "btn_key_long_actions_add"):
                self.btn_key_long_actions_add.setEnabled(False)
            if hasattr(self, "btn_key_long_actions_del"):
                self.btn_key_long_actions_del.setEnabled(False)
            if hasattr(self, "btn_key_long_bottom_add_key"):
                self.btn_key_long_bottom_add_key.setEnabled(False)
            if hasattr(self, "btn_key_long_bottom_edit"):
                self.btn_key_long_bottom_edit.setEnabled(False)
            if hasattr(self, "btn_key_long_bottom_delete"):
                self.btn_key_long_bottom_delete.setEnabled(False)
            self._set_key_long_params_enabled(False)
            self._clear_key_long_params_panel()
        else:
            self._set_delay_b_row_visible(False)
            if hasattr(self, "key_long_actions_table"):
                self._refresh_key_long_actions()

    def _key_action_params_text(self, a: KeyAction) -> str:
        mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
        if mode == "long":
            return "Продолжительный"
        return f"Обычный; x{a.multiplier}; {self._delay_text(a.delay)}"

    def _on_key_press_mode_changed(self, mode: str):
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return
        a = action_from_dict(rec.actions[row])
        if not isinstance(a, KeyAction):
            return

        mode_norm = self._normalize_key_press_mode(mode)
        self._apply_key_press_mode_layout(mode_norm)

        if getattr(a, "press_mode", "normal") == mode_norm:
            if mode_norm == "long":
                self._refresh_key_long_actions()
            return
        a.press_mode = mode_norm
        rec.actions[row] = a.to_dict()
        self._update_actions_table_row(row, a)
        self._save()
        if mode_norm == "long":
            self._refresh_key_long_actions()

    def _on_wait_mode_changed(self, mode: str):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return

        a = action_from_dict(rec.actions[row])
        if not isinstance(a, (WaitAction, WaitEventAction)):
            return

        mode_norm = self._normalize_wait_mode(mode)
        current_mode = "event" if isinstance(a, WaitEventAction) else "time"
        if current_mode == mode_norm:
            return

        if mode_norm == "event":
            base = self._get_base_area(rec)
            rect = self._last_area_global(rec, up_to=row)
            if not rect or not rect.isValid():
                rect = base
            if not rect or not rect.isValid():
                return
            text = ""
            if hasattr(self, "le_wait_event_text"):
                text = str(self.le_wait_event_text.text() or "")
            poll = 1.0
            if hasattr(self, "sp_wait_event_poll"):
                poll = max(0.1, float(self.sp_wait_event_poll.value()))
            ocr_lang = self._default_ocr_lang()
            if hasattr(self, "wait_event_ocr_lang_group"):
                ocr_lang = self._get_ocr_lang_selector(self.wait_event_ocr_lang_group)
            new_a: Action = WaitEventAction.from_global(
                rect,
                base,
                expected_text=text,
                ocr_lang=ocr_lang,
                poll=poll,
            )
        else:
            sec = 1.0
            if isinstance(a, WaitEventAction):
                try:
                    sec = max(0.0, float(getattr(a, "poll", 1.0)))
                except Exception:
                    sec = 1.0
            new_a = WaitAction(delay=Delay("fixed", sec, sec))

        rec.actions[row] = new_a.to_dict()
        self._update_actions_table_row(row, new_a)
        if 0 <= row < self.actions_table.rowCount():
            self.actions_table.selectRow(row)
        self._on_action_selected()
        self._save()

    def _on_area_mode_changed(self, mode: str):
        if getattr(self, "_playing", False):
            return
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return

        a = action_from_dict(rec.actions[row])
        if not isinstance(a, (AreaAction, WordAreaAction)):
            return

        mode_norm = self._normalize_area_mode(mode)
        current_mode = "text" if isinstance(a, WordAreaAction) else "screen"
        if current_mode == mode_norm:
            return

        if mode_norm == "text":
            trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
            delay = getattr(a, "delay", Delay("fixed", 0.1, 0.1))
            delay_copy = Delay.from_dict(delay.to_dict()) if isinstance(delay, Delay) else Delay("fixed", 0.1, 0.1)
            idx = 1
            if hasattr(self, "sp_area_index"):
                try:
                    idx = max(1, int(self.sp_area_index.value()))
                except Exception:
                    idx = 1
            cnt = 1
            if hasattr(self, "sp_area_count"):
                try:
                    cnt = max(1, int(self.sp_area_count.value()))
                except Exception:
                    cnt = 1
            search_max_tries = 100
            if hasattr(self, "sp_area_search_max_tries"):
                try:
                    search_max_tries = max(1, int(self.sp_area_search_max_tries.value()))
                except Exception:
                    search_max_tries = 100

            new_a = WordAreaAction(
                word=str(self.le_area_word.text() or "") if hasattr(self, "le_area_word") else "",
                index=idx,
                count=cnt,
                click=bool(getattr(a, "click", False)),
                multiplier=max(1, int(getattr(a, "multiplier", 1))),
                delay=delay_copy,
                button="left",
                trigger=trig,
                ocr_lang=(
                    self._get_ocr_lang_selector(self.area_ocr_lang_group)
                    if hasattr(self, "area_ocr_lang_group")
                    else self._default_ocr_lang()
                ),
                search_infinite=bool(
                    getattr(self, "cb_area_search_infinite", None) and self.cb_area_search_infinite.isChecked()
                ),
                search_max_tries=search_max_tries,
                search_on_fail=(
                    self._get_single_choice_selector(self.area_search_on_fail_group, default_code="retry")
                    if hasattr(self, "area_search_on_fail_group")
                    else "retry"
                ),
                on_fail_actions=[],
                on_fail_post_mode=self._get_fail_actions_post_mode() if hasattr(self, "_get_fail_actions_post_mode") else "none",
            )
            if trig["kind"] == "mouse" and trig["mouse_button"] in ("left", "middle", "right"):
                new_a.button = trig["mouse_button"]

            if isinstance(a, AreaAction):
                if a.coord == "rel":
                    new_a.coord = "rel"
                    new_a.rx1, new_a.ry1, new_a.rx2, new_a.ry2 = a.rx1, a.ry1, a.rx2, a.ry2
                else:
                    new_a.coord = "abs"
                    new_a.x1, new_a.y1, new_a.x2, new_a.y2 = a.x1, a.y1, a.x2, a.y2
        else:
            trig = normalize_trigger(getattr(a, "trigger", DEFAULT_TRIGGER))
            delay = getattr(a, "delay", Delay("fixed", 0.1, 0.1))
            delay_copy = Delay.from_dict(delay.to_dict()) if isinstance(delay, Delay) else Delay("fixed", 0.1, 0.1)
            new_a = AreaAction(
                click=bool(getattr(a, "click", False)),
                multiplier=max(1, int(getattr(a, "multiplier", 1))),
                delay=delay_copy,
                trigger=trig,
            )
            if isinstance(a, WordAreaAction):
                if a.coord == "rel":
                    new_a.coord = "rel"
                    new_a.rx1, new_a.ry1, new_a.rx2, new_a.ry2 = a.rx1, a.ry1, a.rx2, a.ry2
                else:
                    new_a.coord = "abs"
                    new_a.x1, new_a.y1, new_a.x2, new_a.y2 = a.x1, a.y1, a.x2, a.y2

        rec.actions[row] = new_a.to_dict()
        self._update_actions_table_row(row, new_a)
        if 0 <= row < self.actions_table.rowCount():
            self.actions_table.selectRow(row)
        self._on_action_selected()
        self._save()

    def _set_key_params_enabled(self, en: bool):
        self.key_params.setEnabled(en)
        self._set_delay_b_row_visible(self.cb_random_delay.isChecked())

    def _set_delay_b_row_visible(self, visible: bool):
        visible = bool(visible)
        if hasattr(self, "lbl_delay_b"):
            self.lbl_delay_b.setVisible(visible)
        self.sp_delay_b.setVisible(visible)
        self.sp_delay_b.setEnabled(visible)

    def _set_area_delay_b_row_visible(self, visible: bool):
        visible = bool(visible)
        if hasattr(self, "lbl_area_delay_b"):
            self.lbl_area_delay_b.setVisible(visible)
        self.sp_area_delay_b.setVisible(visible)
        self.sp_area_delay_b.setEnabled(visible)

    def _set_area_search_max_tries_enabled(self, enabled: bool):
        en = bool(enabled)
        if hasattr(self, "lbl_area_search_max_tries"):
            self.lbl_area_search_max_tries.setEnabled(en)
        if hasattr(self, "sp_area_search_max_tries"):
            self.sp_area_search_max_tries.setEnabled(en)

    def _on_area_search_infinite_toggled(self, checked: bool):
        if bool(checked) and hasattr(self, "area_search_on_fail_group"):
            on_fail = self._get_single_choice_selector(self.area_search_on_fail_group, default_code="retry")
            if on_fail == "action":
                self._set_single_choice_selector(self.area_search_on_fail_group, "retry", default_code="retry")
        self._set_area_search_max_tries_enabled(not bool(checked))
        self._apply_area_params()

    def _on_area_search_on_fail_changed(self):
        if not hasattr(self, "area_search_on_fail_group"):
            self._apply_area_params()
            return

        on_fail = self._get_single_choice_selector(self.area_search_on_fail_group, default_code="retry")
        if on_fail == "action" and hasattr(self, "cb_area_search_infinite") and self.cb_area_search_infinite.isChecked():
            self.cb_area_search_infinite.blockSignals(True)
            self.cb_area_search_infinite.setChecked(False)
            self.cb_area_search_infinite.blockSignals(False)
            self._set_area_search_max_tries_enabled(True)

        self._apply_area_params()

    def _set_repeat_b_row_visible(self, visible: bool):
        visible = bool(visible)
        if hasattr(self, "lbl_repeat_b"):
            self.lbl_repeat_b.setVisible(visible)
        self.sp_repeat_b.setVisible(visible)
        self.sp_repeat_b.setEnabled(visible)

    def _on_random_delay_toggled(self, checked: bool):
        # 1) UI
        self._set_delay_b_row_visible(checked)
        # 2) применяем в модель (как было раньше)
        self._apply_key_params()

    def _on_area_random_delay_toggled(self, checked: bool):
        self._set_area_delay_b_row_visible(checked)
        self._apply_area_params()

    def _apply_key_params(self):
        rec = self._current_record()
        row = self._selected_action_row()
        if not rec or row is None:
            return

        a = action_from_dict(rec.actions[row])

        if isinstance(a, KeyAction):
            mode = self._normalize_key_press_mode(getattr(a, "press_mode", "normal"))
            a.press_mode = mode
            if mode == "normal":
                a.multiplier = max(1, int(self.sp_multiplier.value()))
                if self.cb_random_delay.isChecked():
                    a.delay.mode = "range"
                    a.delay.a = float(self.sp_delay_a.value())
                    a.delay.b = float(self.sp_delay_b.value())
                else:
                    a.delay.mode = "fixed"
                    a.delay.a = float(self.sp_delay_a.value())
                    a.delay.b = float(self.sp_delay_a.value())

            rec.actions[row] = a.to_dict()
            if mode == "normal":
                self._set_delay_b_row_visible(self.cb_random_delay.isChecked())
            else:
                self._set_delay_b_row_visible(False)
            self._update_actions_table_row(row, a)
            self._save()
            return

        if isinstance(a, WaitAction):
            if self.cb_random_delay.isChecked():
                a.delay.mode = "range"
                a.delay.a = float(self.sp_delay_a.value())
                a.delay.b = float(self.sp_delay_b.value())
            else:
                a.delay.mode = "fixed"
                a.delay.a = float(self.sp_delay_a.value())
                a.delay.b = float(self.sp_delay_a.value())

            rec.actions[row] = a.to_dict()
            self._set_delay_b_row_visible(self.cb_random_delay.isChecked())
            self._update_actions_table_row(row, a)
            self._save()
            return

    def _last_area_global(self, rec: Record, up_to: Optional[int] = None) -> Optional[QRect]:
        actions = rec.actions if up_to is None else rec.actions[:max(0, up_to)]
        base = self._get_base_area(rec)
        last = base

        for ad in actions:
            a = action_from_dict(ad)
            if isinstance(a, BaseAreaAction):
                base = self._get_base_area(rec)  # учитывает bound_exe
                last = base
            elif isinstance(a, AreaAction):
                last = a.rect_global(base)
            elif isinstance(a, WordAreaAction):
                last = a.search_rect_global(base)
            elif isinstance(a, WaitEventAction):
                last = a.rect_global(base)

        return last

    def _validate_actions_before_play(self, rec: Record) -> bool:
        for idx, ad in enumerate(rec.actions):
            try:
                a = action_from_dict(ad)
            except Exception:
                continue

            if isinstance(a, WordAreaAction):
                if not str(getattr(a, "word", "") or "").strip():
                    self._set_status(
                        f"Ошибка: в действии #{idx + 1} «Область (Текст)» поле «Текст» пустое.",
                        level="error",
                    )
                    if 0 <= idx < self.actions_table.rowCount():
                        self.actions_table.selectRow(idx)
                    return False

            if isinstance(a, WaitEventAction):
                if not str(getattr(a, "expected_text", "") or "").strip():
                    self._set_status(
                        f"Ошибка: в действии #{idx + 1} «Ожидание (Событие)» поле «Текст» пустое.",
                        level="error",
                    )
                    if 0 <= idx < self.actions_table.rowCount():
                        self.actions_table.selectRow(idx)
                    return False

        return True

    def _refresh_bound_process_caption(self, rec: Optional[Record] = None):
        if not hasattr(self, "lbl_bound_process_name"):
            return

        rec = rec if rec is not None else self._current_record()
        if not rec or not bool(getattr(rec, "bind_to_process", False)):
            self.lbl_bound_process_name.setText("Процесс: —")
            return

        bound_exe = str(getattr(rec, "bound_exe", "") or "").strip()
        temp_exe = str(getattr(rec, "bound_exe_override", "") or "").strip()
        shown_exe = temp_exe or bound_exe
        if not shown_exe:
            self.lbl_bound_process_name.setText("Процесс: не задан")
            return

        name = os.path.basename(shown_exe) or shown_exe
        suffix = " (временное)" if temp_exe else ""
        self.lbl_bound_process_name.setText(f"Процесс: {name}{suffix}")

    # ---- Record settings ----
    def _refresh_repeat_ui(self):
        rec = self._current_record()
        if not rec:
            self.repeat_box.setEnabled(False)
            self.btn_repeat_settings.setEnabled(False)
            self._refresh_bound_process_caption(None)
            return
        self.repeat_box.setEnabled(True)
        self.btn_repeat_settings.setEnabled(not getattr(self, "_playing", False))
        rs = rec.repeat

        self.cb_repeat.blockSignals(True)
        self.sp_repeat_count.blockSignals(True)
        self.cb_repeat_random.blockSignals(True)
        self.sp_repeat_a.blockSignals(True)
        self.sp_repeat_b.blockSignals(True)
        self.cb_bind_record_process.blockSignals(True)
        self.cb_move_mouse.blockSignals(True)

        self.cb_repeat.setChecked(rs.enabled)
        self.sp_repeat_count.setValue(rs.count)
        self.cb_repeat_random.setChecked(rs.delay.mode == "range")
        self.sp_repeat_a.setValue(rs.delay.a)
        self.sp_repeat_b.setValue(rs.delay.b)
        self.cb_bind_record_process.setChecked(bool(getattr(rec, "bind_to_process", False)))
        self.cb_move_mouse.setChecked(bool(getattr(rec, "move_mouse", True)))

        self.cb_repeat.blockSignals(False)
        self.sp_repeat_count.blockSignals(False)
        self.cb_repeat_random.blockSignals(False)
        self.sp_repeat_a.blockSignals(False)
        self.sp_repeat_b.blockSignals(False)
        self.cb_bind_record_process.blockSignals(False)
        self.cb_move_mouse.blockSignals(False)

        self._set_repeat_b_row_visible(self.cb_repeat_random.isChecked())
        self._refresh_bound_process_caption(rec)

    def _apply_repeat(self):
        rec = self._current_record()
        if not rec:
            return
        prev_bind = bool(getattr(rec, "bind_to_process", False))
        prev_ctx = self._bound_context_for_record(rec=rec, notify_missing=False)
        prev_display_exe = str(prev_ctx.get("display_exe", "") or "").strip()
        rs = rec.repeat
        rs.enabled = self.cb_repeat.isChecked()
        rs.count = int(self.sp_repeat_count.value())
        rec.bind_to_process = bool(self.cb_bind_record_process.isChecked())
        if rec.bind_to_process:
            if not prev_bind:
                if not prev_display_exe:
                    rec.bind_to_process = False
                    rec.bound_exe = ""
                    rec.bound_exe_override = ""
                    self.cb_bind_record_process.blockSignals(True)
                    self.cb_bind_record_process.setChecked(False)
                    self.cb_bind_record_process.blockSignals(False)
                    self._set_status("Сначала выберите приложение, затем включите привязку записи.", level="error")
                else:
                    rec.bound_exe = prev_display_exe
                    rec.bound_exe_override = ""
                    self._remember_bound_exe_in_history(prev_display_exe)
            elif not str(getattr(rec, "bound_exe", "") or "").strip():
                fallback = prev_display_exe or str(self._bound_exe or "").strip()
                if fallback:
                    rec.bound_exe = fallback
        else:
            rec.bound_exe_override = ""
        rec.move_mouse = self.cb_move_mouse.isChecked()
        if self.cb_repeat_random.isChecked():
            rs.delay.mode = "range"
            rs.delay.a = float(self.sp_repeat_a.value())
            rs.delay.b = float(self.sp_repeat_b.value())
        else:
            rs.delay.mode = "fixed"
            rs.delay.a = float(self.sp_repeat_a.value())
            rs.delay.b = float(self.sp_repeat_a.value())

        self._set_repeat_b_row_visible(self.cb_repeat_random.isChecked())
        self._save()
        self._refresh_global_buttons()
        self._refresh_bound_process_caption(rec)

    # ---- Playback ----
    def play_current(self):
        if self.player and self.player.isRunning():
            QMessageBox.information(self, "Уже играет", "Запись уже проигрывается. Нажмите Стоп.")
            return
        rec = self._current_record()
        if not rec:
            QMessageBox.information(self, "Нет записи", "Выберите запись.")
            return
        if not rec.actions:
            self._set_status("Нечего проигрывать: в записи нет действий.", level="error")
            return
        if not self._validate_actions_before_play(rec):
            return
        ctx = self._bound_context_for_record(rec=rec, notify_missing=True)
        if str(ctx.get("mode", "")) == "record_missing" and bool(ctx.get("enabled", False)):
            return
        effective_exe = str(ctx.get("effective_exe", "") or "").strip()

        self.player = MacroPlayer(
            rec,
            bound_exe=effective_exe,
            stop_word_cfg=self._stop_word_cfg,
            stop_word_enabled_event=self._stop_word_enabled_event,
            start_index=self._get_anchor_index(rec),
        )

        self.player.signals.paused.connect(self._on_player_paused)
        self.player.signals.action_error.connect(self._on_player_action_error)
        self.player.signals.action_ok.connect(self._on_player_action_ok)

        self.player.signals.status.connect(self._set_status)
        self.player.signals.current.connect(self._set_status)
        self.player.signals.cycle.connect(self._set_status)
        self.player.signals.progress.connect(self._on_progress)
        self.player.signals.finished.connect(self._on_finished)
        self.player.signals.action_row.connect(self._on_player_action_row)
        self._set_status("Проигрывание…")
        self._disable_editing(True)

        self.hotkeys.start()

        self._playing = True

        self._error_rows.clear()
        self._clear_all_row_colors()
        self._set_pause_controls(playing=True, paused=False)

        self.player.start()

    def stop_playback(self):
        if self.player and self.player.isRunning():
            self.player.stop("Остановлено кнопкой Стоп")
            self._set_status("Остановка…")

    def _on_finished(self):
        self._playing = False
        self._disable_editing(False)

        state = getattr(self.player, "end_state", "done") if self.player else "done"
        reason = (getattr(self.player, "end_reason", "") or "").strip() if self.player else ""

        if state != "done":
            self._set_status(reason or "Проигрывание остановлено.", level="error")
        else:
            self._set_status("Готово", level="info")

        self._highlight_action_row(-1)
        self._error_rows.clear()
        self._clear_all_row_colors()
        self._set_pause_controls(playing=False, paused=False)

    def _on_progress(self, done: int, total: int):
        self._set_status(f"Прогресс: {done}/{total}")

    def _disable_editing(self, playing: bool):
        en = not playing

        self.btn_add_record.setEnabled(en)
        self.btn_del_record.setEnabled(en)
        self.btn_ren_record.setEnabled(en)
        self.btn_rec_up.setEnabled(en)
        self.btn_rec_down.setEnabled(en)

        self.actions_table.setEnabled(en)
        self.btn_add_area.setEnabled(en)
        self.btn_add_key.setEnabled(en)
        self.btn_edit_action.setEnabled(en)
        self.btn_del_action.setEnabled(en)
        self.btn_act_up.setEnabled(en)
        self.btn_act_down.setEnabled(en)
        self.btn_add_wait.setEnabled(en)
        if hasattr(self, "wait_mode_slider"):
            self.wait_mode_slider.setEnabled(en)
        if hasattr(self, "area_mode_slider"):
            self.area_mode_slider.setEnabled(en)
        if hasattr(self, "key_long_actions_table"):
            self.key_long_actions_table.setEnabled(en)
        if hasattr(self, "btn_key_long_actions_add"):
            self.btn_key_long_actions_add.setEnabled(en)
        if hasattr(self, "btn_key_long_actions_del"):
            self.btn_key_long_actions_del.setEnabled(en)
        if hasattr(self, "btn_key_long_bottom_add_key"):
            self.btn_key_long_bottom_add_key.setEnabled(en)
        if hasattr(self, "btn_key_long_bottom_edit"):
            self.btn_key_long_bottom_edit.setEnabled(en)
        if hasattr(self, "btn_key_long_bottom_delete"):
            self.btn_key_long_bottom_delete.setEnabled(en)
        if hasattr(self, "cb_key_long_hold_range"):
            self.cb_key_long_hold_range.setEnabled(en)
        if hasattr(self, "sp_key_long_hold_a"):
            self.sp_key_long_hold_a.setEnabled(en)
        if hasattr(self, "sp_key_long_hold_b"):
            self.sp_key_long_hold_b.setEnabled(en)
        if hasattr(self, "key_long_activation_slider"):
            self.key_long_activation_slider.setEnabled(en)
        if hasattr(self, "cb_key_long_start_range"):
            self.cb_key_long_start_range.setEnabled(en)
        if hasattr(self, "sp_key_long_start_a"):
            self.sp_key_long_start_a.setEnabled(en)
        if hasattr(self, "sp_key_long_start_b"):
            self.sp_key_long_start_b.setEnabled(en)
        self.btn_record_menu.setEnabled(en)
        self.btn_current_record.setEnabled(en and (self.current_index >= 0))
        self.btn_measure_window.setEnabled(en)


        # Во время проигрывания — всё гасим
        if playing:
            if self.meter.is_running():
                self.meter.stop()
            self.key_params.setEnabled(False)
            self.area_params.setEnabled(False)
            self.wait_event_params.setEnabled(False)
            self.repeat_box.setEnabled(False)
            self.measure_box.setEnabled(False)
            self.btn_repeat_settings.setEnabled(False)
            return

        # После проигрывания — восстанавливаем “как положено”
        self.measure_box.setEnabled(True)
        self.btn_measure_window.setEnabled(True)
        self._refresh_repeat_ui()
        self._on_action_selected()

    def _set_status(self, s: str, level: str = "info"):
        self.lbl_status.setText(s)
        self.lbl_status.setProperty("level", level)

        # принудительно обновляем стиль (Qt не всегда сам подхватывает property)
        self.lbl_status.style().unpolish(self.lbl_status)
        self.lbl_status.style().polish(self.lbl_status)

    def _clear_status_text(self):
        if hasattr(self, "lbl_status"):
            self.lbl_status.setText("")
            self.lbl_status.setProperty("level", "info")
            self.lbl_status.style().unpolish(self.lbl_status)
            self.lbl_status.style().polish(self.lbl_status)

    def _save_settings(self):
        try:
            self._normalize_bound_app_lists()
            payload = {
                "bound_exe": self._bound_exe,
                "bound_exe_enabled": bool(self._bound_exe_enabled),
                "bound_exe_recent": list(self._bound_exe_recent or []),
                "bound_exe_favorites": list(self._bound_exe_favorites or []),
                "stop_word_cfg": self._stop_word_cfg,
                "stop_word_enabled": bool(self._stop_word_enabled_event.is_set()),
                "last_record_index": int(self.current_index) if self.current_index >= 0 else -1,
                "language": i18n.normalize_language(getattr(self, "_ui_language", i18n.DEFAULT_LANGUAGE)),
            }
            config.SETTINGS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            self._settings_payload_cache = dict(payload)
        except Exception:
            pass

    def _load_settings(self):
        payload = self._read_settings_payload()

        self._bound_exe = str(payload.get("bound_exe", "") or "").strip()
        self._bound_exe_enabled = bool(payload.get("bound_exe_enabled", True))
        raw_recent = payload.get("bound_exe_recent", []) or []
        raw_favorites = payload.get("bound_exe_favorites", []) or []
        if not isinstance(raw_recent, list):
            raw_recent = []
        if not isinstance(raw_favorites, list):
            raw_favorites = []
        self._bound_exe_recent = [str(x).strip() for x in raw_recent if str(x).strip()]
        self._bound_exe_favorites = [str(x).strip() for x in raw_favorites if str(x).strip()]
        self._normalize_bound_app_lists()
        self._stop_word_cfg = payload.get("stop_word_cfg", None)

        if bool(payload.get("stop_word_enabled", False)):
            self._stop_word_enabled_event.set()
        else:
            self._stop_word_enabled_event.clear()

        self._ui_language = i18n.normalize_language(payload.get("language"))
        i18n.set_language(self._ui_language)
        if hasattr(self, "app_lang_group"):
            self._set_single_choice_selector(self.app_lang_group, self._ui_language, default_code=i18n.DEFAULT_LANGUAGE)

        try:
            self._last_record_index = int(payload.get("last_record_index", -1))
        except Exception:
            self._last_record_index = -1

    def _refresh_global_buttons(self):
        # --- приложение ---
        ctx = self._bound_context_for_record()
        prefix = str(ctx.get("prefix", "Приложение"))
        mode = str(ctx.get("mode", ""))
        display_exe = str(ctx.get("display_exe", "") or "").strip()

        if not display_exe:
            if mode == "record_missing":
                self.btn_bind_base.setText(f"🎯 {prefix}: не задано")
            else:
                self.btn_bind_base.setText(f"🎯 {prefix}: не выбрано")
            self.btn_bind_base.setProperty("state", "missing")
        elif not self._bound_exe_enabled:
            self.btn_bind_base.setText(f"🎯 {prefix}: выкл ({os.path.basename(display_exe)})")
            self.btn_bind_base.setProperty("state", "missing")
        else:
            r = resolve_bound_base_rect_dip(display_exe)
            if r:
                self.btn_bind_base.setText(f"🎯 {prefix}: {os.path.basename(display_exe)}")
                self.btn_bind_base.setProperty("state", "ok")
            else:
                self.btn_bind_base.setText(f"🎯 {prefix}: нет процесса ({os.path.basename(display_exe)})")
                self.btn_bind_base.setProperty("state", "missing")

        self._refresh_bind_app_menu()
        self.btn_bind_base.style().unpolish(self.btn_bind_base)
        self.btn_bind_base.style().polish(self.btn_bind_base)

        # --- стоп-слово ---
        if not self._stop_word_cfg:
            self.btn_stop_word.setText("🛑 Стоп-слово: не задано")
            self.btn_stop_word.setProperty("state", "missing")
        else:
            en = self._stop_word_enabled_event.is_set()
            w = str(self._stop_word_cfg.get("word", "") or "")
            self.btn_stop_word.setText(f"🛑 Стоп-слово: {'ВКЛ' if en else 'выкл'} ({w})")
            self.btn_stop_word.setProperty("state", "ok" if en else "missing")

        self.btn_stop_word.style().unpolish(self.btn_stop_word)
        self.btn_stop_word.style().polish(self.btn_stop_word)
        self._lock_top_buttons_geometry()

    # ---- Persistence ----
    def _save(self):
        try:
            payload = {"records": [r.to_dict() for r in self.records]}
            config.DATA_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as ex:
            QMessageBox.warning(self, "Ошибка сохранения", str(ex))

    def _load(self):
        self.stop_playback()
        self.records = []
        if config.DATA_PATH.exists():
            try:
                payload = json.loads(config.DATA_PATH.read_text(encoding="utf-8"))
                for rd in payload.get("records", []) or []:
                    self.records.append(Record.from_dict(rd))
            except Exception as ex:
                QMessageBox.warning(self, "Ошибка загрузки", f"Не удалось загрузить файл:\n{ex}")

        if not self.records:
            # start with one empty record for convenience
            self.records = [Record(name="Пример", actions=[],
                                   repeat=RepeatSettings(enabled=False, count=0, delay=Delay("fixed", 0.5, 0.5)))]

        for r in self.records:
            self._ensure_base_and_migrate(r)

        self._refresh_record_list_widget()
        self._set_current_record(0)

    def closeEvent(self, e):
        self.stop_playback()
        self._save()
        self._save_settings()

        self.hotkeys.stop()

        try:
            if hasattr(self, "meter"):
                self.meter.stop()
        except Exception:
            pass

        super().closeEvent(e)

    def _current_record(self) -> Optional[Record]:
        if 0 <= self.current_index < len(self.records):
            return self.records[self.current_index]
        return None
