# Module: geometry helpers for screen/rect conversions.
# Main: virtual_geometry, rect_to_rel, rel_to_rect.
# Example: from atari.core.geometry import rect_to_rel

from PySide6.QtCore import QPoint, QRect
from PySide6.QtGui import QGuiApplication


def virtual_geometry() -> QRect:
    screens = QGuiApplication.screens()
    vg = QRect()
    first = True
    for s in screens:
        if first:
            vg = s.geometry()
            first = False
        else:
            vg = vg.united(s.geometry())
    return vg


def _clamp01(v: float) -> float:
    try:
        v = float(v)
    except Exception:
        return 0.0
    return 0.0 if v < 0.0 else (1.0 if v > 1.0 else v)


def rect_to_rel(base: QRect, r: QRect) -> tuple[float, float, float, float]:
    base = QRect(base).normalized()
    r = QRect(r).normalized()
    denom_x = max(1, base.width() - 1)
    denom_y = max(1, base.height() - 1)

    rx1 = (r.left() - base.left()) / denom_x
    ry1 = (r.top() - base.top()) / denom_y
    rx2 = (r.right() - base.left()) / denom_x
    ry2 = (r.bottom() - base.top()) / denom_y

    return (_clamp01(rx1), _clamp01(ry1), _clamp01(rx2), _clamp01(ry2))


def rel_to_rect(base: QRect, rx1: float, ry1: float, rx2: float, ry2: float) -> QRect:
    base = QRect(base).normalized()
    denom_x = max(1, base.width() - 1)
    denom_y = max(1, base.height() - 1)

    rx1, ry1, rx2, ry2 = _clamp01(rx1), _clamp01(ry1), _clamp01(rx2), _clamp01(ry2)

    x1 = base.left() + int(round(rx1 * denom_x))
    y1 = base.top() + int(round(ry1 * denom_y))
    x2 = base.left() + int(round(rx2 * denom_x))
    y2 = base.top() + int(round(ry2 * denom_y))

    return QRect(QPoint(x1, y1), QPoint(x2, y2)).normalized()
